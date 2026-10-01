from fastapi import APIRouter, HTTPException

from ..ai import adjudicate_comment, chat_about_comment, generate_review_comments, AdjudicationError
from ..pptx import (
    build_document_model,
    build_chunks,
    context_markdown_for_comment,
    locate_quote_in_slide,
    render_markdown,
    writer,
)
from ..models import AdjudicateRequest, AiReviewRequest, ApplyAdjudicationRequest, ChatMessageRequest
from ..storage import store
from .common import document_payload, load_doc_or_404

router = APIRouter(prefix="/api/documents/{doc_id}/comments", tags=["ai"])
doc_router = APIRouter(prefix="/api/documents/{doc_id}", tags=["ai"])

# Above this size (chars, roughly 4 chars/token) we fall back to per-comment
# section context rather than sending the whole document on every request --
# keeps cost/latency sane for very large documents while giving small-to-
# medium documents (this app's expected scale) truly whole-document context.
MAX_FULL_DOC_CONTEXT_CHARS = 60_000


def _context_for_comment(doc: dict, comment_id: str, full_doc_markdown: str | None) -> str | None:
    if full_doc_markdown is not None:
        return full_doc_markdown
    return context_markdown_for_comment(doc, comment_id)


def _full_doc_markdown_if_small_enough(doc: dict) -> str | None:
    md = render_markdown(doc)
    return md if len(md) <= MAX_FULL_DOC_CONTEXT_CHARS else None


def _apply_adjudication(
    pkg, doc_id: str, comment_id: str, *, action: str, reply: str, replacement_text: str | None,
    original_text: str | None, author: str,
) -> dict:
    # PowerPoint has no tracked changes; an edit is applied in place and
    # leaves an automatic resolved "Edited text" comment as its record.
    # `change_id` is that record comment's id.
    change_id = None
    if action == "edit" and replacement_text:
        try:
            change_id = writer.replace_commented_span(
                pkg, comment_id, replacement_text, author, original_text=original_text
            )
        except writer.EditError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    reply_id = None
    if reply:
        try:
            reply_id = writer.add_comment_reply(pkg, comment_id, author, reply)
        except writer.EditError as exc:
            # Never silently drop a reply: resolving the comment anyway would
            # look like success while the reply text vanished.
            raise HTTPException(status_code=400, detail=f"could not add reply: {exc}") from exc

    writer.resolve_comment(pkg, comment_id, done=(action != "no_change"))

    store.save_package(doc_id, pkg)
    store.log_edit(
        doc_id,
        "apply_adjudication",
        author,
        {"comment_id": comment_id, "action": action, "change_id": change_id, "reply_id": reply_id},
    )
    return {"change_id": change_id, "reply_id": reply_id}


@router.post("/{comment_id}/adjudicate")
def adjudicate(doc_id: str, comment_id: str, body: AdjudicateRequest):
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)

    comment = doc["comments"].get(comment_id)
    if comment is None:
        raise HTTPException(status_code=404, detail="comment not found")

    full_doc_markdown = _full_doc_markdown_if_small_enough(doc)
    context_md = _context_for_comment(doc, comment_id, full_doc_markdown)
    if context_md is None:
        raise HTTPException(
            status_code=400, detail="comment has no slide context in the document"
        )

    meta = store.get_document_meta(doc_id)

    try:
        result = adjudicate_comment(
            comment_id=comment_id,
            comment_author=comment["author"],
            comment_text=comment["text"],
            comment_quote=comment.get("quote"),
            context_markdown=context_md,
            doc_title=meta["title"] if meta else None,
            is_full_document=full_doc_markdown is not None,
        )
    except AdjudicationError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    applied = _apply_adjudication(
        pkg, doc_id, comment_id,
        action=result.action, reply=result.reply, replacement_text=result.replacement_text,
        original_text=result.original_text, author=body.author,
    )
    store.log_ai_adjudication(
        doc_id, comment_id, result.action, result.reasoning, applied["reply_id"], applied["change_id"]
    )

    payload = document_payload(doc_id)
    payload["adjudication"] = {
        "action": result.action,
        "reply": result.reply,
        "reasoning": result.reasoning,
        "replacement_text": result.replacement_text,
        "original_text": result.original_text,
        "change_id": applied["change_id"],
        "reply_comment_id": applied["reply_id"],
    }
    return payload


@router.post("/{comment_id}/apply-adjudication")
def apply_adjudication(doc_id: str, comment_id: str, body: ApplyAdjudicationRequest):
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)
    if comment_id not in doc["comments"]:
        raise HTTPException(status_code=404, detail="comment not found")

    applied = _apply_adjudication(
        pkg, doc_id, comment_id,
        action=body.action, reply=body.reply, replacement_text=body.replacement_text,
        original_text=body.original_text, author=body.author,
    )
    store.log_ai_adjudication(
        doc_id, comment_id, body.action, body.reasoning, applied["reply_id"], applied["change_id"]
    )

    payload = document_payload(doc_id)
    payload["adjudication"] = {
        "action": body.action,
        "reply": body.reply,
        "reasoning": body.reasoning,
        "replacement_text": body.replacement_text,
        "original_text": body.original_text,
        "change_id": applied["change_id"],
        "reply_comment_id": applied["reply_id"],
    }
    return payload


@doc_router.post("/suggest-adjudications")
def suggest_adjudications(doc_id: str):
    """Read-only bulk AI pass over every open (unresolved, top-level)
    comment, with full-document context where the document is small enough.
    Never mutates the document -- pair with POST .../apply-adjudication to
    act on a suggestion."""
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)
    meta = store.get_document_meta(doc_id)
    full_doc_markdown = _full_doc_markdown_if_small_enough(doc)

    suggestions: dict[str, dict] = {}
    for comment_id, comment in doc["comments"].items():
        if comment.get("parentId") or comment.get("done") or comment.get("kind") == "edit":
            continue

        context_md = _context_for_comment(doc, comment_id, full_doc_markdown)
        if context_md is None:
            suggestions[comment_id] = {"error": "comment has no slide context in the document"}
            continue

        try:
            result = adjudicate_comment(
                comment_id=comment_id,
                comment_author=comment["author"],
                comment_text=comment["text"],
                comment_quote=comment.get("quote"),
                context_markdown=context_md,
                doc_title=meta["title"] if meta else None,
                is_full_document=full_doc_markdown is not None,
            )
        except AdjudicationError as exc:
            suggestions[comment_id] = {"error": str(exc)}
            continue

        suggestions[comment_id] = {
            "action": result.action,
            "reply": result.reply,
            "reasoning": result.reasoning,
            "replacement_text": result.replacement_text,
            "original_text": result.original_text,
        }

    return {"suggestions": suggestions, "usedFullDocumentContext": full_doc_markdown is not None}


@doc_router.post("/ai-review")
def ai_review(doc_id: str, body: AiReviewRequest = AiReviewRequest()):
    """AI-as-reviewer: reads the deck slide by slide and adds new comments
    (author 'AI Reviewer' by default) attached to specific text, the same
    mechanism the frontend's text-selection popup uses. Re-parses after
    every successful anchor so a later suggestion on the same slide sees the
    current run layout, not a stale one."""
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)
    meta = store.get_document_meta(doc_id)
    chunks = build_chunks(doc)

    added: list[str] = []
    skipped = 0

    for chunk in chunks:
        if not chunk.markdown.strip():
            continue
        try:
            suggestions = generate_review_comments(
                section_markdown=chunk.markdown, doc_title=meta["title"] if meta else None
            )
        except AdjudicationError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        for suggestion in suggestions:
            current_doc = build_document_model(pkg)
            slide = next(s for s in current_doc["slides"] if s["id"] == chunk.slide_id)
            anchor = locate_quote_in_slide(slide, chunk.item_ids, suggestion["quote"])
            if anchor is None:
                skipped += 1
                continue
            try:
                comment_id = writer.add_comment_for_selection(
                    pkg, *anchor, body.author, suggestion["comment"],
                )
                added.append(comment_id)
            except writer.EditError:
                skipped += 1

    store.save_package(doc_id, pkg)
    store.log_edit(doc_id, "ai_review", body.author, {"added": added, "skipped": skipped})

    payload = document_payload(doc_id)
    payload["aiReview"] = {"added": len(added), "skipped": skipped, "commentIds": added}
    return payload


@router.post("/{comment_id}/chat")
def chat_about_comment_route(doc_id: str, comment_id: str, body: ChatMessageRequest):
    """One turn of a persisted, per-comment discussion with the AI assistant
    -- separate from /adjudicate's forced-tool-use decision call, this is
    free-form back-and-forth meant to develop rationale (or talk through a
    'no confident suggestion' case) before the human records a decision.
    The transcript is stored in comment_chats.json, never mixed into the
    decision record itself."""
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)

    comment = doc["comments"].get(comment_id)
    if comment is None:
        raise HTTPException(status_code=404, detail="comment not found")

    full_doc_markdown = _full_doc_markdown_if_small_enough(doc)
    context_md = _context_for_comment(doc, comment_id, full_doc_markdown)
    if context_md is None:
        raise HTTPException(
            status_code=400, detail="comment has no slide context in the document"
        )

    meta = store.get_document_meta(doc_id)
    history = store.get_chats(doc_id).get(comment_id, [])
    store.append_chat_turn(doc_id, comment_id, "user", body.message, body.author)

    try:
        reply = chat_about_comment(
            comment_id=comment_id,
            comment_author=comment["author"],
            comment_text=comment["text"],
            comment_quote=comment.get("quote"),
            context_markdown=context_md,
            doc_title=meta["title"] if meta else None,
            is_full_document=full_doc_markdown is not None,
            history=history,
            user_message=body.message,
        )
    except AdjudicationError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    turns = store.append_chat_turn(doc_id, comment_id, "assistant", reply)
    payload = document_payload(doc_id)
    payload["chat"] = {"comment_id": comment_id, "history": turns}
    return payload
