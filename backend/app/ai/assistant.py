"""Claude-powered comment adjudication.

Given a reviewer's comment and the markdown context around its anchor
(encoded with CriticMarkup so the model can see where comments are
attached), ask Claude to decide how to handle it: propose a text edit
(verbatim original text -> replacement), write a reply without editing, or
decline with a reason. We force structured output via tool-use so the
result is safe to apply programmatically (PowerPoint has no tracked
changes; the edit is applied in place and recorded as a resolved comment).
"""
from __future__ import annotations

import time
from dataclasses import dataclass

import anthropic
from anthropic import Anthropic

from .. import config

SYSTEM_PROMPT = """You are an AI editorial assistant that adjudicates reviewer comments on a \
PowerPoint presentation. PowerPoint has no Track Changes, so when you edit text the change is applied \
in place and recorded as an automatic comment authored "AI Assistant" reading \
"Edited text: old -> new" -- a human can always see and undo what you did.

Slide content is shown to you in Markdown, one "## Slide N: Title" section per slide, using \
CriticMarkup to encode comment anchors:
  {==span==}{>>comment:ID<<}   text a reviewer attached comment ID to (when the reviewer selected \
specific text) -- or, when the comment is attached to a whole text box, every paragraph of that \
box is marked this way
  {>>comment:ID (about the slide as a whole)<<}   a comment on the slide itself, not on any text
Tables appear as Markdown tables, pictures as [Picture: ...], and speaker notes as a quote block.

You are given ONE specific reviewer comment and the slide(s) around it.

Decide exactly one action:
- "edit": the comment asks for a specific, well-defined textual change to text on the slide. Set \
"original_text" to the exact, verbatim, character-for-character text currently on the slide that \
should be replaced (a word, phrase, or one sentence within a single paragraph -- NEVER spanning \
two bullets or paragraphs, and never including the {== ==} markers), and "replacement_text" to what \
it should become. If the reviewer selected specific text, "original_text" should be that text. Keep \
the replacement grammatically consistent with the surrounding text.
- "reply": the comment is a question, a request for clarification, or a discussion point that you \
can meaningfully answer in words without editing the slide text (for example a request to change \
a picture, layout, or speaker notes, which you cannot edit).
- "no_change": you cannot confidently resolve this comment (it's ambiguous, needs information you \
don't have, or is out of scope for a text edit). Explain what a human should do instead.

Always write a short "reply" (1-3 sentences) as if replying directly to the reviewer in the \
comment thread -- it will be posted verbatim as a threaded reply. Always write a brief internal \
"reasoning" for an audit log (not shown to the reviewer).
"""

ADJUDICATE_TOOL = {
    "name": "submit_adjudication",
    "description": "Submit the adjudication decision for this reviewer comment.",
    "input_schema": {
        "type": "object",
        "properties": {
            "action": {
                "type": "string",
                "enum": ["edit", "reply", "no_change"],
            },
            "original_text": {
                "type": "string",
                "description": "Required and ONLY set when action is 'edit': the exact, verbatim "
                "text currently on the slide (within one paragraph) that is being replaced, with no "
                "CriticMarkup or comment markers in it.",
            },
            "replacement_text": {
                "type": "string",
                "description": "Required and ONLY set when action is 'edit': the new text that "
                "replaces original_text, with no CriticMarkup or comment markers in it.",
            },
            "reply": {
                "type": "string",
                "description": "Message posted as a threaded reply to the reviewer's comment.",
            },
            "reasoning": {
                "type": "string",
                "description": "Brief internal rationale for the audit log.",
            },
        },
        "required": ["action", "reply", "reasoning"],
    },
}


REVIEW_SYSTEM_PROMPT = """You are an AI reviewer reading a PowerPoint presentation slide by slide, \
the same way a careful human reviewer would, leaving comments (not silent edits) on specific text \
that needs the author's attention.

Slide content is shown to you in Markdown ("## Slide N: Title"), using CriticMarkup:
  {==span==}{>>comment:ID<<}   text an earlier reviewer already commented on

Read the slide and decide whether it needs any NEW review comments: unclear or ambiguous wording, \
a claim that needs a citation or verification, a potential factual or logical inconsistency with \
the rest of the deck, a missing detail, or an open question a reviewer would reasonably raise. \
Don't invent nitpicks or comment on something that already has a comment attached -- if the slide \
is fine as written, return zero comments. Return at most 3 comments for this slide; quality over \
quantity.

For each comment, "quote" MUST be an exact, verbatim, character-for-character substring of ONE \
paragraph (one bullet, one table cell, or the title) on the slide -- no CriticMarkup syntax, no \
ellipses, no paraphrasing, never spanning two bullets. Keep it short enough to be unique on the \
slide (a few words to one sentence): it will be used to programmatically locate the exact text, so \
it must match exactly or the comment will be silently dropped.
"""

REVIEW_TOOL = {
    "name": "submit_review_comments",
    "description": "Submit any new review comments found on this slide (zero is fine).",
    "input_schema": {
        "type": "object",
        "properties": {
            "comments": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "quote": {
                            "type": "string",
                            "description": "Exact verbatim substring of one paragraph on the slide "
                            "to attach the comment to.",
                        },
                        "comment": {
                            "type": "string",
                            "description": "The review comment text.",
                        },
                    },
                    "required": ["quote", "comment"],
                },
            },
        },
        "required": ["comments"],
    },
}


@dataclass
class AdjudicationResult:
    action: str  # "edit" | "reply" | "no_change"
    replacement_text: str | None
    original_text: str | None
    reply: str
    reasoning: str


class AdjudicationError(RuntimeError):
    pass


def _client() -> Anthropic:
    if not config.ANTHROPIC_API_KEY:
        raise AdjudicationError(
            "ANTHROPIC_API_KEY is not set in the backend environment; AI adjudication is unavailable."
        )
    key = config.ANTHROPIC_API_KEY
    headers = dict(config.EXTRA_HEADERS)
    kwargs: dict = {"base_url": config.ANTHROPIC_BASE_URL or None}
    if config.AUTH_MODE == "bearer":
        # Authorization: Bearer <token> and NO x-api-key header. The SDK would
        # otherwise fall back to the ANTHROPIC_API_KEY environment variable
        # when api_key is None (and prefers x-api-key when both exist), so
        # clear it explicitly after construction.
        if headers:
            kwargs["default_headers"] = headers
        client = Anthropic(api_key=None, auth_token=key, **kwargs)
        client.api_key = None
        return client
    if config.AUTH_MODE == "both":
        # The SDK only ever sends one of its two credentials, so send the
        # Bearer header ourselves alongside x-api-key.
        headers["Authorization"] = f"Bearer {key}"
    if headers:
        kwargs["default_headers"] = headers
    return Anthropic(api_key=key, **kwargs)


def describe_api_error(exc: Exception) -> str:
    """A diagnostic message for a failed AI call: what was requested, what
    came back, and the usual cause -- so a misconfigured gateway can be fixed
    from the message alone, without reading server logs."""
    prefix = "Claude API error"
    if isinstance(exc, anthropic.APIStatusError):
        req = getattr(exc, "request", None)
        where = f"{req.method} {req.url}" if req is not None else "request"
        body = ""
        try:
            body = (exc.response.text or "").strip().replace("\n", " ")[:300]
        except Exception:  # noqa: BLE001
            pass
        hint = {
            401: "the gateway rejected the credential -- try the other Authentication mode "
            "(Bearer token vs. x-api-key) or check the token",
            403: "authenticated but not permitted -- check the token's access, the model, "
            "and any required extra headers",
            404: "the URL or model was not found -- check the base URL (host only, no /v1; some "
            "gateways need a path prefix) and the model id",
            429: "rate limited or out of quota",
        }.get(exc.status_code, "")
        out = f"{prefix}: {where} returned HTTP {exc.status_code}"
        if body:
            out += f" -- {body}"
        if hint:
            out += f". Likely cause: {hint}."
        return out
    if isinstance(exc, anthropic.APIConnectionError):
        req = getattr(exc, "request", None)
        where = f" to {req.url}" if req is not None else ""
        cause = exc.__cause__
        detail = f" ({type(cause).__name__}: {cause})" if cause else ""
        return (
            f"{prefix}: could not connect{where}{detail}. Check the base URL, that this server can "
            "reach it (network/proxy/firewall), and TLS certificates for internal hosts."
        )
    return f"{prefix}: {exc}"


def test_connection() -> dict:
    """Make the smallest real request the Messages API allows (max_tokens=1,
    a one-word prompt) using whatever key/base URL/model are currently
    active, so "it connects" means the exact same client construction and
    endpoint every other AI feature in this app actually uses -- not just a
    key-format check. Raises AdjudicationError on any failure (missing key,
    auth, network, wrong base URL/model), same as every other AI call."""
    client = _client()
    started = time.monotonic()
    try:
        resp = client.messages.create(
            model=config.AI_MODEL,
            max_tokens=1,
            messages=[{"role": "user", "content": "Hi"}],
        )
    except anthropic.APIError as exc:
        raise AdjudicationError(describe_api_error(exc)) from exc
    latency_ms = round((time.monotonic() - started) * 1000)
    return {
        "model": config.AI_MODEL,
        "baseUrl": config.ANTHROPIC_BASE_URL or None,
        "latencyMs": latency_ms,
        "responseId": resp.id,
    }


def _anchor_note(comment_id: str, comment_quote: str | None) -> str:
    if comment_quote:
        return (
            f'The reviewer selected this exact text when commenting: "{comment_quote}" -- it is the span '
            f"marked with {{>>comment:{comment_id}<<}} in the context above."
        )
    return (
        f"This comment is marked with {{>>comment:{comment_id}<<}} in the context above: it is attached "
        "to the marked text box (or to the slide itself if no text is marked), not to specific selected words."
    )


def adjudicate_comment(
    *,
    comment_id: str,
    comment_author: str,
    comment_text: str,
    context_markdown: str,
    comment_quote: str | None = None,
    doc_title: str | None = None,
    is_full_document: bool = False,
) -> AdjudicationResult:
    client = _client()

    context_label = "Full presentation context" if is_full_document else "Slide context"
    user_prompt = f"""Presentation: {doc_title or "(untitled)"}

{context_label} (Markdown with CriticMarkup):
---
{context_markdown}
---

Reviewer comment to adjudicate (id={comment_id}, author={comment_author}):
"{comment_text}"
{_anchor_note(comment_id, comment_quote)}
Call submit_adjudication with your decision."""

    try:
        resp = client.messages.create(
            model=config.AI_MODEL,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            tools=[ADJUDICATE_TOOL],
            tool_choice={"type": "tool", "name": "submit_adjudication"},
            messages=[{"role": "user", "content": user_prompt}],
        )
    except anthropic.APIError as exc:
        raise AdjudicationError(describe_api_error(exc)) from exc

    for block in resp.content:
        if block.type == "tool_use" and block.name == "submit_adjudication":
            data = block.input
            action = data.get("action", "no_change")
            return AdjudicationResult(
                action=action,
                replacement_text=data.get("replacement_text") if action == "edit" else None,
                original_text=data.get("original_text") if action == "edit" else None,
                reply=data.get("reply", ""),
                reasoning=data.get("reasoning", ""),
            )

    raise AdjudicationError("model did not return a submit_adjudication tool call")


CHAT_SYSTEM_PROMPT = """You are an AI editorial assistant having a focused conversation with a human \
reviewer about ONE specific reviewer comment on a PowerPoint presentation, to help them reach an adjudication \
decision (Accept / Reject / Info Only / Defer) with a solid written rationale.

Slide content is shown to you in Markdown ("## Slide N: Title" per slide), using CriticMarkup to \
encode comment anchors:
  {==span==}{>>comment:ID<<}   text a reviewer attached comment ID to (a whole text box is marked \
paragraph by paragraph when no specific text was selected)
  {>>comment:ID (about the slide as a whole)<<}   a comment on the slide itself

The text the comment under discussion is attached to is marked with its ID in the context. Have a normal, focused conversation: answer questions, ask clarifying questions of your \
own when the comment is ambiguous, propose specific rationale wording, and suggest (and revise) a \
decision as the human gives you more context. You are not asked to output structured data here -- \
just talk. Keep replies concise (a few sentences to a short paragraph) unless asked for more detail.

Write every reply in plain prose only -- no Markdown formatting of any kind (no **bold**, *italic*, \
`code`, # headings, bullet or numbered lists, or > blockquotes). The reviewer copies your replies \
directly into a document or spreadsheet to use as rationale text, so any Markdown syntax would show up as literal stray \
characters there instead of formatting; write plain sentences and paragraphs the way you would in an \
email, using a new line or "1)", "2)" if you need to enumerate points."""


def chat_about_comment(
    *,
    comment_id: str,
    comment_author: str,
    comment_text: str,
    context_markdown: str,
    doc_title: str | None,
    is_full_document: bool,
    history: list[dict[str, str]],
    user_message: str,
    comment_quote: str | None = None,
) -> str:
    """One turn of a persisted, per-comment discussion with the reviewer.
    The Messages API is stateless, so every call resends the document/
    comment context as a synthetic opening exchange, then replays the
    stored history, then the new message -- `history` holds only the real
    back-and-forth (what gets persisted), never this synthetic framing."""
    client = _client()

    context_label = "Full presentation context" if is_full_document else "Slide context"
    intro = f"""Presentation: {doc_title or "(untitled)"}

{context_label} (Markdown with CriticMarkup):
---
{context_markdown}
---

Reviewer comment under discussion (id={comment_id}, author={comment_author}):
"{comment_text}"
{_anchor_note(comment_id, comment_quote)}"""

    messages: list[dict[str, str]] = [{"role": "user", "content": intro}]
    if history:
        messages.append(
            {"role": "assistant", "content": "Understood -- I have the document and comment context."}
        )
        messages.extend({"role": t["role"], "content": t["content"]} for t in history)
    messages.append({"role": "user", "content": user_message})

    try:
        resp = client.messages.create(
            model=config.AI_MODEL,
            max_tokens=1024,
            system=CHAT_SYSTEM_PROMPT,
            messages=messages,
        )
    except anthropic.APIError as exc:
        raise AdjudicationError(describe_api_error(exc)) from exc

    reply = "".join(block.text for block in resp.content if block.type == "text")
    if not reply:
        raise AdjudicationError("model did not return a text reply")
    return reply


def generate_review_comments(
    *, section_markdown: str, doc_title: str | None = None
) -> list[dict[str, str]]:
    """AI-as-reviewer: read one slide and propose new review comments,
    each anchored to a verbatim quote from the slide. Returns a list of
    {"quote": ..., "comment": ...} -- callers are responsible for resolving
    each quote to an actual run anchor and may silently drop ones that
    don't match (the model can hallucinate a near-quote)."""
    client = _client()

    user_prompt = f"""Presentation: {doc_title or "(untitled)"}

Slide (Markdown with CriticMarkup):
---
{section_markdown}
---

Call submit_review_comments with any new review comments for this slide."""

    try:
        resp = client.messages.create(
            model=config.AI_MODEL,
            max_tokens=1536,
            system=REVIEW_SYSTEM_PROMPT,
            tools=[REVIEW_TOOL],
            tool_choice={"type": "tool", "name": "submit_review_comments"},
            messages=[{"role": "user", "content": user_prompt}],
        )
    except anthropic.APIError as exc:
        raise AdjudicationError(describe_api_error(exc)) from exc

    for block in resp.content:
        if block.type == "tool_use" and block.name == "submit_review_comments":
            comments = block.input.get("comments", [])
            return [
                {"quote": c["quote"], "comment": c["comment"]}
                for c in comments
                if c.get("quote") and c.get("comment")
            ]

    raise AdjudicationError("model did not return a submit_review_comments tool call")
