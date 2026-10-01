import csv
import io

import openpyxl
from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

from .. import xlsx_roundtrip
from ..pptx import build_document_model, comment_locations, find_comment_anchor, writer
from ..models import (
    CommentReplyRequest,
    CommentResolveRequest,
    CreateCommentRequest,
    SaveDecisionRequest,
)
from ..storage import store
from .common import document_payload, load_doc_or_404

router = APIRouter(prefix="/api/documents/{doc_id}/comments", tags=["comments"])

# Display labels for the Comment Resolution Matrix export -- the internal
# decision codes stay snake_case/lowercase for API stability while the
# exported column matches the printed matrix's wording.
DECISION_LABELS = {
    "accept": "Accept",
    "reject": "Reject",
    "info_only": "Info Only",
    "defer": "Defer",
}


@router.get("")
def list_comments(doc_id: str):
    return document_payload(doc_id)["document"]["comments"]


@router.post("")
def create_comment(doc_id: str, body: CreateCommentRequest):
    """Add a top-level comment, either on a text selection (start/end run
    ids + offsets) or directly on a slide / shape (slide_id[, shape_id])."""
    pkg = load_doc_or_404(doc_id)
    has_selection = None not in (body.start_run_id, body.start_offset, body.end_run_id, body.end_offset)
    try:
        if has_selection:
            comment_id = writer.add_comment_for_selection(
                pkg, body.start_run_id, body.start_offset, body.end_run_id, body.end_offset,
                body.author, body.text,
            )
        elif body.slide_id:
            comment_id = writer.add_comment(
                pkg, slide_id=body.slide_id, shape_id=body.shape_id, author=body.author, text=body.text
            )
        else:
            raise HTTPException(
                status_code=400, detail="provide a text selection or a slide_id to comment on"
            )
    except writer.EditError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    store.save_package(doc_id, pkg)
    store.log_edit(
        doc_id,
        "add_comment",
        body.author,
        {
            "comment_id": comment_id,
            "slide_id": body.slide_id,
            "shape_id": body.shape_id,
            "start_run_id": body.start_run_id,
            "end_run_id": body.end_run_id,
        },
    )
    return document_payload(doc_id)


@router.post("/{comment_id}/reply")
def reply_to_comment(doc_id: str, comment_id: str, body: CommentReplyRequest):
    pkg = load_doc_or_404(doc_id)
    try:
        reply_id = writer.add_comment_reply(pkg, comment_id, body.author, body.text)
    except writer.EditError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    store.save_package(doc_id, pkg)
    store.log_edit(
        doc_id, "comment_reply", body.author, {"parent_comment_id": comment_id, "reply_id": reply_id}
    )
    return document_payload(doc_id)


@router.post("/{comment_id}/resolve")
def resolve_comment(doc_id: str, comment_id: str, body: CommentResolveRequest):
    pkg = load_doc_or_404(doc_id)
    try:
        writer.resolve_comment(pkg, comment_id, body.done)
    except writer.EditError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    store.save_package(doc_id, pkg)
    store.log_edit(
        doc_id, "resolve_comment", body.author, {"comment_id": comment_id, "done": body.done}
    )
    return document_payload(doc_id)


@router.post("/{comment_id}/decision")
def save_comment_decision(doc_id: str, comment_id: str, body: SaveDecisionRequest):
    """Formal Accept/Reject/Info Only/Defer adjudication for the Comment
    Resolution Matrix, with a required rationale -- independent of the
    pptx-level resolve/reopen toggle above, which only flips the
    comment thread's own resolved flag."""
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)
    if comment_id not in doc["comments"]:
        raise HTTPException(status_code=404, detail="comment not found")

    record = store.save_decision(doc_id, comment_id, body.decision, body.reason, body.ref, body.author)
    store.log_edit(
        doc_id,
        "comment_decision",
        body.author,
        {"comment_id": comment_id, "decision": body.decision, "reason": body.reason, "ref": body.ref},
    )
    payload = document_payload(doc_id)
    payload["decision"] = {"comment_id": comment_id, **record}
    return payload


@router.get("/export")
def export_comment_resolution_matrix(doc_id: str):
    """TSV export of every top-level comment plus its current adjudication
    decision, with column headers matching the government/contractor
    Comment Resolution Matrix template's own second header row (the one
    below its merged "Government Columns"/"Contractor Columns" banner)
    word-for-word: Item #, Document, Rev, Ref #, Para #, Comment,
    Contractor Response (Accept / Reject), Contractor Rationale,
    Status (Open / Closed), Date Adjudicated."""
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)
    meta = store.get_document_meta(doc_id) or {}
    decisions = store.get_decisions(doc_id)
    locations = comment_locations(doc)

    doc_ref = meta.get("doc_ref") or meta.get("filename") or ""
    rev = meta.get("rev") or ""

    buf = io.StringIO()
    tsv_writer = csv.writer(buf, delimiter="\t", lineterminator="\n")
    tsv_writer.writerow(
        [
            "Item #", "Document", "Rev", "Ref #", "Para #", "Comment",
            "Contractor Response (Accept / Reject)", "Contractor Rationale",
            "Status (Open / Closed)", "Date Adjudicated", "Comment ID",
        ]
    )

    item_no = 0
    for comment_id, comment in doc["comments"].items():
        if comment.get("parentId") or comment.get("kind") == "edit":
            continue  # replies are threaded under a row; edit records aren't review comments
        item_no += 1
        loc = locations.get(comment_id, {})
        decision = decisions.get(comment_id)
        # A human-entered Ref # (typed into the decision form) always wins
        # over the auto-derived guess -- the derivation is a starting point,
        # not an override of a reviewer's own correction.
        manual_ref = decision.get("ref") if decision else None
        ref_value = manual_ref or loc.get("derivedRef") or ""
        para_value = loc.get("derivedPara") or loc.get("paraRef", "")
        tsv_writer.writerow(
            [
                item_no,
                doc_ref,
                rev,
                ref_value,
                para_value,
                comment.get("text", ""),
                DECISION_LABELS.get(decision["decision"], decision["decision"]) if decision else "",
                decision["reason"] if decision else "",
                "Closed" if decision else "Open",
                decision["decidedAt"] if decision else "",
                # Not part of the paper template -- carried through so a
                # round trip (export, edit in Excel, re-upload to
                # /import-decisions) matches rows back to comments exactly
                # instead of by comment text alone. Safe to ignore/delete.
                comment_id,
            ]
        )

    filename = f"comment_resolution_matrix_{doc_id[:8]}.tsv"
    return Response(
        content=buf.getvalue(),
        media_type="text/tab-separated-values",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/import-decisions")
async def import_comment_resolution_matrix(
    doc_id: str, file: UploadFile = File(...), author: str = Form("Imported")
):
    """Round-trip a filled-in Comment Resolution Matrix .xlsx back into this
    document: for every row with a recognized Contractor Response, save the
    decision/rationale (same record /decision produces) AND push it into the
    actual .pptx as a real threaded reply plus a resolved comment -- so
    opening the file in PowerPoint shows the disposition too, not just this app.
    Matches rows to comments via a hidden "Comment ID" column when present
    (this app's own export, round-tripped through Excel), falling back to
    exact comment text otherwise (a reviewer's own blank template). When a
    row matches no existing comment at all -- the common real-world case of
    a clean master .pptx paired with a separately-maintained review
    spreadsheet -- a brand-new comment is created instead, anchored via
    the row's Ref #/Para # (requirement ID, caption, or numbered section)
    using the same heuristics that generate those columns on export. Rows
    with no recognized response, or with neither an existing comment nor a
    locatable Ref #/Para #, are skipped and reported rather than guessed
    at."""
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="only .xlsx files are supported")

    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)

    data = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"could not read .xlsx file: {exc}") from exc

    ws = next((wb[n] for n in wb.sheetnames if n.strip().lower() == "comments"), wb[wb.sheetnames[0]])

    header = xlsx_roundtrip.find_header_row(ws)
    if header is None:
        raise HTTPException(
            status_code=400,
            detail='could not find a header row with "Comment" and "Contractor Response" columns',
        )
    header_row, col_map = header
    rows = xlsx_roundtrip.parse_rows(ws, header_row, col_map)

    applied: list[dict] = []
    skipped: list[dict] = []
    for row in rows:
        decision = xlsx_roundtrip.decision_from_response(row["response"])
        if decision is None:
            skipped.append(
                {"row": row["row"], "reason": f"no recognized response (\"{row['response']}\")"}
            )
            continue

        comment_id = xlsx_roundtrip.match_comment_id(row, doc["comments"])
        created = False
        if comment_id is None:
            anchor = find_comment_anchor(doc, row.get("ref"), row.get("para"))
            if anchor is None:
                skipped.append(
                    {
                        "row": row["row"],
                        "reason": "no matching open comment, and no Ref #/Para # location found to create one",
                    }
                )
                continue
            try:
                comment_id = writer.add_comment(
                    pkg,
                    slide_id=anchor["slide_id"],
                    shape_id=anchor["shape_id"],
                    quote=anchor["quote"],
                    author="CRM Import",
                    text=row["comment_text"] or "(imported from Comment Resolution Matrix)",
                )
            except writer.EditError:
                skipped.append(
                    {
                        "row": row["row"],
                        "reason": "found a Ref #/Para # location but could not anchor a new comment there",
                    }
                )
                continue
            created = True

        store.save_decision(doc_id, comment_id, decision, row["rationale"], row["ref"], author)

        label = DECISION_LABELS[decision]
        reply_text = f"{label}: {row['rationale']}" if row["rationale"] else label
        try:
            writer.add_comment_reply(pkg, comment_id, author, reply_text)
        except writer.EditError:
            pass  # comment's anchor no longer exists in the body; decision is still recorded
        writer.resolve_comment(pkg, comment_id, done=True)

        applied.append(
            {"row": row["row"], "comment_id": comment_id, "decision": decision, "created": created}
        )

    store.save_package(doc_id, pkg)
    store.log_edit(
        doc_id, "import_crm_xlsx", author, {"applied": len(applied), "skipped": len(skipped)}
    )

    payload = document_payload(doc_id)
    payload["xlsxImport"] = {
        "applied": applied,
        "skipped": skipped,
        "skippedSummary": xlsx_roundtrip.summarize_skips(skipped),
    }
    return payload
