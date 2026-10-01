from fastapi import APIRouter, HTTPException

from ..pptx import writer
from ..models import TextRangeEditRequest
from ..storage import store
from .common import document_payload, load_doc_or_404

router = APIRouter(prefix="/api/documents/{doc_id}", tags=["edits"])


@router.post("/edits/replace-range")
def replace_range(doc_id: str, body: TextRangeEditRequest):
    """Replace a user-selected span of slide text. PowerPoint has no Track
    Changes, so the text is edited in place and an automatic, resolved
    "Edited text: old -> new" comment is left on the shape as the record
    (also kept in this document's edit audit log)."""
    pkg = load_doc_or_404(doc_id)
    try:
        record_id = writer.replace_text_range(
            pkg, body.start_run_id, body.start_offset, body.end_run_id, body.end_offset,
            body.new_text, body.author,
        )
    except writer.EditError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    store.save_package(doc_id, pkg)
    store.log_edit(
        doc_id,
        "replace_text_range",
        body.author,
        {
            "start_run_id": body.start_run_id,
            "start_offset": body.start_offset,
            "end_run_id": body.end_run_id,
            "end_offset": body.end_offset,
            "new_text": body.new_text,
            "edit_record_comment_id": record_id,
        },
    )
    return document_payload(doc_id)
