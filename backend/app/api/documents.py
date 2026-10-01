import mimetypes
import zipfile

from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, Response

from ..pptx.ooxml import COMMENT_AUTHORS_PART, MODERN_AUTHORS_PART, PRESENTATION_PART, qn
from ..pptx import build_document_model
from ..models import SetDocumentMetaRequest
from ..storage import store
from .common import document_payload, load_doc_or_404

router = APIRouter(prefix="/api/documents", tags=["documents"])


@router.post("")
async def upload_document(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".pptx",)):
        raise HTTPException(status_code=400, detail="only .pptx files are supported")
    data = await file.read()
    try:
        doc_id = store.create_document(file.filename, data)
        load_doc_or_404(doc_id)  # fail fast if the file doesn't parse
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"could not parse .pptx: {exc}") from exc
    return document_payload(doc_id)


@router.get("")
def list_documents():
    return {"documents": store.list_documents()}


@router.get("/{doc_id}")
def get_document(doc_id: str):
    return document_payload(doc_id)


@router.patch("/{doc_id}/meta")
def update_document_meta(doc_id: str, body: SetDocumentMetaRequest):
    """Set the document reference/revision shown on every row of the
    Comment Resolution Matrix export (e.g. "UTIC-MK710-A012-0003" / "X1")."""
    if not store.document_exists(doc_id):
        raise HTTPException(status_code=404, detail="document not found")
    store.update_document_meta(doc_id, doc_ref=body.doc_ref, rev=body.rev)
    return document_payload(doc_id)


@router.get("/{doc_id}/markdown")
def get_document_markdown(doc_id: str):
    payload = document_payload(doc_id)
    from ..pptx import render_markdown

    return {"markdown": render_markdown(payload["document"])}


@router.get("/{doc_id}/history")
def get_document_history(doc_id: str):
    load_doc_or_404(doc_id)
    return {
        "edits": store.list_edits(doc_id),
        "ai_adjudications": store.list_ai_adjudications(doc_id),
    }


@router.get("/{doc_id}/export")
def export_document(doc_id: str):
    if not store.document_exists(doc_id):
        raise HTTPException(status_code=404, detail="document not found")
    meta = store.get_document_meta(doc_id)
    filename = meta["filename"] if meta else "presentation.pptx"
    return FileResponse(
        store.working_path(doc_id),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


@router.get("/{doc_id}/media/{slide_id}/{rel_id}")
def get_media(doc_id: str, slide_id: str, rel_id: str):
    """An image used by one slide (relationship ids are per-slide)."""
    pkg = load_doc_or_404(doc_id)
    try:
        slide_idx = int(slide_id.lstrip("s"))
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="media not found") from exc
    resolved = pkg.resolve_media(slide_idx, rel_id)
    if resolved is None:
        raise HTTPException(status_code=404, detail="media not found")
    data, part_name = resolved
    content_type = mimetypes.guess_type(part_name)[0] or "application/octet-stream"
    return Response(content=data, media_type=content_type)


@router.post("/{doc_id}/reset")
def reset_document(doc_id: str):
    if not store.document_exists(doc_id):
        raise HTTPException(status_code=404, detail="document not found")
    store.reset_to_original(doc_id)
    store.log_edit(doc_id, "reset_to_original", "system", {})
    return document_payload(doc_id)


@router.get("/{doc_id}/debug")
def debug_document(doc_id: str):
    """Structural diagnostics only -- zip part names, relationship
    declarations, and element counts, never any slide/comment text -- safe
    to paste back verbatim even from a CUI deck. For tracking down "why does
    this real-world .pptx show 0 comments" reports without ever needing to
    see the file itself."""
    if not store.document_exists(doc_id):
        raise HTTPException(status_code=404, detail="document not found")

    with zipfile.ZipFile(store.working_path(doc_id), "r") as zf:
        part_names = zf.namelist()

    pkg = load_doc_or_404(doc_id)
    slides = []
    for i, part in enumerate(pkg.slide_parts):
        tree = pkg.slide_tree(i)
        rels = pkg.rels(part)
        slides.append(
            {
                "slide": i + 1,
                "part": part,
                "shape_count": len(list(tree.iter(qn("p:sp"), qn("p:pic"), qn("p:graphicFrame")))),
                "table_count": len(list(tree.iter(qn("a:tbl")))),
                "relationships": [{"type": r["Type"].rsplit("/", 1)[-1], "target": r["Target"]} for r in rels],
            }
        )
    doc = build_document_model(pkg)
    return {
        "zip_part_names": part_names,
        "conventional_parts_present": {
            "ppt/commentAuthors.xml (legacy comments)": COMMENT_AUTHORS_PART in part_names,
            "ppt/authors.xml (modern comments)": MODERN_AUTHORS_PART in part_names,
            "ppt/presentation.xml": PRESENTATION_PART in part_names,
        },
        "slide_count": len(pkg.slide_parts),
        "slides": slides,
        "parsed_comment_count": len(doc["comments"]),
    }


@router.delete("/{doc_id}")
def delete_document(doc_id: str):
    if not store.document_exists(doc_id):
        raise HTTPException(status_code=404, detail="document not found")
    store.delete_document(doc_id)
    return {"ok": True}
