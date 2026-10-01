from fastapi import HTTPException

from ..pptx import PptxPackage, build_document_model, build_chunks, comment_locations
from ..storage import store


def load_doc_or_404(doc_id: str) -> PptxPackage:
    if not store.document_exists(doc_id):
        raise HTTPException(status_code=404, detail="document not found")
    return store.load_package(doc_id)


def document_payload(doc_id: str) -> dict:
    pkg = load_doc_or_404(doc_id)
    doc = build_document_model(pkg)
    chunks = [c.to_dict() for c in build_chunks(doc)]
    meta = store.get_document_meta(doc_id)
    decisions = store.get_decisions(doc_id)
    chats = store.get_chats(doc_id)
    # Best-effort Ref #/Para # guesses (the comment's slide; a requirement ID
    # in the commented shape) -- surfaced separately from `document.comments`
    # so the frontend can show/prefill them without the parser needing to
    # know about the Comment Resolution Matrix at all. See
    # pptx/markdown_render.py's comment_locations().
    comment_refs = {
        cid: {"ref": loc.get("derivedRef"), "para": loc.get("derivedPara")}
        for cid, loc in comment_locations(doc).items()
    }
    return {
        "meta": meta,
        "document": doc,
        "chunks": chunks,
        "decisions": decisions,
        "chats": chats,
        "commentRefs": comment_refs,
    }
