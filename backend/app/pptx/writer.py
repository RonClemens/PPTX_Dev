"""Mutations on a live PptxPackage: replace text, and manage comments
(add, reply, resolve) -- all as direct OOXML tree surgery so everything
else in the deck round-trips untouched.

PowerPoint has **no Track Changes**. Where the Word version records an edit
as a tracked del+ins pair, this one edits the text in place (the first
affected run's formatting carries over to the new text, the same thing
PowerPoint does when you type over a selection) and leaves an *edit record*:
an automatic, already-resolved comment on the shape reading
`Edited text: "old" -> "new"`, authored by whoever made the change. That
keeps the before/after visible inside the .pptx itself (in PowerPoint's
Comments pane, as a resolved note) without any app-side database.
"""
from __future__ import annotations

from lxml import etree

from . import comments as comments_mod
from .anchor import locate_quote_in_shape
from .comments import CommentError
from .ooxml import PptxPackage, local, qn
from .parser import PIN_INSET, build_document_model, iter_para_items, parse_slide, run_text


class EditError(ValueError):
    pass


# -- id resolution ------------------------------------------------------------


def parse_slide_id(slide_id: str) -> int:
    try:
        return int(slide_id.lstrip("s"))
    except ValueError as exc:
        raise EditError(f"malformed slide id {slide_id}") from exc


def _shape_el(pkg: PptxPackage, slide_idx: int, cnv_id: str):
    if not (0 <= slide_idx < len(pkg.slide_parts)):
        raise EditError(f"slide s{slide_idx} not found")
    tree = pkg.slide_tree(slide_idx)
    for el in tree.iter(qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"), qn("p:cxnSp")):
        for nv in el:
            if local(nv).startswith("nv"):
                cnv = nv.find(qn("p:cNvPr"))
                if cnv is not None and cnv.get("id") == cnv_id:
                    return el
                break
    raise EditError(f"shape {cnv_id} not found on slide s{slide_idx}")


def _split_shape_id(shape_id: str) -> tuple[int, str]:
    head, _, rest = shape_id.partition(".sh")
    if not rest:
        raise EditError(f"malformed shape id {shape_id}")
    return parse_slide_id(head), rest


def find_paragraph(pkg: PptxPackage, para_id: str):
    """Returns (slide_idx, paragraph_el, shape_el). para_id is
    's<i>.sh<id>.p<k>' or, for a table cell, 's<i>.sh<id>.r<r>.c<c>.p<k>'."""
    parts = para_id.split(".")
    if len(parts) not in (3, 5):
        raise EditError(f"malformed paragraph id {para_id}")
    slide_idx = parse_slide_id(parts[0])
    shape = _shape_el(pkg, slide_idx, parts[1][2:])
    try:
        if len(parts) == 3:
            tx = shape.find(qn("p:txBody"))
            if tx is None:
                raise EditError(f"{para_id} has no text")
            return slide_idx, tx.findall(qn("a:p"))[int(parts[2][1:])], shape
        tbl = shape.find(f".//{qn('a:tbl')}")
        tr = tbl.findall(qn("a:tr"))[int(parts[2][1:])]
        tc = tr.findall(qn("a:tc"))[int(parts[3][1:])]
        return slide_idx, tc.find(qn("a:txBody")).findall(qn("a:p"))[int(parts[4][1:])], shape
    except (IndexError, AttributeError, ValueError) as exc:
        raise EditError(f"paragraph {para_id} not found") from exc


def find_run(pkg: PptxPackage, run_id: str):
    """Returns (run_el, paragraph_el, slide_idx, shape_el, run_index)."""
    para_id, _, r_suffix = run_id.rpartition(".r")
    if not para_id or not r_suffix.isdigit():
        raise EditError(f"malformed run id {run_id}")
    slide_idx, p_el, shape = find_paragraph(pkg, para_id)
    items = list(iter_para_items(p_el))
    i = int(r_suffix)
    if i >= len(items):
        raise EditError(f"run {run_id} not found")
    return items[i], p_el, slide_idx, shape, i


def _shape_pos(pkg: PptxPackage, slide_idx: int, cnv_id: str | None) -> tuple[int, int]:
    if cnv_id is not None:
        sid = f"s{slide_idx}.sh{cnv_id}"
        for s in parse_slide(pkg, slide_idx)["shapes"]:
            if s["id"] == sid:
                return s["x"] + PIN_INSET, s["y"] + PIN_INSET
    return PIN_INSET, PIN_INSET


# -- text edits ---------------------------------------------------------------


def _set_text(r_el, text: str) -> None:
    t = r_el.find(qn("a:t"))
    if t is None:
        t = etree.SubElement(r_el, qn("a:t"))
    t.text = text


def _resolve_selection(pkg: PptxPackage, start_run_id: str, start_offset: int, end_run_id: str, end_offset: int):
    """Validate a user text selection. Returns
    (slide_idx, shape_el, p_el, run_els, start_offset, end_offset) with
    run_els the contiguous a:r elements from the start run through the end run."""
    s_el, s_p, s_slide, s_shape, si = find_run(pkg, start_run_id)
    e_el, e_p, _, _, ei = find_run(pkg, end_run_id)
    if s_p is not e_p:
        raise EditError("selection must stay within a single paragraph")
    if ei < si or (ei == si and end_offset < start_offset):
        si, ei = ei, si
        start_offset, end_offset = end_offset, start_offset
    items = list(iter_para_items(s_p))
    run_els = items[si : ei + 1]
    for el in run_els:
        if local(el) != "r":
            raise EditError(
                "selection includes a line break or field; select plain text within one paragraph"
            )
    first_len = len(run_text(run_els[0]))
    last_len = len(run_text(run_els[-1]))
    start_offset = max(0, min(start_offset, first_len))
    end_offset = max(0, min(end_offset, last_len))
    if len(run_els) == 1 and end_offset < start_offset:
        end_offset = start_offset
    return s_slide, s_shape, s_p, run_els, start_offset, end_offset


def _cnv_id(shape_el) -> str | None:
    for nv in shape_el:
        if local(nv).startswith("nv"):
            cnv = nv.find(qn("p:cNvPr"))
            return cnv.get("id") if cnv is not None else None
    return None


def _clip(text: str, n: int = 160) -> str:
    text = " ".join(text.split())
    return text if len(text) <= n else text[: n - 1] + "…"


def replace_text_range(
    pkg: PptxPackage,
    start_run_id: str,
    start_offset: int,
    end_run_id: str,
    end_offset: int,
    new_text: str,
    author: str,
    *,
    record: bool = True,
) -> str | None:
    """Replace an arbitrary user-selected span (which may cross several
    runs of one paragraph) with new_text. The first affected run keeps its
    formatting and carries the new text; later affected runs are trimmed or
    removed. Returns the id of the edit-record comment (None if record=False
    or the text didn't actually change)."""
    slide_idx, shape, p_el, run_els, so, eo = _resolve_selection(
        pkg, start_run_id, start_offset, end_run_id, end_offset
    )
    first, last = run_els[0], run_els[-1]
    first_text, last_text = run_text(first), run_text(last)
    if first is last:
        old = first_text[so:eo]
        _set_text(first, first_text[:so] + new_text + first_text[eo:])
    else:
        old = first_text[so:] + "".join(run_text(r) for r in run_els[1:-1]) + last_text[:eo]
        _set_text(first, first_text[:so] + new_text)
        for mid in run_els[1:-1]:
            p_el.remove(mid)
        remaining = last_text[eo:]
        if remaining:
            _set_text(last, remaining)
        else:
            p_el.remove(last)
    pkg.mark_dirty(pkg.slide_parts[slide_idx])

    if not record or old == new_text:
        return None
    return add_comment(
        pkg,
        slide_id=f"s{slide_idx}",
        shape_id=f"s{slide_idx}.sh{_cnv_id(shape)}",
        author=author,
        text=f"Edited text: “{_clip(old)}” → “{_clip(new_text)}”",
        kind="edit",
        done=True,
    )


def replace_text_in_shape(
    pkg: PptxPackage, shape_id: str, original: str, replacement: str, author: str
) -> str | None:
    """Replace the first verbatim occurrence of `original` in a shape."""
    slide_idx, cnv_id = _split_shape_id(shape_id)
    shape = next((s for s in parse_slide(pkg, slide_idx)["shapes"] if s["id"] == shape_id), None)
    if shape is None:
        raise EditError(f"shape {shape_id} not found")
    anchor = locate_quote_in_shape(shape, original)
    if anchor is None:
        raise EditError("the original text was not found in that shape (it may have changed)")
    return replace_text_range(pkg, *anchor, replacement, author)


def replace_commented_span(
    pkg: PptxPackage, comment_id: str, new_text: str, author: str, original_text: str | None = None
) -> str | None:
    """Replace the text a comment is about with `new_text`: its stored
    selection, or -- for a comment attached to a whole shape -- an explicit
    `original_text` (verbatim text from that shape, e.g. proposed by the AI)."""
    doc = build_document_model(pkg)
    comment = doc["comments"].get(comment_id)
    if comment is None:
        raise EditError(f"comment {comment_id} not found")
    if not comment["shapeId"]:
        raise EditError("this comment is attached to the slide, not to any text, so there is nothing to replace")
    shape = next(
        s for sl in doc["slides"] for s in sl["shapes"] if s["id"] == comment["shapeId"]
    )
    quote = original_text or comment.get("quote")
    if not quote:
        raise EditError(
            "this comment is attached to a whole shape rather than selected text; "
            "specify which text to replace"
        )
    anchor = locate_quote_in_shape(shape, quote)
    if anchor is None:
        raise EditError("the commented text was not found in its shape (it may have changed)")
    note_id = replace_text_range(pkg, *anchor, new_text, author)
    if comment.get("quote") and comment["quote"] == quote:
        try:
            comments_mod.set_quote(pkg, comments_mod.find_comment(pkg, comment_id), new_text)
        except CommentError:
            pass
    return note_id


# -- comments -----------------------------------------------------------------


def add_comment(
    pkg: PptxPackage,
    *,
    slide_id: str,
    shape_id: str | None = None,
    quote: str | None = None,
    author: str,
    text: str,
    kind: str = "comment",
    done: bool = False,
) -> str:
    """Add a top-level comment to a slide, optionally attached to a shape
    (and, within it, to a selected-text quote)."""
    slide_idx = parse_slide_id(slide_id)
    if not (0 <= slide_idx < len(pkg.slide_parts)):
        raise EditError(f"slide {slide_id} not found")
    cnv_id = None
    if shape_id:
        sl, cnv_id = _split_shape_id(shape_id)
        if sl != slide_idx:
            raise EditError("shape is not on that slide")
        _shape_el(pkg, slide_idx, cnv_id)  # existence check
    x, y = _shape_pos(pkg, slide_idx, cnv_id)
    return comments_mod.add_legacy_comment(
        pkg, slide_idx, author=author, text=text, pos_emu=(x, y),
        shape_id=int(cnv_id) if cnv_id is not None and cnv_id.isdigit() else None,
        quote=quote, kind=kind, done=done,
    )


def add_comment_for_selection(
    pkg: PptxPackage,
    start_run_id: str,
    start_offset: int,
    end_run_id: str,
    end_offset: int,
    author: str,
    text: str,
) -> str:
    """Anchor a brand-new top-level comment to a user-selected span: the
    comment is attached to the selection's shape and remembers the selected
    text (PowerPoint comments can't point at a text range themselves)."""
    slide_idx, shape, _p, run_els, so, eo = _resolve_selection(
        pkg, start_run_id, start_offset, end_run_id, end_offset
    )
    if len(run_els) == 1:
        quote = run_text(run_els[0])[so:eo]
    else:
        quote = (
            run_text(run_els[0])[so:]
            + "".join(run_text(r) for r in run_els[1:-1])
            + run_text(run_els[-1])[:eo]
        )
    quote = quote.strip()
    if not quote:
        raise EditError("selection is empty")
    return add_comment(
        pkg, slide_id=f"s{slide_idx}", shape_id=f"s{slide_idx}.sh{_cnv_id(shape)}",
        quote=quote, author=author, text=text,
    )


def add_comment_reply(pkg: PptxPackage, parent_comment_id: str, author: str, text: str) -> str:
    try:
        parent = comments_mod.find_comment(pkg, parent_comment_id)
    except CommentError as exc:
        raise EditError(str(exc)) from exc
    return comments_mod.add_reply(pkg, parent, author, text)


def resolve_comment(pkg: PptxPackage, comment_id: str, done: bool = True) -> None:
    try:
        rec = comments_mod.find_comment(pkg, comment_id)
    except CommentError as exc:
        raise EditError(str(exc)) from exc
    comments_mod.set_done(pkg, rec, done)
