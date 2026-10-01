"""PowerPoint comments: reading both on-disk flavors, writing the legacy one.

PowerPoint has two comment formats:

* **Legacy** (PowerPoint 2007-2016, still read by every version): one
  `ppt/comments/commentN.xml` part per slide (`p:cmLst` of `p:cm`), authors
  in `ppt/commentAuthors.xml`. A comment is positioned by a bare `p:pos`
  (x, y in 1/576 inch) -- it is *not* attached to a shape or text range.
  Replies are `p:cm` elements carrying a `p15:threadingInfo/p15:parentCm`
  extension. There is no resolved flag.
* **Modern** (Microsoft 365): `ppt/comments/modernComment_*.xml`
  (`p188:cmLst`), authors in `ppt/authors.xml`, replies nested in
  `p188:replyLst`, `status="resolved"` for resolved threads, anchored to a
  shape by an `ac:spMk` moniker.

Both are *read* into one flat list of `CommentRec`. We *write* the legacy
format for new comments and replies (universally readable, and simple
enough to emit correctly), and modern-style edits (reply, resolve) on
comments that already live in a modern part.

What legacy comments can't express -- which shape a comment is about, the
exact text a reviewer selected, a resolved flag, and whether it's one of
this app's automatic "edit record" notes -- is stored in a private
`p:extLst` extension on the `p:cm` (`pdev:state`). PowerPoint ignores
extensions it doesn't know, so this is harmless to it, and it keeps the
file self-contained: no sidecar database is needed to round-trip.
"""
from __future__ import annotations

import datetime as dt
import uuid
from dataclasses import dataclass, field

from lxml import etree

from .ooxml import (
    COMMENT_AUTHORS_PART,
    CT_COMMENT_AUTHORS,
    CT_COMMENTS,
    MODERN_AUTHORS_PART,
    NS,
    PRESENTATION_PART,
    REL_COMMENT_AUTHORS,
    REL_COMMENTS,
    REL_MODERN_COMMENTS,
    PptxPackage,
    local,
    qn,
)

# Private extension URI for our own per-comment state.
PDEV_EXT_URI = "{7B9A6C1E-3F52-4C0B-9E1A-5D7F0C2A4B10}"
THREADING_EXT_URI = "{C676402C-5697-4E1C-873F-D02D1690AC5C}"  # p15:threadingInfo

EMU_PER_LEGACY_UNIT = 914400 / 576  # legacy p:pos is in 1/576 inch


class CommentError(ValueError):
    pass


def now_legacy() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000")


def now_modern() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


@dataclass
class CommentRec:
    id: str
    flavor: str  # "legacy" | "modern"
    slide_idx: int
    author: str
    initials: str | None
    date: str | None
    text: str
    done: bool
    parent_id: str | None
    pos_emu: tuple[int, int] | None
    shape_id: int | None
    quote: str | None
    kind: str  # "comment" | "edit"
    el: etree._Element = field(repr=False, default=None)  # type: ignore[assignment]
    part: str = ""


def _initials(name: str) -> str:
    parts = [p for p in name.replace("_", " ").split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


# -- legacy authors ---------------------------------------------------------


def _legacy_author_root(pkg: PptxPackage, create: bool = False):
    if pkg.has_part(COMMENT_AUTHORS_PART):
        return pkg.tree(COMMENT_AUTHORS_PART)
    if not create:
        return None
    root = etree.Element(qn("p:cmAuthorLst"), nsmap={"p": NS["p"]})
    pkg.add_part(COMMENT_AUTHORS_PART, root)
    pkg.add_override(COMMENT_AUTHORS_PART, CT_COMMENT_AUTHORS)
    pkg.add_rel(PRESENTATION_PART, REL_COMMENT_AUTHORS, COMMENT_AUTHORS_PART)
    return root


def _legacy_authors(pkg: PptxPackage) -> dict[str, tuple[str, str | None]]:
    root = _legacy_author_root(pkg)
    out: dict[str, tuple[str, str | None]] = {}
    if root is not None:
        for a in root:
            if local(a) == "cmAuthor":
                out[a.get("id", "")] = (a.get("name") or "Unknown", a.get("initials"))
    return out


def _ensure_legacy_author(pkg: PptxPackage, name: str) -> etree._Element:
    root = _legacy_author_root(pkg, create=True)
    for a in root:
        if local(a) == "cmAuthor" and a.get("name") == name:
            return a
    ids = [int(a.get("id")) for a in root if local(a) == "cmAuthor" and (a.get("id") or "").isdigit()]
    new_id = (max(ids) + 1) if ids else 0
    a = etree.SubElement(root, qn("p:cmAuthor"))
    a.set("id", str(new_id))
    a.set("name", name)
    a.set("initials", _initials(name))
    a.set("lastIdx", "0")
    a.set("clrIdx", str(new_id % 10))
    pkg.mark_dirty(COMMENT_AUTHORS_PART)
    return a


# -- reading ----------------------------------------------------------------


def _pdev_state(cm: etree._Element) -> etree._Element | None:
    for el in cm.iter(qn("pdev:state")):
        return el
    return None


def _ensure_pdev_state(cm: etree._Element) -> etree._Element:
    state = _pdev_state(cm)
    if state is not None:
        return state
    ext_lst = cm.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(cm, qn("p:extLst"))
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", PDEV_EXT_URI)
    return etree.SubElement(ext, qn("pdev:state"), nsmap={"pdev": NS["pdev"]})


def _text_of_txbody(tx) -> str:
    paras = []
    for p in tx.iter("{%s}p" % NS["a"]):
        paras.append("".join(t.text or "" for t in p.iter("{%s}t" % NS["a"])))
    return "\n".join(paras)


def _read_legacy(pkg: PptxPackage, slide_idx: int, part: str, authors) -> list[CommentRec]:
    recs: list[CommentRec] = []
    root = pkg.tree(part)
    for cm in root:
        if local(cm) != "cm":
            continue
        aid, idx = cm.get("authorId", ""), cm.get("idx", "")
        name, initials = authors.get(aid, ("Unknown", None))
        pos_el = cm.find(qn("p:pos"))
        pos = None
        if pos_el is not None:
            try:
                pos = (round(int(pos_el.get("x", "0")) * EMU_PER_LEGACY_UNIT),
                       round(int(pos_el.get("y", "0")) * EMU_PER_LEGACY_UNIT))
            except ValueError:
                pos = None
        text_el = cm.find(qn("p:text"))
        parent_id = None
        for pc in cm.iter(qn("p15:parentCm")):
            parent_id = f"{pc.get('authorId')}-{pc.get('idx')}"
        state = _pdev_state(cm)
        shape_id = None
        if state is not None and (state.get("shape") or "").isdigit():
            shape_id = int(state.get("shape"))
        recs.append(
            CommentRec(
                id=f"{aid}-{idx}",
                flavor="legacy",
                slide_idx=slide_idx,
                author=name,
                initials=initials,
                date=cm.get("dt"),
                text=(text_el.text or "") if text_el is not None else "",
                done=state is not None and state.get("done") == "1",
                parent_id=parent_id,
                pos_emu=pos,
                shape_id=shape_id,
                quote=(state.get("quote") or None) if state is not None else None,
                kind=(state.get("kind") if state is not None and state.get("kind") else "comment"),
                el=cm,
                part=part,
            )
        )
    return recs


def _modern_authors(pkg: PptxPackage) -> dict[str, tuple[str, str | None]]:
    out: dict[str, tuple[str, str | None]] = {}
    if pkg.has_part(MODERN_AUTHORS_PART):
        for a in pkg.tree(MODERN_AUTHORS_PART):
            if local(a) == "author":
                out[a.get("id", "")] = (a.get("name") or "Unknown", a.get("initials"))
    return out


def _read_modern(pkg: PptxPackage, slide_idx: int, part: str, authors) -> list[CommentRec]:
    recs: list[CommentRec] = []
    for cm in pkg.tree(part):
        if local(cm) != "cm":
            continue
        cid = cm.get("id", "")
        name, initials = authors.get(cm.get("authorId", ""), ("Unknown", None))
        shape_id = None
        for el in cm.iter():
            if local(el) == "spMk" and (el.get("id") or "").isdigit():
                shape_id = int(el.get("id"))
                break
        tx = next((c for c in cm if local(c) == "txBody"), None)
        recs.append(
            CommentRec(
                id=cid, flavor="modern", slide_idx=slide_idx, author=name, initials=initials,
                date=cm.get("created"), text=_text_of_txbody(tx) if tx is not None else "",
                done=cm.get("status") in ("resolved", "closed"), parent_id=None, pos_emu=None,
                shape_id=shape_id, quote=None, kind="comment", el=cm, part=part,
            )
        )
        reply_lst = next((c for c in cm if local(c) == "replyLst"), None)
        if reply_lst is not None:
            for rp in reply_lst:
                if local(rp) != "reply":
                    continue
                rname, rinit = authors.get(rp.get("authorId", ""), ("Unknown", None))
                rtx = next((c for c in rp if local(c) == "txBody"), None)
                recs.append(
                    CommentRec(
                        id=rp.get("id", ""), flavor="modern", slide_idx=slide_idx, author=rname,
                        initials=rinit, date=rp.get("created"),
                        text=_text_of_txbody(rtx) if rtx is not None else "", done=False,
                        parent_id=cid, pos_emu=None, shape_id=shape_id, quote=None, kind="comment",
                        el=rp, part=part,
                    )
                )
    return recs


def read_comments(pkg: PptxPackage) -> list[CommentRec]:
    legacy_authors = _legacy_authors(pkg)
    modern_authors = _modern_authors(pkg)
    recs: list[CommentRec] = []
    for i, slide_part in enumerate(pkg.slide_parts):
        for part in pkg.related_parts(slide_part, REL_COMMENTS):
            recs.extend(_read_legacy(pkg, i, part, legacy_authors))
        for part in pkg.related_parts(slide_part, REL_MODERN_COMMENTS):
            recs.extend(_read_modern(pkg, i, part, modern_authors))
    return recs


def find_comment(pkg: PptxPackage, comment_id: str) -> CommentRec:
    for rec in read_comments(pkg):
        if rec.id == comment_id:
            return rec
    raise CommentError(f"comment {comment_id} not found")


# -- writing ----------------------------------------------------------------


def _ensure_legacy_comments_part(pkg: PptxPackage, slide_idx: int) -> str:
    slide_part = pkg.slide_parts[slide_idx]
    existing = pkg.related_part(slide_part, REL_COMMENTS)
    if existing:
        return existing
    n = 1
    while pkg.has_part(f"ppt/comments/comment{n}.xml"):
        n += 1
    part = f"ppt/comments/comment{n}.xml"
    pkg.add_part(part, etree.Element(qn("p:cmLst"), nsmap={"p": NS["p"]}))
    pkg.add_override(part, CT_COMMENTS)
    pkg.add_rel(slide_part, REL_COMMENTS, part)
    return part


def add_legacy_comment(
    pkg: PptxPackage,
    slide_idx: int,
    *,
    author: str,
    text: str,
    pos_emu: tuple[int, int],
    shape_id: int | None = None,
    quote: str | None = None,
    kind: str = "comment",
    done: bool = False,
    parent: CommentRec | None = None,
) -> str:
    part = _ensure_legacy_comments_part(pkg, slide_idx)
    root = pkg.tree(part)
    author_el = _ensure_legacy_author(pkg, author)
    idx = int(author_el.get("lastIdx", "0")) + 1
    author_el.set("lastIdx", str(idx))
    pkg.mark_dirty(COMMENT_AUTHORS_PART)

    cm = etree.SubElement(root, qn("p:cm"))
    cm.set("authorId", author_el.get("id"))
    cm.set("dt", now_legacy())
    cm.set("idx", str(idx))
    pos = etree.SubElement(cm, qn("p:pos"))
    pos.set("x", str(max(0, round(pos_emu[0] / EMU_PER_LEGACY_UNIT))))
    pos.set("y", str(max(0, round(pos_emu[1] / EMU_PER_LEGACY_UNIT))))
    etree.SubElement(cm, qn("p:text")).text = text

    ext_lst = etree.SubElement(cm, qn("p:extLst"))
    if parent is not None:
        p_aid, _, p_idx = parent.id.partition("-")
        ext = etree.SubElement(ext_lst, qn("p:ext"))
        ext.set("uri", THREADING_EXT_URI)
        ti = etree.SubElement(ext, qn("p15:threadingInfo"), nsmap={"p15": NS["p15"]})
        ti.set("timeZoneBias", "0")
        pc = etree.SubElement(ti, qn("p15:parentCm"))
        pc.set("authorId", p_aid)
        pc.set("idx", p_idx)
    if shape_id is not None or quote or kind != "comment" or done:
        state = _ensure_pdev_state(cm)
        if shape_id is not None:
            state.set("shape", str(shape_id))
        if quote:
            state.set("quote", quote)
        if kind != "comment":
            state.set("kind", kind)
        if done:
            state.set("done", "1")
    if len(ext_lst) == 0:
        cm.remove(ext_lst)

    pkg.mark_dirty(part)
    return f"{author_el.get('id')}-{idx}"


def _ensure_modern_author(pkg: PptxPackage, name: str) -> str:
    root = pkg.tree(MODERN_AUTHORS_PART)
    for a in root:
        if local(a) == "author" and a.get("name") == name:
            return a.get("id")
    new_id = "{" + str(uuid.uuid4()).upper() + "}"
    a = etree.SubElement(root, "{%s}author" % NS["p188"])
    a.set("id", new_id)
    a.set("name", name)
    a.set("initials", _initials(name))
    a.set("userId", name)
    a.set("providerId", "None")
    pkg.mark_dirty(MODERN_AUTHORS_PART)
    return new_id


def _modern_txbody(text: str) -> etree._Element:
    A = NS["a"]
    tx = etree.Element("{%s}txBody" % NS["p188"], nsmap={"a": A})
    etree.SubElement(tx, "{%s}bodyPr" % A)
    etree.SubElement(tx, "{%s}lstStyle" % A)
    p = etree.SubElement(tx, "{%s}p" % A)
    r = etree.SubElement(p, "{%s}r" % A)
    etree.SubElement(r, "{%s}rPr" % A).set("lang", "en-US")
    etree.SubElement(r, "{%s}t" % A).text = text
    return tx


def add_reply(pkg: PptxPackage, parent: CommentRec, author: str, text: str) -> str:
    if parent.flavor == "legacy":
        return add_legacy_comment(
            pkg, parent.slide_idx, author=author, text=text,
            pos_emu=parent.pos_emu or (0, 0), shape_id=parent.shape_id, parent=parent,
        )
    top = parent.el
    reply_lst = next((c for c in top if local(c) == "replyLst"), None)
    if reply_lst is None:
        reply_lst = etree.Element("{%s}replyLst" % NS["p188"])
        tx = next((c for c in top if local(c) == "txBody"), None)
        if tx is not None:
            tx.addprevious(reply_lst)
        else:
            top.append(reply_lst)
    rid = "{" + str(uuid.uuid4()).upper() + "}"
    rp = etree.SubElement(reply_lst, "{%s}reply" % NS["p188"])
    rp.set("id", rid)
    rp.set("authorId", _ensure_modern_author(pkg, author))
    rp.set("created", now_modern())
    rp.append(_modern_txbody(text))
    pkg.mark_dirty(parent.part)
    return rid


def set_done(pkg: PptxPackage, rec: CommentRec, done: bool) -> None:
    if rec.flavor == "modern":
        rec.el.set("status", "resolved" if done else "active")
    else:
        state = _ensure_pdev_state(rec.el)
        if done:
            state.set("done", "1")
        elif "done" in state.attrib:
            del state.attrib["done"]
    pkg.mark_dirty(rec.part)


def set_quote(pkg: PptxPackage, rec: CommentRec, quote: str | None) -> None:
    """Update the stored selected-text quote (legacy comments only; modern
    comments carry no quote of ours)."""
    if rec.flavor != "legacy":
        return
    state = _ensure_pdev_state(rec.el)
    if quote:
        state.set("quote", quote)
    elif "quote" in state.attrib:
        del state.attrib["quote"]
    pkg.mark_dirty(rec.part)
