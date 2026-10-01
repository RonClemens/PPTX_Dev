"""Low-level OOXML (.pptx) access: namespaces, zip I/O, relationships.

A .pptx is a zip of XML parts. Parts we need to read are parsed lazily into
live lxml trees (`PptxPackage.tree`); only parts a writer explicitly marks
dirty are re-serialized on save, so every slide, layout, theme, media file,
etc. that we didn't touch round-trips byte-for-byte. Comment parts
(ppt/commentAuthors.xml, ppt/comments/*.xml) are created on demand, with
their relationships and content-type overrides, when a deck has none yet.
"""
from __future__ import annotations

import io
import posixpath
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "pdev": "urn:pptx-dev:comment-state",
}


def qn(tag: str) -> str:
    """Convert 'p:sp' -> '{namespace}sp'."""
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def local(el) -> str:
    return etree.QName(el).localname


PRESENTATION_PART = "ppt/presentation.xml"
CONTENT_TYPES_PART = "[Content_Types].xml"
COMMENT_AUTHORS_PART = "ppt/commentAuthors.xml"
MODERN_AUTHORS_PART = "ppt/authors.xml"

REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
REL_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_NOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_COMMENT_AUTHORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
REL_MODERN_COMMENTS = "http://schemas.microsoft.com/office/2018/10/relationships/comments"

CT_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
CT_COMMENT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"

EMU_PER_INCH = 914400


def resolve_target(source_part: str, target: str) -> str:
    """A relationship Target (relative to the source part's directory, or
    absolute with a leading '/') -> a package-root-relative part name."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def rels_part_name(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def _serialize(el) -> bytes:
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


class PptxPackage:
    """An in-memory, mutable representation of a .pptx package."""

    def __init__(self) -> None:
        self.parts: dict[str, bytes] = {}
        self._trees: dict[str, etree._Element] = {}
        self._dirty: set[str] = set()
        self.slide_parts: list[str] = []
        self.slide_size: tuple[int, int] = (9144000, 6858000)

    # -- load / save ------------------------------------------------------

    @classmethod
    def load(cls, path: str | Path) -> "PptxPackage":
        pkg = cls()
        with zipfile.ZipFile(path, "r") as zf:
            for info in zf.infolist():
                pkg.parts[info.filename] = zf.read(info.filename)
        if PRESENTATION_PART not in pkg.parts:
            raise ValueError("not a .pptx package: ppt/presentation.xml is missing")

        pres = pkg.tree(PRESENTATION_PART)
        rid_to_part = {r["Id"]: resolve_target(PRESENTATION_PART, r["Target"]) for r in pkg.rels(PRESENTATION_PART)
                       if r["Type"] == REL_SLIDE}
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        if sld_id_lst is not None:
            for sld_id in sld_id_lst:
                part = rid_to_part.get(sld_id.get(qn("r:id")))
                if part and part in pkg.parts:
                    pkg.slide_parts.append(part)
        sz = pres.find(qn("p:sldSz"))
        if sz is not None and sz.get("cx") and sz.get("cy"):
            pkg.slide_size = (int(sz.get("cx")), int(sz.get("cy")))
        return pkg

    def save(self, path: str | Path) -> None:
        for part in self._dirty:
            self.parts[part] = _serialize(self._trees[part])
        self._dirty.clear()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            names = list(self.parts)
            if CONTENT_TYPES_PART in names:  # conventionally the first entry
                names.remove(CONTENT_TYPES_PART)
                names.insert(0, CONTENT_TYPES_PART)
            for name in names:
                zf.writestr(name, self.parts[name])
        Path(path).write_bytes(buf.getvalue())

    # -- parts ------------------------------------------------------------

    def has_part(self, part: str) -> bool:
        return part in self.parts or part in self._trees

    def tree(self, part: str) -> etree._Element:
        if part not in self._trees:
            self._trees[part] = etree.fromstring(self.parts[part])
        return self._trees[part]

    def mark_dirty(self, part: str) -> None:
        if part not in self._trees:
            self.tree(part)
        self._dirty.add(part)

    def add_part(self, part: str, root: etree._Element) -> None:
        """Register a brand-new XML part (kept as a live tree, saved dirty)."""
        self.parts.setdefault(part, b"")
        self._trees[part] = root
        self._dirty.add(part)

    def add_override(self, part: str, content_type: str) -> None:
        ct = self.tree(CONTENT_TYPES_PART)
        name = "/" + part
        for o in ct.findall(qn("ct:Override")):
            if o.get("PartName") == name:
                return
        o = etree.SubElement(ct, qn("ct:Override"))
        o.set("PartName", name)
        o.set("ContentType", content_type)
        self.mark_dirty(CONTENT_TYPES_PART)

    # -- relationships ----------------------------------------------------

    def rels(self, part: str) -> list[dict[str, str]]:
        rp = rels_part_name(part)
        if not self.has_part(rp):
            return []
        return [
            {"Id": r.get("Id", ""), "Type": r.get("Type", ""), "Target": r.get("Target", ""),
             "TargetMode": r.get("TargetMode", "")}
            for r in self.tree(rp)
        ]

    def related_part(self, part: str, rel_type: str) -> str | None:
        for r in self.rels(part):
            if r["Type"] == rel_type and r["TargetMode"] != "External":
                target = resolve_target(part, r["Target"])
                if self.has_part(target):
                    return target
        return None

    def related_parts(self, part: str, rel_type: str) -> list[str]:
        out = []
        for r in self.rels(part):
            if r["Type"] == rel_type and r["TargetMode"] != "External":
                target = resolve_target(part, r["Target"])
                if self.has_part(target):
                    out.append(target)
        return out

    def add_rel(self, source_part: str, rel_type: str, target_part: str) -> str:
        rp = rels_part_name(source_part)
        if not self.has_part(rp):
            self.add_part(
                rp, etree.fromstring(f'<Relationships xmlns="{NS["pr"]}"/>'.encode("utf-8"))
            )
        rels = self.tree(rp)
        nums = [int(r.get("Id")[3:]) for r in rels if (r.get("Id") or "").startswith("rId") and r.get("Id")[3:].isdigit()]
        new_id = f"rId{(max(nums) + 1) if nums else 1}"
        rel = etree.SubElement(rels, qn("pr:Relationship"))
        rel.set("Id", new_id)
        rel.set("Type", rel_type)
        rel.set("Target", posixpath.relpath(target_part, posixpath.dirname(source_part) or "."))
        self.mark_dirty(rp)
        return new_id

    # -- slides -----------------------------------------------------------

    def slide_tree(self, idx: int) -> etree._Element:
        return self.tree(self.slide_parts[idx])

    def layout_part(self, slide_part: str) -> str | None:
        return self.related_part(slide_part, REL_LAYOUT)

    def master_part(self, layout_part: str | None) -> str | None:
        return self.related_part(layout_part, REL_MASTER) if layout_part else None

    def theme_part(self, master_part: str | None) -> str | None:
        return self.related_part(master_part, REL_THEME) if master_part else None

    def resolve_media(self, slide_idx: int, rel_id: str) -> tuple[bytes, str] | None:
        """(bytes, part_name) for a slide's image relationship."""
        if not (0 <= slide_idx < len(self.slide_parts)):
            return None
        slide_part = self.slide_parts[slide_idx]
        for r in self.rels(slide_part):
            if r["Id"] == rel_id and r["TargetMode"] != "External":
                part = resolve_target(slide_part, r["Target"])
                data = self.parts.get(part)
                if data is not None:
                    return data, part
        return None
