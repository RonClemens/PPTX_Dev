"""Document JSON model -> per-slide Markdown (the AI's context).

Each slide is one chunk: a `## Slide N: Title` heading, then the slide's
text shapes in z-order (bullets indented by level), tables as Markdown
tables, pictures/charts as bracketed placeholders, and speaker notes as a
quote block. Comment anchors are encoded with CriticMarkup so both humans
and the AI can read them as plain text:

  commented text   {==span text==}{>>comment:<id><<}
  slide-level      {>>comment:<id> (about the slide as a whole)<<}

PowerPoint has no tracked changes, so unlike the Word version there are no
`{++ins++}`/`{--del--}` markers to render.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .anchor import iter_shape_paragraphs


def _para_text(paragraph: dict[str, Any]) -> str:
    return "".join(r.get("text", "") for r in paragraph.get("runs", []))


def shape_text(shape: dict[str, Any]) -> str:
    """Plain text of a shape or table (paragraphs joined by newlines)."""
    return "\n".join(_para_text(p) for p in iter_shape_paragraphs(shape))


def _top_level_comments_by_shape(doc: dict[str, Any]) -> tuple[dict[str, list[dict]], dict[str, list[dict]]]:
    by_shape: dict[str, list[dict]] = {}
    by_slide: dict[str, list[dict]] = {}
    for c in doc.get("comments", {}).values():
        if c.get("parentId") or c.get("kind") == "edit":
            continue
        if c.get("shapeId"):
            by_shape.setdefault(c["shapeId"], []).append(c)
        else:
            by_slide.setdefault(c["slideId"], []).append(c)
    return by_shape, by_slide


def _mark(text: str, comment: dict[str, Any]) -> str:
    return f"{{=={text}==}}{{>>comment:{comment['id']}<<}}"


def _paragraph_lines(paragraphs: list[dict[str, Any]], comments: list[dict[str, Any]]) -> list[str]:
    """Paragraph texts (with bullets) with this shape's comments marked up."""
    texts = [_para_text(p) for p in paragraphs]
    for c in comments:
        quote = (c.get("quote") or "").strip()
        placed = False
        if quote:
            for i, t in enumerate(texts):
                if quote in t:
                    texts[i] = t.replace(quote, _mark(quote, c), 1)
                    placed = True
                    break
        if not placed:
            # Whole-shape comment: every non-empty paragraph is its own
            # {==...==} span, with the comment tag after the last one.
            nonempty = [i for i, t in enumerate(texts) if t.strip()]
            for n, i in enumerate(nonempty):
                texts[i] = "{==" + texts[i] + "==}"
                if n == len(nonempty) - 1:
                    texts[i] += f"{{>>comment:{c['id']}<<}}"
    lines = []
    for p, t in zip(paragraphs, texts):
        if not t.strip():
            continue
        if p.get("bullet"):
            lines.append(f"{'  ' * p.get('level', 0)}- {t}")
        else:
            lines.append(t)
    return lines


def table_markdown(table: dict[str, Any], comments: list[dict[str, Any]]) -> str:
    lines = []
    for ri, row in enumerate(table.get("rows", [])):
        texts = []
        for cell in row.get("cells", []):
            cell_text = " ".join(_para_text(p) for p in cell.get("paragraphs", [])).strip()
            texts.append(cell_text.replace("|", "\\|") or " ")
        lines.append("| " + " | ".join(texts) + " |")
        if ri == 0:
            lines.append("| " + " | ".join(["---"] * len(texts)) + " |")
    md = "\n".join(lines)
    for c in comments:
        quote = (c.get("quote") or "").strip()
        if quote and quote in md:
            md = md.replace(quote, _mark(quote, c), 1)
        else:
            md += f"\n{{>>comment:{c['id']}<<}}"
    return md


def shape_markdown(shape: dict[str, Any], comments: list[dict[str, Any]]) -> str:
    t = shape.get("type")
    if t == "table":
        return table_markdown(shape, comments)
    if t == "picture":
        label = shape.get("descr") or shape.get("name") or "picture"
        md = f"[Picture: {label}]"
    elif t == "other":
        md = f"[{shape.get('label') or 'Embedded object'}: {shape.get('name') or ''}]".replace(": ]", "]")
    elif t == "shape":
        md = "\n".join(_paragraph_lines(shape.get("paragraphs", []), comments))
        if not md.strip() and comments:
            md = f"[Shape: {shape.get('name') or 'unnamed'}]" + "".join(
                f"{{>>comment:{c['id']}<<}}" for c in comments
            )
        return md
    else:
        return ""
    for c in comments:
        md += f"{{>>comment:{c['id']}<<}}"
    return md


def slide_markdown(slide: dict[str, Any], by_shape, by_slide) -> str:
    heading = f"## Slide {slide['index']}: {slide['title']}" if slide.get("title") else f"## Slide {slide['index']}"
    if slide.get("hidden"):
        heading += " (hidden)"
    parts = [heading]
    for c in by_slide.get(slide["id"], []):
        parts.append(f"{{>>comment:{c['id']} (about the slide as a whole)<<}}")
    for shape in slide.get("shapes", []):
        if shape.get("isTitle"):
            for c in by_shape.get(shape["id"], []):
                parts.append(f"{{>>comment:{c['id']} (on the slide title)<<}}")
            continue
        md = shape_markdown(shape, by_shape.get(shape["id"], []))
        if md.strip():
            parts.append(md)
    if slide.get("notes"):
        parts.append("\n".join("> " + ln for ln in ("Speaker notes:\n" + slide["notes"]).splitlines()))
    return "\n\n".join(parts)


def render_markdown(doc: dict[str, Any]) -> str:
    by_shape, by_slide = _top_level_comments_by_shape(doc)
    return "\n\n".join(slide_markdown(s, by_shape, by_slide) for s in doc.get("slides", []))


@dataclass
class Chunk:
    id: str
    heading: str | None
    level: int
    start_index: int
    end_index: int
    markdown: str
    item_ids: list[str]
    heading_number: str | None = None
    slide_id: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "heading": self.heading,
            "level": self.level,
            "startIndex": self.start_index,
            "endIndex": self.end_index,
            "markdown": self.markdown,
            "itemIds": self.item_ids,
            "slideId": self.slide_id,
        }


def build_chunks(doc: dict[str, Any]) -> list[Chunk]:
    """One chunk per slide."""
    by_shape, by_slide = _top_level_comments_by_shape(doc)
    chunks: list[Chunk] = []
    for i, slide in enumerate(doc.get("slides", [])):
        chunks.append(
            Chunk(
                id=f"slide{i}",
                heading=slide.get("title") or f"Slide {slide['index']}",
                level=1,
                start_index=i,
                end_index=i,
                markdown=slide_markdown(slide, by_shape, by_slide),
                item_ids=[s["id"] for s in slide.get("shapes", [])],
                heading_number=f"Slide {slide['index']}",
                slide_id=slide["id"],
            )
        )
    return chunks


def context_markdown_for_comment(doc: dict[str, Any], comment_id: str) -> str | None:
    """Markdown of the slide containing the anchor for `comment_id`, used as
    AI context when the whole deck is too big to send."""
    comment = doc.get("comments", {}).get(comment_id)
    if comment is None:
        return None
    for chunk in build_chunks(doc):
        if chunk.slide_id == comment["slideId"]:
            return chunk.markdown
    return None


# -- Best-effort Ref #/Para # derivation for the Comment Resolution Matrix --
#
# Ref # is the slide a comment lives on ("Slide 3"). Para # is a requirement
# ID like "ACA-HRS-0190" found in the commented shape's own text, when there
# is one; otherwise callers fall back to this app's internal chunk:shape
# reference rather than guessing wrong.

# "ACA-HRS-0190", "SSS-00340.021" -- a letter prefix, 0-4 more hyphen/
# underscore-separated segments, then a required final hyphen/underscore +
# numeric segment with an optional dotted-decimal suffix (".021"). Case-
# insensitive, hyphen or underscore as separator. Loose by design; tune as
# real requirement-ID conventions turn up.
_REQUIREMENT_ID_RE = re.compile(
    r"\b[A-Za-z]{2,10}(?:[-_][A-Za-z0-9]{1,10}){0,4}[-_]\d{1,6}(?:\.\d{1,6})?\b"
)
_SLIDE_REF_RE = re.compile(r"^\s*(?:slide\s*#?\s*)?(\d+)\s*$", re.IGNORECASE)


def _normalize_locator(value: str | None) -> str:
    return " ".join((value or "").strip().lower().split())


def _derive_para_ref(shape: dict[str, Any] | None) -> str | None:
    if not shape:
        return None
    m = _REQUIREMENT_ID_RE.search(shape_text(shape))
    return m.group(0) if m else None


def find_comment_anchor(
    doc: dict[str, Any], ref: str | None, para: str | None
) -> dict[str, Any] | None:
    """Reverse of comment_locations' derivedRef/derivedPara guess: given a
    Ref #/Para # from an imported CRM row that doesn't match any existing
    comment (a clean master deck paired with a separately-maintained review
    spreadsheet), find where a brand-new comment should be attached.
    Para # (a requirement ID -- specific to one shape) is preferred over
    Ref # ("Slide N" or just N -- only slide-level, so it anchors on the
    slide's title). Returns {"slide_id", "shape_id", "quote"} or None when
    nothing recognizable matches, rather than guessing."""
    ref_n = _normalize_locator(ref)
    para_n = _normalize_locator(para)
    if not ref_n and not para_n:
        return None

    slides = doc.get("slides", [])
    if para_n:
        for slide in slides:
            for shape in slide.get("shapes", []):
                derived = _derive_para_ref(shape)
                if derived and _normalize_locator(derived) == para_n:
                    return {"slide_id": slide["id"], "shape_id": shape["id"], "quote": derived}

    m = _SLIDE_REF_RE.match(ref or "")
    if m:
        n = int(m.group(1))
        for slide in slides:
            if slide["index"] == n:
                title = next((s for s in slide.get("shapes", []) if s.get("isTitle")), None)
                return {"slide_id": slide["id"], "shape_id": title["id"] if title else None, "quote": None}
    return None


def comment_locations(doc: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Map each comment id to where it lives -- the slide (chunk) and the
    shape within it -- plus a best-effort guess at the Comment Resolution
    Matrix's Ref # (the slide) and Para # (requirement ID in the shape)."""
    shapes_by_id = {s["id"]: s for sl in doc.get("slides", []) for s in sl.get("shapes", [])}
    chunk_by_slide = {c.slide_id: c for c in build_chunks(doc)}
    locations: dict[str, dict[str, Any]] = {}
    for cid, c in doc.get("comments", {}).items():
        chunk = chunk_by_slide.get(c["slideId"])
        if chunk is None:
            continue
        item_id = c.get("shapeId") or c["slideId"]
        locations[cid] = {
            "sectionId": chunk.id,
            "sectionHeading": chunk.heading,
            "itemId": item_id,
            "paraRef": f"{chunk.id}:{item_id}",
            "derivedRef": chunk.heading_number,
            "derivedPara": _derive_para_ref(shapes_by_id.get(c.get("shapeId") or "")),
        }
    return locations
