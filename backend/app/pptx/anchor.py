"""Resolve a verbatim text quote (e.g. from an AI review pass, or a comment's
stored selection) to the run_id + character-offset anchors that writer.py's
selection-based operations (replace_text_range, add_comment_for_selection)
expect.

This is the AI-review equivalent of what the frontend's text-selection
popup does with a browser Range: turn "this exact span of text" into
(start_run_id, start_offset, end_run_id, end_offset).
"""
from __future__ import annotations

from typing import Any, Iterator


def iter_shape_paragraphs(shape: dict[str, Any]) -> Iterator[dict[str, Any]]:
    """Every paragraph of a shape: its own, or all table-cell paragraphs."""
    if shape.get("type") == "table":
        for row in shape.get("rows", []):
            for cell in row.get("cells", []):
                yield from cell.get("paragraphs", [])
    else:
        yield from shape.get("paragraphs", [])


def locate_quote_in_paragraph(
    paragraph: dict[str, Any], quote: str
) -> tuple[str, int, str, int] | None:
    """Find `quote` as an exact substring of `paragraph`'s text and return
    (start_run_id, start_offset, end_run_id, end_offset). None if the
    paragraph has no matching text (wrong paragraph, or the quote isn't
    verbatim). Line breaks are not part of any run's anchor space, so a
    quote spanning one can't be located."""
    quote = quote.strip()
    if not quote:
        return None

    pieces: list[tuple[str, str]] = [
        (run["id"], run.get("text", ""))
        for run in paragraph.get("runs", [])
        if run.get("type") == "text" and run.get("text")
    ]
    if not pieces:
        return None

    full_text = "".join(text for _, text in pieces)
    idx = full_text.find(quote)
    if idx == -1:
        return None
    end_idx = idx + len(quote)

    def locate(char_pos: int, prefer_start_of_next: bool) -> tuple[str, int]:
        cursor = 0
        for run_id, text in pieces:
            run_end = cursor + len(text)
            boundary_matches = run_end > char_pos if prefer_start_of_next else run_end >= char_pos
            if boundary_matches:
                return run_id, char_pos - cursor
            cursor = run_end
        run_id, text = pieces[-1]
        return run_id, len(text)

    start_run_id, start_offset = locate(idx, prefer_start_of_next=True)
    end_run_id, end_offset = locate(end_idx, prefer_start_of_next=False)
    return start_run_id, start_offset, end_run_id, end_offset


def locate_quote_in_shape(
    shape: dict[str, Any], quote: str
) -> tuple[str, int, str, int] | None:
    """First paragraph of `shape` containing `quote` verbatim."""
    for paragraph in iter_shape_paragraphs(shape):
        located = locate_quote_in_paragraph(paragraph, quote)
        if located:
            return located
    return None


def locate_quote_in_slide(
    slide: dict[str, Any], shape_ids: list[str], quote: str
) -> tuple[str, int, str, int] | None:
    """Try locate_quote_in_shape against each shape in `shape_ids` (in
    order) and return the first match. Used to find which shape within a
    slide a quote belongs to."""
    by_id = {s["id"]: s for s in slide.get("shapes", [])}
    for sid in shape_ids:
        shape = by_id.get(sid)
        if not shape:
            continue
        located = locate_quote_in_shape(shape, quote)
        if located:
            return located
    return None
