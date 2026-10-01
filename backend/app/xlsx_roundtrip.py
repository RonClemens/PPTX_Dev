"""Parses an uploaded Comment Resolution Matrix .xlsx (filled in by a
government/contractor reviewer, possibly outside this app entirely) back
into decisions this app can apply.

Column headers are matched by normalized name rather than fixed position,
so this works whether the sheet is this app's own TSV export opened and
edited in Excel (which carries a hidden "Comment ID" column for exact
matching) or the reviewer's own blank template filled out independently
(matched by exact comment text instead). The real template's first row is
often a merged "Government Columns"/"Contractor Columns" banner, so the
actual header row is located by scanning rather than assumed to be row 1.
"""
from __future__ import annotations

from typing import Any

NORMALIZED_HEADER_ALIASES: dict[str, str] = {
    "comment id": "comment_id",
    "item #": "item",
    "item": "item",
    "document": "document",
    "rev": "rev",
    "ref #": "ref",
    "ref": "ref",
    "para #": "para",
    "para": "para",
    "comment": "comment",
    "contractor response": "response",
    "contractor response (accept / reject)": "response",
    "contractor response (accept/reject)": "response",
    "response": "response",
    "response (accept / reject)": "response",
    "response (accept/reject)": "response",
    "contractor rationale": "rationale",
    "rationale": "rationale",
    "status": "status",
    "status (open / closed)": "status",
    "status (open/closed)": "status",
    "date adjudicated": "date",
}

RESPONSE_TO_DECISION: dict[str, str] = {
    "accept": "accept",
    "reject": "reject",
    "info only": "info_only",
    "info_only": "info_only",
    "defer": "defer",
}


def _normalize(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().split())


def find_header_row(ws, max_scan_rows: int = 10) -> tuple[int, dict[str, int]] | None:
    """Scan the first few rows for the one containing (at least) recognizable
    "Comment" and "Contractor Response" headers, however they're worded --
    that's the real column header row, even if row 1 is a merged section
    banner above it. Returns (1-indexed row number, {field: 1-indexed col})."""
    max_col = ws.max_column or 1
    scan_rows = min(max_scan_rows, ws.max_row or max_scan_rows)
    for r in range(1, scan_rows + 1):
        col_map: dict[str, int] = {}
        for c in range(1, max_col + 1):
            field = NORMALIZED_HEADER_ALIASES.get(_normalize(ws.cell(row=r, column=c).value))
            if field and field not in col_map:
                col_map[field] = c
        if "comment" in col_map and "response" in col_map:
            return r, col_map
    return None


def _cell_text(value: Any) -> str:
    """A section number like "3.3" typed into a General-formatted cell gets
    silently auto-typed by Excel as a numeric value (unlike "3.2.2", which
    has two dots and so stays text) -- str()'d naively, a whole-number float
    would print as "6.0" instead of "6". Nothing recovers a trailing zero
    Excel already dropped (e.g. "3.10" -> 3.1), but this at least avoids
    inventing new noise for the common whole-number case."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_rows(ws, header_row: int, col_map: dict[str, int]) -> list[dict[str, Any]]:
    """Every non-blank data row below the header, as plain dicts -- blank
    rows (no comment text and no response) are skipped entirely."""

    def cell(row: int, field: str) -> str:
        col = col_map.get(field)
        if not col:
            return ""
        return _cell_text(ws.cell(row=row, column=col).value)

    rows: list[dict[str, Any]] = []
    last_row = ws.max_row or header_row
    for r in range(header_row + 1, last_row + 1):
        comment_text = cell(r, "comment")
        response = cell(r, "response")
        if not comment_text and not response:
            continue
        rows.append(
            {
                "row": r,
                "comment_id": cell(r, "comment_id") or None,
                "ref": cell(r, "ref"),
                "para": cell(r, "para"),
                "comment_text": comment_text,
                "response": response,
                "rationale": cell(r, "rationale"),
            }
        )
    return rows


def decision_from_response(response: str) -> str | None:
    return RESPONSE_TO_DECISION.get(_normalize(response))


def summarize_skips(skipped: list[dict[str, Any]], max_rows_shown: int = 10) -> list[dict[str, Any]]:
    """Group skipped rows by their exact reason (a blank template skips every
    row with the identical "no recognized response" message, and dumping 50+
    duplicate lines in the UI is worse than useless) -- one summary entry per
    distinct reason, with a capped preview of which rows and a total count."""
    order: list[str] = []
    rows_by_reason: dict[str, list[int]] = {}
    for s in skipped:
        reason = s["reason"]
        if reason not in rows_by_reason:
            rows_by_reason[reason] = []
            order.append(reason)
        rows_by_reason[reason].append(s["row"])
    return [
        {
            "reason": reason,
            "count": len(rows_by_reason[reason]),
            "rows": rows_by_reason[reason][:max_rows_shown],
        }
        for reason in order
    ]


def match_comment_id(row: dict[str, Any], comments: dict[str, Any]) -> str | None:
    """Prefer the hidden Comment ID column (exact, unambiguous) when present
    and still valid; otherwise fall back to matching a top-level comment by
    exact (whitespace/case-normalized) text. Returns None if nothing or more
    than one comment matches -- an ambiguous match is treated as no match
    rather than guessed at."""
    cid = row.get("comment_id")
    if cid and cid in comments and not comments[cid].get("parentId"):
        return cid

    target = _normalize(row["comment_text"])
    if not target:
        return None
    matches = [
        c["id"] for c in comments.values() if not c.get("parentId") and _normalize(c["text"]) == target
    ]
    return matches[0] if len(matches) == 1 else None
