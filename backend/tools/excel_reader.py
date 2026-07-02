"""
tools/excel_reader.py
=====================
Excel structure utilities shared across all comp pipelines.
No LLM calls — pure openpyxl helpers.

Public API
----------
find_best_sheet(wb, keywords: set[str]) -> str
    Return the sheet name most likely to contain comp data.

find_header_row(rows: list, min_text_cells: int = 3) -> int
    Return 0-based index of the first header-like row.
"""

import re


def find_best_sheet(wb, keywords: set) -> str:
    """Return the sheet name most likely to contain comp data.

    Scoring: primary = how many keyword-matching cells appear in the best row;
    tiebreaker = total non-empty rows.  Avoids picking pivot/summary sheets
    that have more rows but wrong column names.

    keywords should be a set of lowercase strings (≥ 3 chars) derived from
    the output column names of the comp type being parsed.
    """
    def _score(name):
        ws = wb[name]
        if not hasattr(ws, "iter_rows"):
            return (0, 0)
        best_hits = total_rows = 0
        for row in ws.iter_rows():
            vals = [str(c.value).lower() for c in row
                    if c.value not in (None, "") and str(c.value).strip()]
            if not vals:
                continue
            total_rows += 1
            hits = sum(1 for v in vals if any(kw in v for kw in keywords))
            best_hits = max(best_hits, hits)
        return (best_hits, total_rows)

    return max(wb.sheetnames, key=_score)


def find_header_row(rows: list, min_text_cells: int = 3) -> int:
    """Return 0-based index of the first row that looks like a header."""
    for i, row in enumerate(rows):
        text_vals = [c for c in row if c is not None and isinstance(c, str) and c.strip()]
        if len(text_vals) >= min_text_cells:
            return i
    return 0


def sheet_keywords(output_fields: list) -> set:
    """Build the keyword set for find_best_sheet from an _OUTPUT_FIELDS list.

    output_fields is a list of (col_name, internal_key, description) tuples.
    Returns the set of lowercase words (≥ 3 chars) from the column names.
    """
    return {
        k
        for col, *_ in output_fields
        for k in re.sub(r"[^a-z0-9 ]", " ", col.lower()).split()
        if len(k) >= 3
    }
