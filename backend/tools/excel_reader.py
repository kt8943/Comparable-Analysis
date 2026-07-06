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


def split_tables(rows: list, keywords: set = None,
                 min_text_cells: int = 3, min_kw_frac: float = 0.4) -> list:
    """Split a sheet's rows into STACKED tables, each with its own header row.

    A single sheet can hold several tables with *different* column layouts (e.g. a
    "GFA / Price" table above a "Site Area / Max GFA / Price" table). Mapping the whole
    sheet with one header row then mis-maps every table below the first. This returns
    ``[(header_idx, headers, data_rows), …]`` — one entry per detected header row.

    A row is treated as a header when it has ≥ ``min_text_cells`` text cells and (when
    ``keywords`` is given) a high fraction of them match a known column keyword.
    Consecutive header rows collapse to the first (multi-line headers); data rows that
    themselves look like headers are dropped (secondary header lines).
    """
    def _is_header(row) -> bool:
        text_vals = [str(c).strip() for c in row
                     if c is not None and isinstance(c, str) and str(c).strip()]
        if len(text_vals) < min_text_cells:
            return False
        if not keywords:
            return True
        hits = sum(1 for v in text_vals if any(kw in v.lower() for kw in keywords))
        return hits >= 2 and hits >= len(text_vals) * min_kw_frac

    head_idxs = []
    for i, row in enumerate(rows):
        if _is_header(row):
            if head_idxs and i - head_idxs[-1] == 1:
                continue   # multi-line header → keep the first line only
            head_idxs.append(i)
    if not head_idxs:
        head_idxs = [find_header_row(rows)]

    out = []
    for j, h in enumerate(head_idxs):
        end = head_idxs[j + 1] if j + 1 < len(head_idxs) else len(rows)
        headers = [str(c) if c is not None else "" for c in rows[h]]
        data = [r for r in rows[h + 1:end]
                if any(c not in (None, "") for c in r) and not _is_header(r)]
        if data:
            out.append((h, headers, data))
    if not out:   # nothing had data rows — fall back to a single best-guess table
        h = head_idxs[0]
        out = [(h, [str(c) if c is not None else "" for c in rows[h]],
                [r for r in rows[h + 1:] if any(c not in (None, "") for c in r)])]
    return out


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
