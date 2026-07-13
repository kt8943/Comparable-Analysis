"""
backend/tools/report_period.py
==============================
Infer a **report period** (e.g. "Q2 2025") from a broker file's TITLE (filename) or
COVER PAGE, and use it to backfill comps that have no row-level transaction date.

Why: broker comp tables often omit a date per comp, but the report itself is dated
("2025 Q2 Colliers … Outlook.pdf", or "Q2 2025" on the cover). A comp in that report
is at most as recent as the report, so an inferred period is a reasonable, clearly-
marked fallback (prefixed "~" for circa) — it fills the blank Date column and lets the
recency filter work. It NEVER overwrites a real date.

Public API:
  infer_report_period(path)        → "Q2 2025" | "Jun 2025" | "2025" | ""
  backfill_missing_dates(records, date_field, excel_files, pdf_files, image_files)
                                   → int (number of blank dates filled), in place
"""
from __future__ import annotations

import re
from pathlib import Path

_MONTHS = ("jan", "feb", "mar", "apr", "may", "jun",
           "jul", "aug", "sep", "oct", "nov", "dec")


def _yr4(y) -> int:
    y = int(y)
    return 2000 + y if y < 100 else y


def _match_quarter(s: str) -> str:
    for pat, qi, yi in (
        (r"q\s*([1-4])[\s'./_-]*(20\d{2}|\d{2})(?!\d)", 1, 2),        # Q2 2025 / Q2'25
        (r"(?<!\d)([1-4])\s*q[\s'./_-]*(20\d{2}|\d{2})(?!\d)", 1, 2),  # 2Q25 / 2Q 2025
        (r"(20\d{2})[\s./_-]*q\s*([1-4])(?!\d)", 2, 1),              # 2025 Q2
    ):
        m = re.search(pat, s)
        if m:
            return f"Q{m.group(qi)} {_yr4(m.group(yi))}"
    return ""


def _match_half(s: str) -> str:
    # Map a half-year to the quarter that ENDS it (H1→Q2, H2→Q4) — conservative recency.
    # Number-first ("1H 2024") is tried before "H1 2024" so the half digit is never
    # confused with the first digit of the year.
    for pat, hi, yi in (
        (r"(?<!\d)([12])\s*h[\s'./_-]*(20\d{2}|\d{2})(?!\d)", 1, 2),  # 1H 2024 / 1H24
        (r"(20\d{2})[\s./_-]*h\s*([12])(?!\d)", 2, 1),               # 2024 H1
        (r"h\s*([12])[\s'./_-]+(20\d{2})(?!\d)", 1, 2),              # H1 2024
    ):
        m = re.search(pat, s)
        if m:
            return f"Q{'2' if m.group(hi) == '1' else '4'} {_yr4(m.group(yi))}"
    return ""


def _match_month(s: str) -> str:
    m = re.search(
        r"(?<![a-z])(" + "|".join(_MONTHS) + r")[a-z]*\.?[\s'./_-]*(20\d{2}|'?\d{2})(?!\d)", s)
    if m:
        yr = m.group(2).lstrip("'")
        return f"{m.group(1).title()} {_yr4(yr)}"
    return ""


def _match_year(s: str) -> str:
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", s)
    return m.group(1) if m else ""


def _cover_text(path) -> str:
    """First page (PDF) or first rows (Excel) — cheap read for the cover period."""
    p = Path(path)
    suf = p.suffix.lower()
    if suf == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(p) as pdf:
                return "\n".join((pg.extract_text() or "") for pg in pdf.pages[:1])
        except Exception:
            return ""
    if suf in (".xlsx", ".xls"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            ws = wb.active
            out = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= 15:
                    break
                out.append(" ".join(str(c) for c in row if c not in (None, "")))
            return "\n".join(out)
        except Exception:
            return ""
    return ""


def infer_report_period(path) -> str:
    """Report period from the filename first (clean titles), then the cover page.
    Returns "" when nothing plausible is found. A bare year is only accepted from the
    filename, not the cover (avoids matching a random year in body text)."""
    name = re.sub(r"\s+", " ", Path(path).name.lower())
    p = (_match_quarter(name) or _match_half(name)
         or _match_month(name) or _match_year(name))
    if p:
        return p
    cover = re.sub(r"\s+", " ", _cover_text(path).lower())
    return _match_quarter(cover) or _match_half(cover) or _match_month(cover)


def backfill_missing_dates(records: list, date_field: str,
                           excel_files=None, pdf_files=None, image_files=None,
                           mark: str = "~") -> int:
    """Fill blank ``date_field`` on each record with its source file's report period
    (prefixed ``mark`` for 'circa'). Uses the record's ``_source`` label to pick the
    right file when several are present; falls back to the single detected period.
    Never overwrites a non-empty date. Returns how many were filled."""
    excel_files = excel_files or []
    pdf_files   = pdf_files or []
    image_files = image_files or []

    periods: dict[str, str] = {}
    for kind, files in (("excel", excel_files), ("pdf", pdf_files), ("image", image_files)):
        for i, fp in enumerate(files, 1):
            label = f"{kind}_{i}" if len(files) > 1 else kind
            try:
                per = infer_report_period(fp)
            except Exception:
                per = ""
            if per:
                periods[label] = per
    if not periods:
        return 0
    default = next(iter(periods.values())) if len(periods) == 1 else ""

    n = 0
    for r in records:
        if str(r.get(date_field) or "").strip():
            continue
        per = periods.get(r.get("_source")) or default
        if per:
            r[date_field] = f"{mark}{per}"
            n += 1
    return n
