#!/usr/bin/env python3
"""
pdf_extractor.py
================
4-stage PDF comparable data extraction pipeline.
Shared by scan_input_sales_comps.py, scan_input_rent_comps.py,
and scan_input_land_comps.py.

Stage 1  PAGE DISCOVERY   — extract_text() keyword scan; one entry per page.
                            Robust against unusual PDF layouts.

Stage 2  TABLE EXTRACTION — extract_tables() on each flagged page (original,
                            reliable approach).  Tables are then filtered by
                            column-mapping score: a table with fewer than
                            MIN_MAPPED columns matched to the target schema is
                            discarded.  This removes leasing/land/other-type
                            tables that share a page with the target section.

Stage 3  COLUMN MAPPING   — tools/column_mapper.map_columns() tiered approach:
                            exact → keyword → fuzzy → LLM last resort.
                            Col-map computed in Stage 2 filter is reused here —
                            no duplicate LLM calls.

Stage 4  RECORD ASSEMBLY  — structured rows → dicts via col_map + unit_map;
                            text-only pages → Ollama direct extraction.
"""

import json
import math
import re
import urllib.request
from pathlib import Path

from tools.column_mapper import map_columns


# ─── constants ────────────────────────────────────────────────────────────────

_NAME_KEYS = frozenset({"property_name", "site_name", "building_name"})

# Generic asset-class / aggregate labels.  When a "property name" is essentially
# just one of these words it is a category heading or a summary-table row
# (e.g. an "Investment Activity by Property Type" breakdown), not a real
# property — such rows are dropped.  Matched on the whole name, so genuine names
# that merely contain the word ("Office Tower", "Retail Park") are unaffected.
_CATEGORY_LABELS = frozenset({
    "residential", "commercial", "industrial", "hospitality",
    "retail", "office", "mixed", "mixed use", "mixed/others",
    "mixed others", "others", "logistics", "warehouse",
    "total", "subtotal", "grand total", "sector", "property type",
})


def _is_category_label(name: str) -> bool:
    """True if the name is just an asset-class / aggregate label, not a property."""
    n = re.sub(r"\s+", " ", re.sub(r"[^\w/ ]", " ", name.lower())).strip()
    return n in _CATEGORY_LABELS or n.replace("/", " ").strip() in _CATEGORY_LABELS


# ─── Stage 1: page discovery ─────────────────────────────────────────────────

def find_relevant_pages(pdf_path: str, section_keywords: list,
                        max_pages: int = 60) -> list:
    """
    Stage 1: scan every page with extract_text() for keyword phrases.

    Returns list of dicts:
        {page_num, section_title, matched_keywords, has_table, text_preview}
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber required: pip install pdfplumber")

    # ── Optional embedding tier (reuse the column-mapper fastembed model) ──────
    # Catches transaction-table pages whose heading is a SEMANTIC match but not an
    # exact substring of any keyword (e.g. "Headline Deals" ≈ "Key Transactions").
    # Only applied to pages that actually contain a table, to limit false positives.
    _emb_model, _kw_vecs, _np = None, None, None
    try:
        from tools.column_mapper import (_EMBED_AVAILABLE, _get_embed_model,
                                          _embed as _embed_text)
        import numpy as _np
        if _EMBED_AVAILABLE and section_keywords:
            _emb_model = _get_embed_model()
            _kw_vecs   = _np.array([_embed_text(_emb_model, kw) for kw in section_keywords])
    except Exception as _e:   # fastembed/numpy missing → skip semantic tier
        _emb_model = None
    _EMB_THRESHOLD = 0.60

    matches = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages[:max_pages], 1):
            text = (page.extract_text() or "").strip()
            if not text:
                continue

            text_lower  = text.lower()
            page_tables = page.extract_tables() or []
            has_table   = any(len(t) > 1 and len(t[0]) > 1 for t in page_tables)

            matched = [kw for kw in section_keywords if kw.lower() in text_lower]

            # Keyword not in body text — check inside table cells (some PDFs
            # embed the section title in a merged table header cell).
            if not matched:
                for tbl in page_tables:
                    for row in tbl[:3]:
                        for cell in row:
                            cell_matched = [kw for kw in section_keywords
                                            if kw.lower() in str(cell or "").lower()]
                            if cell_matched:
                                matched = cell_matched
                                break
                        if matched:
                            break
                    if matched:
                        break

            # Embedding tier — only for table-bearing pages with no exact match.
            emb_title, emb_score = "", 0.0
            if not matched and has_table and _emb_model is not None:
                # Collect heading-like lines across the WHOLE page, then embed the
                # most title-like ones first (a table title such as "TABLE 2: Major
                # Transactions" can sit well below the top of the extracted text).
                scored = []
                for ln in text.splitlines():
                    ls = ln.strip()
                    if not (4 <= len(ls) <= 80 and len(ls.split()) >= 2):
                        continue
                    is_title = bool(re.match(r'(?i)^\s*(?:table|exhibit|schedule|'
                                             r'section)\b', ls))
                    scored.append((0 if is_title else 1, ls))
                scored.sort(key=lambda x: x[0])          # explicit titles first
                cand = [s for _, s in scored[:25]]
                for ln in cand:
                    # Strip "TABLE 2:" / "FIGURE 3 —" prefixes and trailing quarter/year
                    # so the comparison focuses on the heading noun phrase itself.
                    clean = re.sub(r'^\s*(?:table|figure|exhibit|chart|appendix)\s*'
                                   r'\d*\s*[:.\-]?\s*', '', ln, flags=re.I)
                    clean = re.sub(r'[,\-–]?\s*(?:q[1-4]\s*)?\b(?:19|20)\d{2}\b.*$', '',
                                   clean, flags=re.I).strip()
                    for variant in {ln, clean}:
                        if not variant:
                            continue
                        try:
                            s = float(_kw_vecs.dot(_embed_text(_emb_model, variant)).max())
                        except Exception:
                            continue
                        if s > emb_score:
                            emb_score, emb_title = s, ln
                if emb_score >= _EMB_THRESHOLD:
                    matched = [f"~embed({emb_score:.2f})"]

            if not matched:
                continue

            section_title = ""
            if emb_title and matched and matched[0].startswith("~embed"):
                section_title = emb_title[:120]
            else:
                for line in text.splitlines():
                    if any(kw.lower() in line.lower() for kw in matched):
                        section_title = line.strip()[:120]
                        break
            if not section_title:
                section_title = matched[0]

            matches.append({
                "page_num":         i,
                "section_title":    section_title,
                "matched_keywords": matched,
                "has_table":        has_table,
                "text_preview":     text[:400],
            })
            print(f"    Page {i:>3}: {section_title[:60]!r}"
                  f"  keywords={matched}  table={has_table}")

    return matches


# ─── Stage 2 + 3: table extraction with schema-relevance filter ──────────────

def _header_quality(tbls: list) -> int:
    return sum(1 for t in tbls if t
               for cell in (t[0] or []) if str(cell or "").strip())


def _is_col_hdr_row(cells: list) -> bool:
    """
    True if the row looks like column headers: every non-empty, non-wrapped
    cell's first word is ALL CAPS and contains at least one letter.

    Catches 'PRICE (S$ Million)' (first word PRICE is caps) while
    rejecting 'Portfolio of three properties...' (first word Portfolio is mixed-case).
    Requires at least 2 non-empty cells so single-cell title rows are excluded.
    """
    ne = [str(c or "").strip() for c in cells
          if str(c or "").strip() and "\n" not in str(c or "")]
    if len(ne) < 2:
        return False
    return all(
        c.split() and c.split()[0].upper() == c.split()[0]
        and any(ch.isalpha() for ch in c.split()[0])
        for c in ne
    )



def _hdr_has_name(headers: list) -> bool:
    joined = " ".join(str(h or "").lower() for h in headers)
    return any(kw in joined for kw in ("property", "building", "asset", "location", "site"))


def _hdr_has_price(headers: list) -> bool:
    joined = " ".join(str(h or "").lower() for h in headers)
    return any(kw in joined for kw in ("price", "consideration", "value", "amount"))


def _merge_h_fragments(tables: list) -> list:
    """
    Merge same-row-count tables that are horizontal fragments of one logical table.

    pdfplumber splits a single PDF table when a column has its own complete
    border box (e.g. a shaded 'PURCHASE PRICE' column). The fragments have
    identical row counts but no single fragment has both a name column and a
    price column. Concatenate their columns row-wise to reconstruct the full table.
    """
    if len(tables) <= 1:
        return tables

    name_idxs  = [i for i, t in enumerate(tables)
                  if _hdr_has_name(t[0]) and not _hdr_has_price(t[0])]
    price_idxs = [i for i, t in enumerate(tables)
                  if _hdr_has_price(t[0]) and not _hdr_has_name(t[0])]

    if not name_idxs or not price_idxs:
        return tables

    merged_set: set = set()
    merges: list    = []
    for ni in name_idxs:
        for pi in price_idxs:
            if ni in merged_set or pi in merged_set:
                continue
            nt, pt = tables[ni], tables[pi]
            if len(nt) == len(pt):
                merged = [nr + pr for nr, pr in zip(nt, pt)]
                merges.append((min(ni, pi), merged))
                merged_set.update([ni, pi])

    if not merges:
        return tables

    result: dict = {i: t for i, t in enumerate(tables) if i not in merged_set}
    for pos, merged in merges:
        result[pos] = merged
        print(f"      [merge] joined name+price fragments → "
              f"{len(merged[0])} columns, {len(merged)-1} data rows")
    return [result[k] for k in sorted(result)]


def _img2table_page_tables(pdf_path: str, page_num: int) -> list:
    """
    img2table + easyocr fallback for a single page.
    Renders the page as an image and detects table structure visually —
    catches grid lines that pdfplumber misses.
    Returns raw tables in the same list-of-lists format as pdfplumber.
    Returns [] if img2table/easyocr are not installed or extraction fails.
    """
    try:
        import fitz as _fitz
    except ImportError:
        try:
            import pdf2image as _pdf2image  # second render option
        except ImportError:
            print("      [img2table] skipped — install pymupdf or pdf2image")
            return []
        _fitz = None

    try:
        from img2table.document import PDF as _I2TPDF
        from img2table.ocr import EasyOCR as _EasyOCR
    except ImportError:
        print("      [img2table] skipped — pip install img2table easyocr")
        return []

    try:
        import io
        ocr = _EasyOCR(lang=["en"], gpu=False)
        # Use img2table's built-in PDF support directly
        doc = _I2TPDF(src=pdf_path, pages=[page_num - 1])  # 0-indexed
        extracted = doc.extract_tables(ocr=ocr, implicit_rows=True, borderless_tables=False)
        # extracted is {page_idx: [Table, ...]}
        page_tables = extracted.get(page_num - 1, [])
        result = []
        for tbl in page_tables:
            df = tbl.df
            if df is None or df.empty:
                continue
            data = [[str(c or "").strip() for c in row]
                    for row in df.values.tolist()]
            # Use column names as header row if they look like real headers
            col_names = [str(c or "").strip() for c in df.columns.tolist()]
            if any(h and not h.startswith("0") for h in col_names):
                data = [col_names] + data
            if len(data) >= 2:
                result.append(data)
        return result
    except Exception as exc:
        print(f"      [img2table] failed: {exc}")
        return []


def _is_caps_label(cell: str) -> bool:
    """All-caps header token: contains a letter and has no lowercase letters."""
    return any(ch.isalpha() for ch in cell) and cell == cell.upper()


def _row_kind(row: list) -> str:
    """
    Classify a raw camelot row as 'blank' | 'title' | 'hdrfrag' | 'data'.

    'hdrfrag' = a header fragment: every non-empty cell is an all-caps label
    with no digits (the line of a multi-line header that 'stream' split out).
    A digit or a mixed-case word anywhere marks the row as data.
    """
    ne = [str(c or "").strip() for c in row if str(c or "").strip()]
    if not ne:
        return "blank"
    has_data = any(
        any(ch.isdigit() for ch in c)
        or (c != c.upper() and any(ch.islower() for ch in c))
        for c in ne
    )
    if len(ne) == 1:
        c = ne[0]
        if has_data or len(c) > 40 or re.match(r"(TABLE|FIGURE|SOURCE|NOTE)\b", c, re.I):
            return "title"
        return "hdrfrag" if _is_caps_label(c) else "data"
    if has_data:
        return "data"
    return "hdrfrag" if all(_is_caps_label(c) for c in ne) else "data"


def _collapse_multirow_header(tbl: list) -> list:
    """
    Reconstruct a header that camelot 'stream' split across several physical
    rows because the source PDF printed the header on multiple lines
    (e.g. Savills GLS tables: 'SUCCESSFUL TENDER / PRICE / (S$ MILLION)').

    Detects the leading run of header-fragment rows (after any title rows) and
    merges them column-wise, top-to-bottom, into one header row. Only fires
    when the fragments are COMPLEMENTARY — together they fill more columns than
    any single fragment row — which is the signature of a split multi-line
    header rather than two genuine all-caps data rows.
    """
    if not tbl or len(tbl) < 3:
        return tbl

    i, n = 0, len(tbl)
    while i < n and _row_kind(tbl[i]) in ("title", "blank"):
        i += 1

    band = []
    while i < n and _row_kind(tbl[i]) == "hdrfrag":
        band.append(tbl[i])
        i += 1

    if len(band) < 2:
        return tbl

    filled_union = {c for r in band for c, v in enumerate(r) if str(v or "").strip()}
    max_single   = max(sum(1 for v in r if str(v or "").strip()) for r in band)
    if len(filled_union) <= max_single:
        return tbl  # not complementary — likely real data rows, leave untouched

    ncols  = max(len(r) for r in band)
    merged = []
    for c in range(ncols):
        parts = [str(r[c]).strip() for r in band
                 if c < len(r) and str(r[c] or "").strip()]
        merged.append(" ".join(parts))

    print(f"      [multirow-hdr] merged {len(band)} header rows → {merged}")
    return [merged] + tbl[i:]


def _is_unit_subtitle_row(row: list) -> bool:
    """
    True if every non-empty cell is only a parenthetical unit annotation, e.g.
    '(SGD million)', '(SGD/psf)'.  This is the second line of a two-line column
    header (a unit subtitle), not a transaction — it should be folded into the
    header rather than parsed as data (Colliers 'Price / (SGD million)' style).
    """
    ne = [str(c or "").strip() for c in row if str(c or "").strip()]
    if not ne:
        return False
    return all(
        re.fullmatch(r"\(.*\)", line.strip())
        for cell in ne for line in cell.split("\n") if line.strip()
    )


def _merge_transaction_cont_rows(headers: list, rows: list) -> list:
    """
    Reassemble transaction rows whose cells wrapped onto several physical rows.

    camelot 'stream' splits a multi-line cell (a long property name, or a long
    seller list) into separate rows.  A real transaction is anchored by the row
    carrying a numeric PRICE; wrapped fragments (rows with text but no price)
    are attached to the NEAREST anchor by row distance, with ties broken toward
    the anchor that still has an empty cell where the fragment has text.  This
    keeps a wrapped name/seller from gluing onto a different, already-complete
    property above or below it.

    After merging, cleans up:
    - Property name: strip leading/embedded row-number digits ("1\\nSouth Beach"
      → "South Beach"; "Citadines Raffles \\n3 Place" → "Citadines Raffles Place")
    - Price cell: keep only the first line before '\\n' when it is numeric
    """
    import re as _re

    hn = lambda h: h.lower().replace('\n', ' ')
    price_col = next(
        (i for i, h in enumerate(headers)
         if 'price' in hn(h) and 'unit' not in hn(h)),
        None
    )
    if price_col is None:
        return rows

    name_col = next(
        (i for i, h in enumerate(headers)
         if any(k in hn(h) for k in
                ('property', 'name', 'building', 'site', 'location', 'asset'))),
        0
    )

    # A price marks a real transaction. Allow a leading approximation symbol
    # ('~490.0', '≈490') so those rows still anchor instead of being treated as
    # fragments.
    _starts_digit = _re.compile(r'^[~≈]?\s*\d')
    _price_of = lambda r: r[price_col].strip() if price_col < len(r) else ''

    # Anchors = rows whose price cell starts with a digit (one real transaction each).
    anchor_idxs = [i for i, r in enumerate(rows) if _starts_digit.match(_price_of(r))]
    if not anchor_idxs:
        return rows  # no price anchors — leave rows untouched

    # collected[anchor] = list of (row_idx, row) pieces that belong to that anchor.
    collected: dict = {a: [(a, rows[a])] for a in anchor_idxs}
    for i, row in enumerate(rows):
        if i in collected:
            continue
        if not any(str(v or '').strip() for v in row):
            continue  # blank row
        frag_cols = [c for c, v in enumerate(row) if str(v or '').strip()]

        # A name fragment beginning with a lowercase letter is the tail of the
        # line above (e.g. 'portfolio' in 'Mapletree Industrial Trust portfolio'),
        # so attach it to the nearest anchor ABOVE rather than letting distance
        # pull it onto the next record below.  Uppercase fragments (which may be
        # the START of the next record's wrapped name) keep the distance rule.
        frag_name = str(row[name_col]).strip() if name_col < len(row) else ''
        above = [a for a in anchor_idxs if a < i]
        if frag_name[:1].islower() and above:
            collected[max(above)].append((i, row))
            continue

        def _key(a):
            dist = abs(a - i)
            # prefer an anchor that still has an empty cell where this fragment
            # has text (a complete record needs no more name/seller text)
            has_room = any(c < len(rows[a]) and not str(rows[a][c] or '').strip()
                           for c in frag_cols)
            return (dist, not has_room)

        collected[min(anchor_idxs, key=_key)].append((i, row))

    # Build one row per anchor; join its pieces top-to-bottom in original order.
    merged: list = []
    for a in anchor_idxs:
        pieces  = sorted(collected[a], key=lambda p: p[0])
        ncols   = max(len(p[1]) for p in pieces)
        out_row = [''] * ncols
        for _, piece in pieces:
            for ci, val in enumerate(piece):
                v = str(val or '').strip()
                if v:
                    out_row[ci] = (out_row[ci] + ' ' + v).strip() if out_row[ci] else v
        merged.append(out_row)

    # Clean up cell values
    for pos, row in enumerate(merged, 1):
        # Property name: strip a leading row-INDEX only — a number that equals
        # this row's position in the table (the sequential 1,2,3… that some
        # reports print before each deal). A leading number that does NOT match
        # the position is a real street number ("78 Shenton Way", "21 Carpenter")
        # and is kept.
        if name_col < len(row):
            p = _re.sub(r'\n\d+\s*', ' ', row[name_col])  # "Raffles \n3 Place" → "Raffles  Place"
            p = _re.sub(rf'^0*{pos}[\s\n]+', '', p)       # "1\nSouth Beach"/"2 Mapletree" → name
            row[name_col] = ' '.join(p.split())
        # Price: keep only the first numeric line (drop overflowed unit-price text)
        if price_col < len(row):
            pv = row[price_col]
            if '\n' in pv:
                first_line = pv.split('\n')[0].strip()
                if _re.search(r'[\d,]+\.?\d*', first_line):
                    row[price_col] = first_line
            # Pull out the numeric price, ignoring a leading '~'/'≈' or trailing
            # footnote markers ('1,231.7**', '~490.0').
            m = _re.search(r'[\d,]+\.?\d*', row[price_col])
            if m:
                row[price_col] = m.group(0)

    return merged


def _camelot_raw_tables(pdf_path: str, page_num: int) -> list:
    """
    Unified three-phase table extraction:

    Phase 1 — camelot lattice (whole page):
        Uses actual PDF border lines. Precise — won't misread multi-column
        article text as table columns. Works for PDFs with real table borders.
        If complete tables (header + data rows) are found, return immediately.

    Phase 2 — pdfplumber column hints + camelot stream (per section):
        For PDFs where only colored/teal header rows have border lines (e.g.
        alternating-colour industrial tables). pdfplumber finds the header bbox
        and column x-positions; camelot stream then extracts ALL rows in that
        section (including white-background rows that lattice misses).

    Phase 3 — camelot stream (whole page):
        Last resort. May pick up multi-column text as false tables on complex
        page layouts, but better than returning nothing.
    """
    try:
        import camelot
    except ImportError:
        raise ImportError("camelot-py required: pip install 'camelot-py[cv]'")

    result: list = []

    # ── Phase 1: lattice — precise border-line detection ──────────────────────
    try:
        lat_list = camelot.read_pdf(pdf_path, flavor="lattice", pages=str(page_num))
        complete = []
        for tbl in lat_list:
            df = tbl.df
            if df is None or df.empty:
                continue
            rows = [[str(v or "").strip() for v in row] for row in df.values.tolist()]
            if len(rows) <= 1:
                continue
            first_ne = [h for h in rows[0] if h]
            # Skip bordered boxes containing paragraph text (quotes, bios, etc.)
            # — their cells contain line-wrapped text (\n) or are very long sentences
            if any('\n' in h for h in first_ne):
                continue
            if first_ne and sum(len(h) for h in first_ne) / len(first_ne) > 60:
                continue
            complete.append(rows)
        if complete:
            print(f"      [camelot] lattice: {len(complete)} complete table(s)")
            return complete
    except Exception as exc:
        print(f"      [camelot] lattice failed: {exc}")

    # ── Phase 2: pdfplumber column hints + camelot stream per section ─────────
    _STRICT = {
        "vertical_strategy":   "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance":      5,
        "join_tolerance":      5,
        "min_words_vertical":  1,
    }
    sections: list = []
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num - 1]
            ph   = page.height

            tbl_objs = sorted(page.find_tables(_STRICT), key=lambda o: o.bbox[1])
            hdr_objs = []
            for obj in tbl_objs:
                try:
                    data = obj.extract()
                    if data and _is_col_hdr_row(data[0]):
                        hdr_objs.append(obj)
                except Exception:
                    pass

            for idx, obj in enumerate(hdr_objs):
                x0, top, x1, bot = obj.bbox
                next_top = (hdr_objs[idx + 1].bbox[1]
                            if idx + 1 < len(hdr_objs) else ph)
                try:
                    cells    = sorted([c for c in obj.rows[0].cells if c],
                                      key=lambda c: c[0])
                    col_seps = [c[2] for c in cells[:-1]]
                    right_x  = cells[-1][2]
                    left_x   = cells[0][0]
                except Exception:
                    col_seps, right_x, left_x = [], x1, x0

                # table_regions: "x1,y1,x2,y2" in PDF coords (y=0 at bottom)
                region = (f"{left_x:.0f},{ph - next_top:.0f},"
                          f"{right_x:.0f},{ph - top:.0f}")
                cols_s = ",".join(f"{x:.0f}" for x in col_seps)
                sections.append((region, cols_s))
                print(f"      [camelot] section {idx+1}: region={region}")

    except Exception as exc:
        print(f"      [camelot] pdfplumber phase: {exc}")

    if sections:
        for region, cols_s in sections:
            try:
                kwargs: dict = dict(flavor="stream", pages=str(page_num),
                                    table_regions=[region], row_tol=2)
                if cols_s:
                    kwargs["columns"] = [cols_s]
                for tbl in camelot.read_pdf(pdf_path, **kwargs):
                    df = tbl.df
                    if df is None or df.empty:
                        continue
                    rows = [[str(v or "").strip() for v in row]
                            for row in df.values.tolist()]
                    if rows:
                        result.append(rows)
            except Exception as exc:
                print(f"      [camelot] section failed: {exc}")
        if result:
            return result

    # ── Phase 2.5: text-anchor detection ──────────────────────────────────────
    # Match the FULL anchor phrase word-by-word so we land on the actual section
    # title, not a coincidental word in article text. Use the anchor's x-position
    # as the left boundary of the region to exclude side-bar content (bios, etc.)
    _ANCHORS = [
        "KEY SALES TRANSACTIONS", "SIGNIFICANT PRIVATE TRANSACTIONS",
        "PRIVATE TRANSACTIONS", "KEY TRANSACTIONS", "NOTABLE TRANSACTIONS",
        "MAJOR TRANSACTIONS", "KEY DEALS", "SIGNIFICANT DEALS",
    ]
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num - 1]
            ph, pw = page.height, page.width
            words = page.extract_words(x_tolerance=3, y_tolerance=3)

            anchor_y, anchor_x0 = None, None
            for anchor in _ANCHORS:
                aw = anchor.split()
                n  = len(aw)
                for i in range(len(words) - n + 1):
                    if all(words[i+j]['text'].upper() == aw[j].upper()
                           for j in range(n)):
                        anchor_y  = words[i]['top']
                        anchor_x0 = words[i]['x0']
                        break
                if anchor_y is not None:
                    break

            if anchor_y is not None:
                region_x0 = max(0, anchor_x0 - 10)

                # Determine the table's right boundary from column-header words.
                # This excludes charts/sidebars that share the same horizontal band.
                _COL_HDR_KWS = {'Seller', 'Buyer', 'Property', 'Price',
                                 'Tenant', 'Date', 'Area', 'Floor', 'Size'}
                col_hdr_wds = sorted(
                    [w for w in words
                     if w['text'] in _COL_HDR_KWS
                     and w['x0'] >= region_x0
                     and abs(w['top'] - anchor_y) <= 50],
                    key=lambda w: w['x0']
                )
                if col_hdr_wds:
                    tbl_right = min(
                        max(w['x1'] for w in col_hdr_wds) + 30, pw)
                    col_seps  = ",".join(
                        f"{w['x0']:.0f}" for w in col_hdr_wds[1:])
                else:
                    tbl_right, col_seps = pw, ""

                camelot_y_top = ph - max(0, anchor_y - 20)
                # camelot table_regions: "x1,y1,x2,y2" — y1 is TOP (larger PDF y)
                region = (f"{region_x0:.0f},{camelot_y_top:.0f},"
                          f"{tbl_right:.0f},0")
                print(f"      [camelot] text-anchor y={anchor_y:.0f} "
                      f"x={anchor_x0:.0f} tbl_right={tbl_right:.0f} "
                      f"→ region={region}")

                for flavor in ["lattice", "stream"]:
                    try:
                        kwargs: dict = dict(pages=str(page_num),
                                            table_regions=[region])
                        if flavor == "stream":
                            kwargs["row_tol"] = 2
                            if col_seps:
                                kwargs["columns"] = [col_seps]
                        found = []
                        for tbl in camelot.read_pdf(pdf_path, flavor=flavor,
                                                    **kwargs):
                            df = tbl.df
                            if df is None or df.empty:
                                continue
                            rows = [[str(v or "").strip() for v in row]
                                    for row in df.values.tolist()]
                            if len(rows) <= 1:
                                continue
                            if not any(c for c in rows[0]):
                                continue
                            found.append(rows)
                        if found:
                            result.extend(found)
                            print(f"      [camelot] text-anchor {flavor}: "
                                  f"{len(found)} table(s)")
                            break
                    except Exception as exc:
                        print(f"      [camelot] text-anchor {flavor} failed:"
                              f" {exc}")

                if not result:
                    # pdfplumber text+text fallback — reuse already-open page
                    # handle and the same tbl_right boundary.
                    try:
                        crop = page.crop((region_x0,
                                          max(0, anchor_y - 10),
                                          tbl_right, page.height))
                        print(f"      [pdfplumber] x-bounded crop "
                              f"x=[{region_x0:.0f},{tbl_right:.0f}]")
                        pdp_tbls = crop.extract_tables({
                            "vertical_strategy":    "text",
                            "horizontal_strategy":  "text",
                            "snap_tolerance":       3,
                            "join_tolerance":       3,
                            "min_words_horizontal": 1,
                        }) or []
                        for tbl in pdp_tbls:
                            rows = [[str(c or "").strip() for c in row]
                                    for row in tbl]
                            if len(rows) > 1:
                                result.append(rows)
                        if result:
                            print(f"      [pdfplumber] text+text: "
                                  f"{len(result)} table(s)")
                    except Exception as exc:
                        print(f"      [pdfplumber] text+text failed: {exc}")

                if result:
                    return result
    except Exception as exc:
        print(f"      [camelot] text-anchor phase failed: {exc}")

    # ── Phase 3: whole-page stream fallback ───────────────────────────────────
    print(f"      [camelot] page {page_num}: whole-page stream fallback")
    try:
        for tbl in camelot.read_pdf(pdf_path, flavor="stream",
                                    pages=str(page_num), row_tol=2):
            df = tbl.df
            if df is None or df.empty:
                continue
            rows = [[str(v or "").strip() for v in row] for row in df.values.tolist()]
            if rows:
                result.append(rows)
    except Exception as exc:
        print(f"      [camelot] stream fallback failed: {exc}")

    return result


def _is_prose_table(tbl: list) -> bool:
    """
    True if a camelot 'table' is really article prose (multi-column body text)
    rather than a data table — e.g. the commentary paragraphs Savills prints
    beside its transaction tables.

    A real comp table has at least one mostly-numeric column (price/area); a
    prose blob has none and its cells are mostly long sentences.
    """
    if not tbl or len(tbl) < 2:
        return False
    ncols = max(len(r) for r in tbl)
    for c in range(ncols):
        col = [str(r[c]).strip() for r in tbl if c < len(r) and str(r[c] or "").strip()]
        if col and sum(1 for v in col if re.match(r"^[\d,]+\.?\d*$", v)) / len(col) >= 0.5:
            return False  # has a numeric column → structured table, not prose

    cells = [str(c or "").strip() for row in tbl for c in row if str(c or "").strip()]
    if len(cells) < 4:
        return False
    sentence = sum(1 for c in cells if len(c) > 25 and len(c.split()) >= 5)
    return sentence / len(cells) >= 0.4


def _pdfplumber_line_tables(pdf_path: str, page_num: int) -> list:
    """
    Line/border-based table extraction via pdfplumber — fallback for pages where
    camelot's stream drowns a real (bordered / coloured-header) table in the
    surrounding article text. Returns list-of-lists like _camelot_raw_tables.
    """
    try:
        import pdfplumber
    except ImportError:
        return []
    out: list = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num - 1]
            for strat in ({"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                          {"vertical_strategy": "lines", "horizontal_strategy": "text"}):
                for t in (page.extract_tables(strat) or []):
                    rows = [[str(c or "").strip() for c in r]
                            for r in t if any(str(c or "").strip() for c in r)]
                    if len(rows) >= 2 and len(rows[0]) >= 3:
                        out.append(rows)
                if out:
                    break
    except Exception as exc:
        print(f"      [pdfplumber-lines] failed: {exc}")
    return out


def _has_header_row(tbl: list) -> bool:
    """
    True if any of the table's first rows looks like a column-header row:
    >=2 non-empty cells, none containing a digit (headers hold labels, not
    dates/prices/areas), and no cell is a long sentence.  Used to decide when
    camelot dropped the header (or returned prose) and pdfplumber should retry.
    """
    if not tbl:
        return False
    for row in _collapse_multirow_header(tbl)[:3]:
        ne = [str(c or "").replace("\n", " ").strip()
              for c in row if str(c or "").strip()]
        if len(ne) >= 2 and all(not re.search(r"\d", c) and len(c.split()) <= 6
                                for c in ne):
            return True
    return False


def _table_prices(tbl: list) -> set:
    """
    Set of price-like tokens in a table — numbers carrying a comma or a decimal
    (e.g. '1,133.0', '152.8', '1,034,600'), commas stripped.  Used to tell
    whether two extractors found the same set of deals.
    """
    out = set()
    for row in tbl:
        for c in row:
            v = str(c or "").strip().lstrip("~≈ ").rstrip("*†‡ ")
            if re.fullmatch(r"[\d,]+(\.\d+)?", v) and ("," in v or "." in v):
                out.add(v.replace(",", ""))
    return out


def _map_cols(headers: list, rows: list, field_schema: list,
              llm_cfg: dict) -> tuple:
    """Tiered column mapping via tools/column_mapper.map_columns()."""
    col_to_key = {col: key for col, key, _ in field_schema}
    ocfg = llm_cfg.get("ollama", {}) if llm_cfg else {}
    return map_columns(
        headers=headers,
        sample_rows=rows[:3],
        output_fields=field_schema,
        col_to_key=col_to_key,
        base_url=ocfg.get("base_url", "http://localhost:11434"),
        model=ocfg.get("model", "qwen2.5:3b"),
        llm_cfg=llm_cfg,
    )


def extract_page_tables(pdf_path: str, page_infos: list,
                        field_schema: list = None,
                        base_url: str = "", model: str = "",
                        llm_cfg: dict = None,
                        reject_table_headers: list = None) -> list:
    """
    Stage 2: extract tables from each flagged page.

    Tables are always preferred.  Text fallback is used only when the page
    has no grid-line tables at all (has_table=False from Stage 1).

    Returns list of dicts:
        source="table" → {page_num, section_title, headers, rows}
        source="text"  → {page_num, section_title, raw_text}
    """
    results = []

    for info in page_infos:
        page_num  = info["page_num"]
        sec_title = info.get("section_title", "")

        raw_tbls = _camelot_raw_tables(pdf_path, page_num)
        raw_tbls = _merge_h_fragments(raw_tbls)

        # Decide whether to fall back to pdfplumber's line-based extraction, which
        # keeps multi-line cells intact. Two cases, both only adopt pdfplumber
        # tables that have a real header and aren't prose:
        #   (a) camelot produced NO table with a real header — it returned prose,
        #       or grabbed the data but dropped the header row.
        #   (b) camelot HAS a header but FRAGMENTED a multi-line table (split a
        #       long name across rows). Detected when pdfplumber recovers the same
        #       deals (price set) in fewer rows — i.e. it consolidated the cells.
        if raw_tbls and not any(_has_header_row(t) for t in raw_tbls):
            _pp = [t for t in _merge_h_fragments(_pdfplumber_line_tables(pdf_path, page_num))
                   if _has_header_row(t) and not _is_prose_table(t)]
            if _pp:
                print(f"    Page {page_num:>3}: camelot lacked a usable header — "
                      f"using {len(_pp)} pdfplumber line table(s)")
                raw_tbls = _pp
        elif raw_tbls:
            _cam_real   = [t for t in raw_tbls if not _is_prose_table(t)]
            _cam_prices = set().union(*(_table_prices(t) for t in _cam_real)) if _cam_real else set()
            if _cam_prices:
                _pp = [t for t in _merge_h_fragments(_pdfplumber_line_tables(pdf_path, page_num))
                       if _has_header_row(t) and not _is_prose_table(t)]
                if _pp:
                    _pp_prices = set().union(*(_table_prices(t) for t in _pp))
                    if (_cam_prices.issubset(_pp_prices)
                            and sum(len(t) for t in _pp) < sum(len(t) for t in _cam_real)):
                        print(f"    Page {page_num:>3}: camelot fragmented "
                              f"({sum(len(t) for t in _cam_real)} rows) — using "
                              f"{len(_pp)} pdfplumber line table(s) "
                              f"({sum(len(t) for t in _pp)} rows)")
                        raw_tbls = _pp

        print(f"    Page {page_num:>3}: {len(raw_tbls)} table(s) found")

        found_any = False
        _orphaned_hdr: list = []

        for tbl_idx, tbl in enumerate(raw_tbls):
            if not tbl:
                continue

            tbl = _collapse_multirow_header(tbl)

            headers = [str(c or "").strip() for c in tbl[0]]
            rows    = [
                [str(c or "").strip() for c in row]
                for row in tbl[1:]
                if any(c not in (None, "", " ") for c in row)
            ]
            if rows and not any(headers):
                headers, rows = rows[0], rows[1:]
                print(f"      table {tbl_idx+1}: promoted row 1 as headers")

            if not rows:
                _ne     = [h for h in headers if h]
                _has_nl = any('\n' in h for h in _ne)
                _is_col_hdr = not _has_nl and _is_col_hdr_row(headers)

                if (not _is_col_hdr) and _orphaned_hdr:
                    print(f"      table {tbl_idx+1}: single data row — using saved column headers")
                    rows    = [headers]
                    headers = _orphaned_hdr
                elif _is_col_hdr:
                    _orphaned_hdr = headers
                    print(f"      table {tbl_idx+1}: header-only fragment saved — {_ne[:4]}")
                    continue
                else:
                    print(f"      table {tbl_idx+1}: skipped — 0 data rows, headers={_ne[:4]}")
                    continue

            _nonempty = [h for h in headers if h]
            if len(headers) >= 3 and len(_nonempty) == 1 and rows:
                print(f"      table {tbl_idx+1}: title row ({_nonempty[0][:60]!r}) — promoting next row as headers")
                headers = rows[0]
                rows    = rows[1:]

            _nonempty = [h for h in headers if h]
            if any('\n' in h for h in _nonempty):
                _flat = [h.replace('\n', ' ').strip() for h in headers]
                if _is_col_hdr_row(_flat):
                    # genuine multi-line column header (e.g. pdfplumber line table
                    # 'TRANSACTION\nDATE' / 'PRICE\n(S$ MILLION)') — flatten the
                    # newlines and keep it as the header rather than discarding.
                    headers = _flat
                    print(f"      table {tbl_idx+1}: flattened multi-line header → {headers[:4]}")
                elif _orphaned_hdr:
                    print(f"      table {tbl_idx+1}: data-as-header — restoring saved column headers")
                    rows    = [headers] + rows
                    headers = _orphaned_hdr
                else:
                    print(f"      table {tbl_idx+1}: data-as-header — no column context, skipping")
                    continue

            if not rows:
                continue

            _short_frac = sum(1 for h in headers if len(h) < 3) / max(len(headers), 1)
            if _short_frac > 0.4 and rows:
                _merged = [f"{h} {r}".strip() for h, r in zip(headers, rows[0])]
                if (sum(1 for h in _merged if len(h) >= 3) >
                        sum(1 for h in headers if len(h) >= 3)
                        and not any('\n' in h for h in _merged)):
                    headers = _merged
                    rows = rows[1:]
                    print(f"      table {tbl_idx+1}: merged header rows → {headers[:4]}")

            # Re-check after possible header merge — merged sub-headers can contain \n
            _nonempty = [h for h in headers if h]
            if any('\n' in h for h in _nonempty):
                print(f"      table {tbl_idx+1}: skipped — headers contain newlines after merge")
                continue
            if not rows:
                continue

            # Fold a unit-subtitle row (e.g. '(SGD million)' / '(SGD/psf)') into
            # the header so it doesn't pollute the first transaction's cells.
            if rows and _is_unit_subtitle_row(rows[0]):
                sub = rows[0]
                # Take only the FIRST unit line per cell: when camelot merges two
                # columns, the subtitle cell holds several units (e.g.
                # '(SGD million)\n(SGD/psf)') — only the first belongs to this
                # column, and piling both on would confuse field mapping.
                headers = [
                    (h + " " + str(sub[j] or "").split("\n")[0].strip()).strip()
                    if j < len(sub) and str(sub[j] or "").strip() else h
                    for j, h in enumerate(headers)
                ]
                rows = rows[1:]
                print(f"      table {tbl_idx+1}: folded unit-subtitle row into header → {headers[:4]}")
            if not rows:
                continue

            rows = _merge_transaction_cont_rows(headers, rows)
            if not rows:
                continue

            # Reject tables that belong to a different comp type (e.g. GLS / land
            # tender tables in an asset-sales run). Caller passes header keywords
            # that mark a table as out-of-scope.
            if reject_table_headers:
                _hjoin = " ".join(str(h or "") for h in headers).lower()
                if any(kw in _hjoin for kw in reject_table_headers):
                    print(f"      table {tbl_idx+1}: skipped — out-of-scope table "
                          f"(header matched reject list): {headers[:4]}")
                    continue

            print(f"      table {tbl_idx+1}: {len(rows)} data rows, headers={headers[:4]}")
            results.append({
                "page_num":      page_num,
                "section_title": sec_title,
                "headers":       headers,
                "rows":          rows,
                "source":        "table",
            })
            found_any = True

        if not found_any:
            print(f"    Page {page_num:>3}: camelot found no tables — trying img2table")
            ocr_tbls = _img2table_page_tables(pdf_path, page_num)
            for tbl_idx, tbl in enumerate(ocr_tbls):
                if not tbl or len(tbl) < 2:
                    continue
                headers = [str(c or "").strip() for c in tbl[0]]
                rows    = [
                    [str(c or "").strip() for c in row]
                    for row in tbl[1:]
                    if any(c not in (None, "", " ") for c in row)
                ]
                if not rows:
                    continue
                print(f"      [img2table] table {tbl_idx+1}: {len(rows)} rows, "
                      f"headers={headers[:4]}")
                results.append({
                    "page_num":      page_num,
                    "section_title": sec_title,
                    "headers":       headers,
                    "rows":          rows,
                    "source":        "table",
                })
                found_any = True

        if not found_any:
            text = ""
            try:
                import pdfplumber
                with pdfplumber.open(pdf_path) as _pdf:
                    text = (_pdf.pages[page_num - 1].extract_text() or "").strip()
            except Exception:
                pass
            if text:
                results.append({
                    "page_num":      page_num,
                    "section_title": sec_title,
                    "raw_text":      text,
                    "source":        "text",
                })
                print(f"    Page {page_num:>3}: no usable tables — text extraction")

    return results


# ─── GPT-4o vision: full PDF → images → single request ───────────────────────

def _gpt_extract_full_pdf(
    pdf_path: str,
    section_keywords: list,
    field_schema: list,
    subj_tokens: set,
    llm_cfg: dict,
    max_pages: int = 60,
    batch_size: int = 8,
) -> list:
    """
    GPT-4o vision path — replaces Stages 1-4 entirely for OpenAI provider.

    Renders pages as images and sends them to GPT-4o **in batches** (default 8
    pages/request) rather than one giant request. Batching (a) keeps each request
    within size/token limits so large PDFs don't silently fail — the failure mode
    that made a 2nd/large PDF return nothing — and (b) lets us send higher-detail
    images so borderless / dense tables are actually read. Rows from all batches
    are aggregated (dedup happens in the caller).
    """
    try:
        import fitz
    except ImportError:
        raise ImportError("pymupdf required for GPT vision path: pip install pymupdf")

    import base64
    from tools.llm_client import openai_chat as _openai_chat

    doc     = fitz.open(pdf_path)
    n_pages = min(len(doc), max_pages)

    field_list  = "\n".join(f'  "{k}": {d}' for _, k, d in field_schema)
    keyword_str = ", ".join(f'"{kw}"' for kw in section_keywords[:10])

    system = (
        "You are a real estate data extraction specialist. "
        "Your skill: given PDF page images, locate every comparable transaction table, "
        "read all rows accurately, and return the data as a structured JSON array. "
        "Rules:\n"
        "- Search EVERY page — there may be multiple tables across different pages.\n"
        "- Extract ALL rows from ALL tables found.\n"
        "- Output ONLY a valid JSON array. No markdown, no explanation, no extra text.\n"
        "- One JSON object per transaction row.\n"
        "- Copy property names exactly as printed — do not paraphrase or abbreviate."
    )

    n_batches = max(1, (n_pages + batch_size - 1) // batch_size)
    print(f"  [GPT-4o] {n_pages} page(s) → {n_batches} batch(es) of ≤{batch_size} "
          f"at detail='high' ...")

    all_items: list = []
    for b in range(n_batches):
        start = b * batch_size
        end   = min(start + batch_size, n_pages)
        user_content: list = [{"type": "text", "text": (
            f"These are pages {start + 1}-{end} of a real estate report.\n"
            f"Find tables under headings like: {keyword_str}.\n"
            f"Extract every row and return as a JSON array.\n\n"
            f"Fields to extract:\n{field_list}"
        )}]
        for i in range(start, end):
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
            b64 = base64.b64encode(pix.tobytes("png")).decode()
            user_content.append({"type": "text", "text": f"[Page {i + 1}]"})
            user_content.append({"type": "image_url", "image_url": {
                "url": f"data:image/png;base64,{b64}", "detail": "high",
            }})
        messages = [{"role": "system", "content": system},
                    {"role": "user",   "content": user_content}]
        try:
            print(f"  [GPT-4o] batch {b + 1}/{n_batches} — pages {start + 1}-{end} …")
            raw = _openai_chat(llm_cfg, messages, timeout=300)
            raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
            raw = re.sub(r"\n?```$", "", raw)
            m   = re.search(r"\[[\s\S]*\]", raw)
            if not m:
                print(f"  [GPT-4o] batch {b + 1}: no JSON array in response")
                continue
            extracted = json.loads(m.group(0))
            if isinstance(extracted, list):
                all_items.extend(x for x in extracted if isinstance(x, dict))
        except Exception as e:
            print(f"  [GPT-4o] batch {b + 1} failed: {e}")
            continue
    doc.close()

    result = []
    for item in all_items:
        name = str(next((item.get(k, "") for k in _NAME_KEYS if item.get(k)), ""))
        if _skip_subject(name, subj_tokens):
            print(f"      SKIP {name!r:.60s}  — matches subject")
            continue
        price = item.get("price_sgd_m") or item.get("price_psf_gfa") or ""
        print(f"      KEEP {name!r:.55s}  price={price!r:.20s}")
        result.append(item)
    return result


# ─── Stage 4: record assembly ─────────────────────────────────────────────────

def _skip_subject(name: str, subj_tokens: set) -> bool:
    if not subj_tokens or not name:
        return False
    meaningful = {t for t in subj_tokens if not t.isdigit()}
    if not meaningful:
        return False
    row_toks = set(re.sub(r"\W+", " ", name.lower()).split())
    needed   = max(1, math.ceil(len(meaningful) * 0.75))
    return len(row_toks & meaningful) >= needed


def _from_table(headers: list, rows: list, col_map: dict, unit_map: dict,
                subj_tokens: set) -> list:
    """Build record dicts from structured table rows using col_map + unit_map."""
    # Forward-fill the property_name column: PDFs with visually merged/spanning
    # cells produce empty strings in pdfplumber for all rows after the first.
    name_col = col_map.get("property_name")
    if name_col is not None:
        rows = [list(r) for r in rows]
        last_name = ""
        for row in rows:
            if name_col < len(row) and row[name_col]:
                last_name = row[name_col]
            elif name_col < len(row) and not row[name_col] and last_name:
                row[name_col] = last_name

    records = []
    for row in rows:
        rec = {}
        for field_key, col_idx in col_map.items():
            if col_idx is None or col_idx >= len(row):
                continue
            val = row[col_idx]
            if val in ("", None):
                continue
            mult = unit_map.get(field_key, 1.0)
            if mult != 1.0:
                try:
                    val = str(float(re.sub(r"[, ]", "", val)) * mult)
                except (ValueError, AttributeError):
                    pass
            rec[field_key] = val
        if not rec:
            continue
        name = str(next((rec.get(k, "") for k in _NAME_KEYS if rec.get(k)), ""))
        if _is_category_label(name):
            print(f"      SKIP (category label, not a property): {name!r}")
            continue
        if len(name) > 80:
            print(f"      SKIP (garbage — name too long): {name[:60]!r}")
            continue
        if name.count('\n') >= 3:
            print(f"      SKIP (garbage — name has multiple newlines): {name[:60]!r}")
            continue
        price_raw = str(rec.get("price_sgd_m", ""))
        if len(price_raw.split()) >= 4:
            print(f"      SKIP (garbage — price looks like text): {price_raw[:40]!r}")
            continue
        if _skip_subject(name, subj_tokens):
            print(f"      SKIP {name!r:.60s}  — matches subject")
            continue
        price = rec.get("price_sgd_m") or rec.get("price_psf_gfa") or ""
        print(f"      KEEP {name!r:.55s}  price={price!r:.20s}")
        records.append(rec)

    # Tag all records with detected price unit (for downstream display conversion)
    _price_unit_tag = "B" if unit_map.get("price_sgd_m", 1.0) >= 1000 else "M"
    for rec in records:
        rec.setdefault("_price_unit", _price_unit_tag)
    return records


def _ollama_free(base_url: str, model: str, messages: list, timeout: int = 120) -> str:
    """Ollama call without JSON mode — for free-form array output."""
    payload = json.dumps({
        "model": model, "messages": messages,
        "stream": False, "options": {"temperature": 0},
    }).encode()
    req = urllib.request.Request(
        f"{base_url.rstrip('/')}/api/chat", data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())["message"]["content"]


def _from_text(text: str, section_title: str, field_schema: list,
               subj_tokens: set, llm_cfg: dict) -> list:
    """LLM extraction from raw page text — used both as Ollama fallback and GPT primary path."""
    field_list = "\n".join(f'  "{k}": {d}' for _, k, d in field_schema)
    system = (
        "You are a real estate data extraction assistant. "
        "The text below is extracted from a PDF report page and may contain a table. "
        "Identify the table, detect its columns, and extract every data row as a JSON object. "
        "Return ONLY a valid JSON array — no markdown fences, no explanation. "
        "One object per property/transaction row. Use the exact property name as written."
    )
    user = (
        f"Section: {section_title}\n\n"
        f"Fields to extract:\n{field_list}\n\n"
        f"PAGE TEXT:\n---\n{text[:12000]}\n---"
    )
    messages = [
        {"role": "system", "content": system},
        {"role": "user",   "content": user},
    ]
    try:
        provider = (llm_cfg or {}).get("provider", "ollama")
        if provider == "openai":
            from tools.llm_client import openai_chat as _openai_chat
            raw = _openai_chat(llm_cfg, messages)
        else:
            ocfg     = (llm_cfg or {}).get("ollama", {})
            base_url = ocfg.get("base_url", "http://localhost:11434")
            model    = ocfg.get("model",    "qwen2.5:3b")
            raw = _ollama_free(base_url, model, messages)
        raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
        raw = re.sub(r"\n?```$", "", raw)
        m   = re.search(r"\[[\s\S]*\]", raw)
        if not m:
            return []
        extracted = json.loads(m.group(0))
        if not isinstance(extracted, list):
            return []
        result = []
        for item in extracted:
            if not isinstance(item, dict):
                continue
            name = str(next((item.get(k, "") for k in _NAME_KEYS if item.get(k)), ""))
            if _skip_subject(name, subj_tokens):
                continue
            result.append(item)
        return result
    except Exception as e:
        print(f"      [warning] text extraction failed: {e}")
        return []


# ─── Deduplication ────────────────────────────────────────────────────────────

def _norm_price(p: str) -> str:
    """Strip footnote markers and whitespace before price comparison."""
    return re.sub(r"[*†‡\s]", "", str(p or ""))


def _dedup(records: list) -> list:
    """
    Remove duplicates by exact name match OR by truncation match.

    Truncation match: when pdfplumber's loose strategy finds the same table
    as the strict strategy but with cells cut short, the property names differ
    ('MCL Land's ass' vs 'MCL Land's assets').  We catch these by comparing
    the space-stripped first 15 characters of both names when the prices match.
    """
    kept_norms:  list = []   # normalized names of kept records
    kept_prices: list = []   # corresponding price strings
    kept_recs:   list = []   # references to the kept record dicts (parallel to above)
    out: list = []

    for rec in records:
        raw   = str(next((rec.get(k) for k in _NAME_KEYS if rec.get(k)), ""))
        norm  = re.sub(r"\W+", " ", raw.lower()).strip()
        price = str(rec.get("price_sgd_m") or "").strip()

        if not norm:
            out.append(rec)
            continue

        dup_idx, dup_reason = None, None
        for j, (seen_norm, seen_price) in enumerate(zip(kept_norms, kept_prices)):
            if seen_norm == norm:
                dup_idx, dup_reason = j, "exact"
                break
            # Truncation check: normalised price match + compact prefix overlap
            if price and _norm_price(price) == _norm_price(seen_price):
                a = norm.replace(" ", "")
                b = seen_norm.replace(" ", "")
                short, long_ = (a, b) if len(a) <= len(b) else (b, a)
                if len(short) >= 8 and long_.startswith(short[:min(len(short), 15)]):
                    dup_idx, dup_reason = j, "truncated"
                    break

        if dup_reason is not None:
            # A comp whose name wrapped onto several PDF rows is emitted as a
            # name-only fragment (no price) followed by the data row (with price).
            # Merge the duplicate into the kept record — filling any empty fields
            # and adopting a price when the kept copy lacks one — instead of
            # discarding it (which would otherwise drop the priced copy).
            kept = kept_recs[dup_idx]
            for k, v in rec.items():
                if v not in ("", None) and kept.get(k) in ("", None):
                    kept[k] = v
            if not kept_prices[dup_idx] and price:
                kept_prices[dup_idx] = price
            print(f"    [dedup] merged {dup_reason} duplicate: {raw!r:.60s}")
        else:
            kept_norms.append(norm)
            kept_prices.append(price)
            kept_recs.append(rec)
            out.append(rec)

    return out


def _merge_tenant_fragments(rows: list, col_map: dict) -> list:
    """
    Stitch wrapped tenant cells back together (rent/lease tables).

    A long tenant name can wrap onto several PDF lines; the extractor then splits
    it across rows where only the tenant column has text (property/area empty),
    and the data row's own tenant cell may come out blank, e.g.:

        ['', '', 'Tech-Component Resources Pte', '', '']    # wrap line 1
        ['Admirax', 'Sembawang', '', '9,000', 'New Lease']  # data row, tenant blank
        ['', '', 'Ltd', '', '']                             # wrap line 2

    The tenant-only fragments are attached to the owning data row (the nearby data
    row whose tenant is blank), so the comp becomes ONE row with the full tenant —
    which also means deleting that row in the preview removes everything.
    Runs only when a tenant column exists, so sales/land tables are untouched.
    """
    t = col_map.get("tenant")
    name_c = (col_map.get("building_name") or col_map.get("property_name")
              or col_map.get("site_name"))
    anchors = [c for c in (name_c, col_map.get("nla_sf")) if c is not None]
    if t is None or not anchors:
        return rows

    def _cell(r, c):
        return str(r[c]).strip() if (c is not None and c < len(r)) else ""
    def _is_data(r):
        return any(_cell(r, c) for c in anchors)
    def _is_frag(r):
        return (not _is_data(r)) and bool(_cell(r, t))

    rows = [list(r) for r in rows]
    n = len(rows)
    consumed = [False] * n

    # Pass A: a data row with a BLANK tenant absorbs the tenant-only fragment
    # rows immediately above and below it (the wrapped name lines).
    for i in range(n):
        if not _is_data(rows[i]) or _cell(rows[i], t):
            continue
        above, a = [], i - 1
        while a >= 0 and not consumed[a] and _is_frag(rows[a]):
            above.append(a); a -= 1
        above.reverse()
        below, b = [], i + 1
        while b < n and not consumed[b] and _is_frag(rows[b]):
            below.append(b); b += 1
        parts = [_cell(rows[k], t) for k in above + below]
        if parts and t < len(rows[i]):
            rows[i][t] = " ".join(p for p in parts if p)
            for k in above + below:
                consumed[k] = True

    # Pass B: any leftover tenant fragment is a continuation — append to the
    # previous data row's tenant.
    last_data = None
    for i in range(n):
        if consumed[i]:
            continue
        if _is_data(rows[i]):
            last_data = i
        elif _is_frag(rows[i]) and last_data is not None and t < len(rows[last_data]):
            own = _cell(rows[last_data], t)
            rows[last_data][t] = (own + " " + _cell(rows[i], t)).strip()
            consumed[i] = True

    return [rows[i] for i in range(n) if not consumed[i]]


# ─── Stage 3 → 4 orchestration ───────────────────────────────────────────────

def map_to_schema(page_tables: list, field_schema: list,
                  subject_name: str, llm_cfg: dict, dedup: bool = True) -> list:
    """
    Stage 3 + 4: for each page entry build records.

    Tables that already have col_map from Stage 2 (schema-filter pass) reuse
    it directly — no second LLM call.  Text-only pages call Ollama directly.

    dedup : merge duplicate names (default). Set False for rent/lease comps,
            where the same building legitimately appears multiple times (one row
            per lease deal — different tenants/floors/areas).
    """
    subj_tokens = (set(re.sub(r"\W+", " ", subject_name.lower()).split())
                   if subject_name else set())

    has_any_table = any(e["source"] == "table" for e in page_tables)

    # Count tables per page for labelling
    from collections import Counter
    _page_tbl_count: Counter = Counter(
        e["page_num"] for e in page_tables if e["source"] == "table"
    )
    _page_tbl_seen: Counter = Counter()

    all_records: list = []
    for entry in page_tables:
        pg = entry["page_num"]
        if entry["source"] == "text":
            if has_any_table:
                print(f"    Page {pg:>3}: text-only page skipped"
                      f" (table data present on other pages)")
                continue
            recs = _from_text(
                entry["raw_text"], entry["section_title"],
                field_schema, subj_tokens, llm_cfg,
            )
            print(f"    Page {pg:>3}: {len(recs)} record(s) from text")
        else:
            _page_tbl_seen[pg] += 1
            n_of = (f"table {_page_tbl_seen[pg]}/{_page_tbl_count[pg]}"
                    if _page_tbl_count[pg] > 1 else "")
            label = f"Page {pg:>3}" + (f" {n_of}" if n_of else "")
            col_map, unit_map = _map_cols(
                entry["headers"], entry["rows"], field_schema, llm_cfg,
            )
            print(f"      col_map: { {k: (entry['headers'][v] if v is not None and v < len(entry['headers']) else None) for k, v in col_map.items()} }")
            if entry["rows"]:
                print(f"      row[0]: {entry['rows'][0]}")
            _rows = entry["rows"]
            if col_map.get("tenant") is not None:
                _rows = _merge_tenant_fragments(_rows, col_map)
            recs = _from_table(
                entry["headers"], _rows, col_map, unit_map, subj_tokens,
            )
            print(f"    {label}: {len(recs)} record(s)")
        all_records.extend(recs)

    if not dedup:
        return all_records
    deduped = _dedup(all_records)
    removed = len(all_records) - len(deduped)
    if removed:
        print(f"  [PDF] Dedup: removed {removed} duplicate(s) (listed above)")
    return deduped


# ─── PUBLIC API ───────────────────────────────────────────────────────────────

def extract_pdf_records(
    pdf_path: str,
    section_keywords: list,
    field_schema: list,
    llm_cfg: dict,
    subject_name: str = "",
    max_pages: int = 60,
    reject_table_headers: list = None,
    dedup: bool = True,
) -> list:
    """
    Full 4-stage PDF extraction pipeline.  Public API unchanged.

    Parameters
    ----------
    pdf_path         : path to the PDF file
    section_keywords : section heading phrases to search for
    field_schema     : list of (display_name, internal_key, description) tuples
    llm_cfg          : LLM config dict from deal config
    subject_name     : subject property name to exclude from results
    max_pages        : max pages to scan (default 60)

    Returns
    -------
    List of raw dicts {internal_key: raw_string_value}.
    """
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        raise ImportError("pdfplumber required: pip install pdfplumber")

    pdf_name = Path(pdf_path).name

    provider = (llm_cfg or {}).get("provider", "ollama")

    # ── GPT-4o vision path (Stages 1-4 replaced by one multimodal call) ───────
    if provider == "openai":
        print(f"\n  [PDF] GPT-4o vision — scanning full PDF for tables ...")
        subj_tokens = (set(re.sub(r"\W+", " ", subject_name.lower()).split())
                       if subject_name else set())
        records = _gpt_extract_full_pdf(
            pdf_path, section_keywords, field_schema,
            subj_tokens, llm_cfg, max_pages,
        )
        records = _dedup(records) if dedup else records
        print(f"  [PDF] {len(records)} record(s) extracted")
        return records

    # ── Ollama path: 4-stage pdfplumber + rule-based pipeline ─────────────────
    print(f"\n  [PDF Stage 1] Locating relevant pages in {pdf_name!r} ...")
    page_infos = find_relevant_pages(pdf_path, section_keywords, max_pages)
    if not page_infos:
        print(f"  [PDF] No pages matched.  Keywords searched: {section_keywords}")
        return []
    print(f"  [PDF] {len(page_infos)} page(s) matched: "
          f"{[p['page_num'] for p in page_infos]}")

    print(f"\n  [PDF Stage 2] Extracting + filtering tables ...")
    page_tables = extract_page_tables(
        pdf_path, page_infos,
        field_schema=field_schema,
        llm_cfg=llm_cfg,
        reject_table_headers=reject_table_headers,
    )
    if not page_tables:
        print(f"  [PDF] No content found on matched pages.")
        return []

    tbl_count  = sum(1 for t in page_tables if t["source"] == "table")
    txt_count  = sum(1 for t in page_tables if t["source"] == "text")
    total_rows = sum(len(t.get("rows", [])) for t in page_tables)
    print(f"  [PDF] {tbl_count} table(s) ({total_rows} rows), "
          f"{txt_count} text-only page(s)")

    print(f"\n  [PDF Stage 3+4] Assembling records ...")
    records = map_to_schema(page_tables, field_schema, subject_name, llm_cfg, dedup=dedup)
    print(f"  [PDF] {len(records)} record(s) extracted")
    return records
