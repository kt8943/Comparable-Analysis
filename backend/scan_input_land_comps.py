#!/usr/bin/env python3
"""
scan_input_land_comps.py
========================
Reads a user-provided Excel of land sale comparables, uses Ollama to
intelligently map columns and classify records, then produces the
13-column formatted output Excel (company template) and an optional
Mapbox location map.

No web search — all data comes from the input file specified by
``land_input_file`` in the deal config.

Output schema  (13 columns)
---------------------------
  Property | Map Marker | Date of Launch | Land Zoning | Land Tenure (Y)
  Site Area (SF) | Max GFA (SF) | Price (SGD M) | Price (SGD psf ppr)
  Adj. Price (SGD psf ppr) | Location | Quality | Comment

Pipeline
--------
  1  GEOCODE    Geocode the subject property via Mapbox
  2  PARSE      Read input Excel; Ollama auto-detects column mapping
  3  CLASSIFY   Ollama assigns location, quality per comp
  4  GEOCODE    Geocode each comp; sort by distance from subject
  5  RENDER     Formatted 13-column Excel; optional Mapbox map PNG

Usage
-----
    python3 scan_input_land_comps.py --config configs/deal_config_88_Cecil.json
    python3 scan_input_land_comps.py --config configs/deal_config_88_Cecil.json --map

Config keys used
----------------
    land_input_file       : path to the Excel with land sale comp data  (required)
    output_file           : used to derive the output folder
    parameters.max_comps  : maximum comps to include  (default 10)
    parameters.bala_yield : Bala table yield           (default 0.06)
"""

import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).parent))
# ── Corporate proxy TLS fix (trust OS cert store; no-op without truststore) ────
from tools import corp_ssl  # noqa: F401  — must import before any HTTPS call
# ── Windows UTF-8 fix ─────────────────────────────────────────────────────────
for _stream in (_sys.stdout, _sys.stderr):
    try:
        if getattr(_stream, "encoding", "utf-8").lower().replace("-", "") != "utf8":
            _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass

import argparse
import json
import re
import urllib.request
from pathlib import Path

import openpyxl

from generate_land_comps_map import render_map
from generate_comps_map_base import geocode_any as geocode_with_fallbacks, build_geocode_queries as _build_geocode_queries, near_country_centroid as _near_country_centroid
from generate_land_comps_table import (
    get_land_schema, bala_factor,
    subject_to_row, comp_to_row, build_workbook,
)
import generate_global_land_comps_table as _global_land_tbl
from tools.calculations import haversine_km as _haversine_km, parse_num as _num
from tools.json_utils import fix_json as _fix_json, split_json_arrays as _split_json_arrays
from tools.llm_client import ollama_post as _ollama_post, apply_refinement as _apply_refinement
from tools.excel_reader import find_best_sheet as _find_best_sheet, find_header_row as _find_header_row, sheet_keywords as _sheet_keywords, split_tables as _split_tables
from tools.vision_llm import call_vision_llm as _call_vision_llm
from tools.column_mapper import map_columns as _map_cols_tool
from tools.geo_utils import write_geo_sidecar


# ─────────────────────────────────────────────────────────────────────────────
# GENERAL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _parse_tenure_yrs(val):
    """Parse tenure string → integer years. 'Freehold' / 999-yr → 0."""
    if val is None:
        return None
    s = str(val).strip()
    if "freehold" in s.lower():
        return 0
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if m:
        n = float(m.group(1))
        return 0 if n >= 999 else round(n, 1)
    return None


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL PARSING  (Ollama-driven column mapping)
# ─────────────────────────────────────────────────────────────────────────────

# Output column names drive column detection — no hard-coded input synonyms.
# Ollama is shown the OUTPUT column display names and finds matching input columns.
_OUTPUT_FIELDS = [
    # (output_col_display_name,  internal_key,   description_for_ollama)
    ("Property",            "site_name",     "Name of the land site, development, parcel, project, asset, or lot — e.g. 'Property', 'Site', 'Project', 'Asset', 'Lot', 'Location', 'Name'"),
    ("Property Address",    "address",       "Street address or location of the site"),
    ("Date of Launch",      "launch_date",   "Transaction, tender award, tender closing, or launch date"),
    ("Land Zoning",         "land_zoning",   "Planning or land use zone (Commercial, White, B1)"),
    ("Land Tenure (Y)",     "tenure",        "Tenure description: 99-year Leasehold, Freehold, 30-year"),
    ("Site Area (SF)",      "site_area_sf",  "Site or land area in square feet"),
    ("Max GFA (SF)",        "max_gfa_sf",    "Maximum allowable gross floor area in square feet"),
    ("Price (SGD M)",       "price_sgd_m",   "Total sale or tender price in millions (any currency) — e.g. 'Price', 'Land Price', 'Transaction Price', 'Successful Tender Price', 'Tendered Price', 'Winning Bid'"),
    ("Price (SGD psf ppr)", "price_psf_ppr", "Land price per floor area per plot ratio (any currency) — e.g. 'psf ppr', 'per sqm per PR', 'tendered price psf ppr', 'land rate'"),
    ("Comment",             "remarks",       "Remarks, comments, or notes on the transaction"),
    ("Sale Type",           "sale_type",     "Transaction type: GLS Confirmed, GLS Reserve, En Bloc, Private"),
]
_OUTPUT_COL_TO_KEY = {col: key for col, key, _ in _OUTPUT_FIELDS}

_sheet_kws  = _sheet_keywords(_OUTPUT_FIELDS)
_best_sheet = lambda wb: _find_best_sheet(wb, _sheet_kws)


def _map_columns(headers: list, sample_rows: list,
                 base_url: str, model: str, llm_cfg: dict = None) -> tuple:
    return _map_cols_tool(headers, sample_rows, _OUTPUT_FIELDS,
                          _OUTPUT_COL_TO_KEY, base_url, model, llm_cfg=llm_cfg)


def parse_input_excel(input_file: str, base_url: str, model: str,
                      subject_name: str = "", llm_cfg: dict = None,
                      _segment: tuple = None) -> list:
    """
    Read any Excel file of land comps. Uses LLM (GPT or Ollama) to map columns.
    Returns list of dicts with standardised field names.

    A single sheet may hold several STACKED tables with different column layouts;
    each is detected and parsed with its OWN header row (via ``_segment`` recursion).
    """
    if _segment is not None:
        best, headers, data_rows = _segment
        print(f"  Headers: {[h for h in headers if h]}")
    else:
        wb   = openpyxl.load_workbook(input_file, data_only=True)
        best = _best_sheet(wb)
        ws   = wb[best]
        rows = [tuple(c.value for c in row) for row in ws.iter_rows()]

        segments = _split_tables(rows, _sheet_kws)
        if len(segments) > 1:
            print(f"  Sheet {best!r}: {len(segments)} stacked tables detected — "
                  "parsing each with its own header row.")
            _all, _pu0 = [], None
            for _k, (_hi, _sh, _sd) in enumerate(segments, 1):
                print(f"\n  ══ Table {_k}/{len(segments)} — header row {_hi + 1}, "
                      f"{len(_sd)} data row(s) ══")
                _recs, _pu = parse_input_excel(
                    input_file, base_url, model, subject_name=subject_name,
                    llm_cfg=llm_cfg, _segment=(best, _sh, _sd))
                _all += _recs
                if _pu0 is None:
                    _pu0 = _pu
            return _all, _pu0 or "M"

        _hi, headers, data_rows = segments[0]
        print(f"  Sheet: {best!r}  |  Header row: {_hi + 1}  |  Data rows: {len(data_rows)}")
        print(f"  Headers: {[h for h in headers if h]}")

    # Pick sample rows for Ollama: prefer rows with >= 3 non-empty cells
    # (avoids section sub-header rows that only have 1–2 cells filled)
    sample_rows = [r for r in data_rows
                   if sum(1 for c in r if c not in (None, "")) >= 3][:3]
    if not sample_rows:
        sample_rows = data_rows[:3]

    # Tiered column mapping: exact → keyword → fuzzy → Ollama (last resort)
    print(f"  Mapping columns …")
    col_map, unit_map = _map_columns(headers, sample_rows, base_url, model, llm_cfg=llm_cfg)
    _price_unit_detected = "B" if unit_map.get("price_sgd_m", 1.0) >= 1000 else "M"

    # Extract year from sheet name first, then fall back to launch date column header.
    _date_header_year = None
    _sheet_ym = re.search(r"\b(20\d{2})\b", best)
    if _sheet_ym:
        _date_header_year = _sheet_ym.group(1)
        print(f"  [date-year] {_date_header_year!r} extracted from sheet name {best!r}")
    else:
        _date_col_idx = col_map.get("launch_date")
        if _date_col_idx is not None and _date_col_idx < len(headers):
            _col_ym = re.search(r"\b(20\d{2})\b", headers[_date_col_idx])
            if _col_ym:
                _date_header_year = _col_ym.group(1)
                print(f"  [date-year] {_date_header_year!r} extracted from column header {headers[_date_col_idx]!r}")

    def _get(row, field):
        idx = col_map.get(field)
        if idx is not None and idx < len(row):
            return row[idx]
        return None

    def _get_num(row, field):
        val = _num(_get(row, field))
        if val is None:
            return None
        return val * unit_map.get(field, 1.0)

    subj_tokens = (set(re.sub(r"\W+", " ", subject_name.lower()).split())
                   if subject_name else set())

    records = []
    for row in data_rows:
        name = str(_get(row, "site_name") or "").strip()
        if not name:
            continue

        # Skip totals / averages rows
        if re.search(r"\b(total|average|avg|summary)\b", name, re.I):
            continue

        # Skip subject property if it appears in the comps
        if subj_tokens:
            name_tokens = set(re.sub(r"\W+", " ", name.lower()).split())
            if len(name_tokens & subj_tokens) >= max(2, len(subj_tokens) * 0.75):
                continue

        price_m   = _get_num(row, "price_sgd_m")
        # Safety net: if header had no unit marker, detect from value magnitude
        if price_m is not None and price_m > 100_000:
            price_m = round(price_m / 1_000_000, 3)
        price_psf = _get_num(row, "price_psf_ppr")
        site_area = _get_num(row, "site_area_sf")
        max_gfa   = _get_num(row, "max_gfa_sf")

        # Need at least a price value
        if price_m is None and price_psf is None:
            continue

        # Compute price_psf_ppr if missing but price_m + max_gfa are available
        if price_psf is None and price_m is not None and max_gfa:
            price_psf = round(price_m * 1_000_000 / max_gfa)

        tenure_raw = _get(row, "tenure")
        tenure_yrs = _parse_tenure_yrs(tenure_raw)
        addr       = str(_get(row, "address")     or "").strip()
        zoning     = str(_get(row, "land_zoning")  or "").strip()
        launch_dt  = str(_get(row, "launch_date")  or "").strip()
        if launch_dt and _date_header_year and not re.search(r"\b(?:19|20)\d{2}\b", launch_dt):
            launch_dt = f"{launch_dt} {_date_header_year}"
        remarks    = str(_get(row, "remarks")       or "").strip()
        sale_type  = str(_get(row, "sale_type")     or "").strip()

        # Combine sale_type + remarks into a single comment
        parts = []
        if sale_type:
            parts.append(sale_type)
        if remarks and remarks.lower() not in sale_type.lower():
            parts.append(remarks)
        comment = "  ".join(parts)

        records.append({
            "raw_description": f"{name}\n{addr}" if addr else name,
            "property_name":   name,
            "address":         addr,
            "launch_date":     launch_dt,
            "land_zoning":     zoning,
            "tenure_raw":      str(tenure_raw or "").strip(),
            "tenure_yrs":      tenure_yrs,
            "site_area_sf":    int(site_area) if site_area else None,
            "max_gfa_sf":      int(max_gfa)   if max_gfa   else None,
            "price_sgd_m":     price_m,
            "price_psf_ppr":   price_psf,
            "comment":         comment,
            "location":        "",
            "quality":         "",
            "_source":         "excel",
            "_price_unit":     _price_unit_detected,
        })

    return records, _price_unit_detected


# ─────────────────────────────────────────────────────────────────────────────
# PDF INPUT PARSING
# ─────────────────────────────────────────────────────────────────────────────

_PDF_SECTION_KEYWORDS = [
    "Land Sales", "Additional Land Sales", "GLS Sales",
    "Land Bids", "Land Comparables", "Additional Land Comparables",
    "Government Land Sales", "GLS Tenders", "GLS Results",
    "Successful Tenders", "Tender Results", "Sites Awarded", "State Land",
    "Land Transactions", "Additional Land Transactions",
    "En Bloc Sales", "Land Acquisitions",
    "Land Sale Evidence", "Collective Sales",
    # Broker market-report tables listing land/site deals
    "Major Transactions", "Major Deals", "Selected Transactions",
    "Notable Transactions", "Key Transactions",
]


def _parse_pdf_records(pdf_path: str, llm_cfg: dict,
                       subject_name: str = "") -> list:
    """
    Extract land comp records from a PDF using the shared 4-stage
    pdf_extractor pipeline (pdfplumber page discovery → table detection →
    field mapping → record assembly).

    Returns the same record format as parse_input_excel() so downstream
    classification and geocoding work unchanged.
    """
    from pdf_extractor import extract_pdf_records

    # Reject asset / private-sector investment-sales tables that share a land page
    # (e.g. Savills "Top Investment Sales in the Private Sector"). They are building
    # deals, not land — identified by a BUYER/VENDOR column or income/leasing
    # metrics. Land / GLS tables use "SUCCESSFUL TENDERER" instead, so these
    # markers never appear in a genuine land table.
    _ASSET_TABLE_MARKERS = [
        "buyer", "vendor", "purchaser",
        "npi yield", "cap rate", "capitalisation rate",
        "net lettable", "nla", "tenant",
    ]

    # _OUTPUT_FIELDS already includes address, site_name, all land fields
    raw_records = extract_pdf_records(
        pdf_path, _PDF_SECTION_KEYWORDS, _OUTPUT_FIELDS,
        llm_cfg, subject_name=subject_name,
        reject_table_headers=_ASSET_TABLE_MARKERS,
    )
    if not raw_records:
        return []

    subj_tokens = set(re.sub(r"\W+", " ", subject_name.lower()).split()) if subject_name else set()

    records = []
    for item in raw_records:
        if not isinstance(item, dict):
            continue

        name = str(item.get("site_name") or "").strip()
        if not name:
            continue

        # Skip summary/average rows and subject property
        if re.search(r"\b(total|average|avg|summary)\b", name, re.I):
            continue
        if subj_tokens:
            name_tokens = set(re.sub(r"\W+", " ", name.lower()).split())
            if len(name_tokens & subj_tokens) >= max(2, len(subj_tokens) * 0.75):
                continue

        price_m   = _num(item.get("price_sgd_m"))
        if price_m is not None and price_m > 100_000:
            price_m = round(price_m / 1_000_000, 3)
        price_psf = _num(item.get("price_psf_ppr"))
        site_area = _num(item.get("site_area_sf"))
        max_gfa   = _num(item.get("max_gfa_sf"))

        if price_m is None and price_psf is None:
            continue

        # Date sanity: non-empty date with no 4-digit year is a sentence fragment.
        launch_dt = str(item.get("launch_date") or "").strip()
        if launch_dt and not re.search(r"\b(?:19|20)\d{2}\b", launch_dt):
            continue

        # Derive price psf from total price + GFA when not explicitly given
        if price_psf is None and price_m is not None and max_gfa:
            price_psf = round(price_m * 1_000_000 / max_gfa)

        tenure_raw = item.get("tenure")
        tenure_yrs = _parse_tenure_yrs(tenure_raw)
        addr       = str(item.get("address")    or "").strip()
        zoning     = str(item.get("land_zoning") or "").strip()
        remarks    = str(item.get("remarks")     or "").strip()
        sale_type  = str(item.get("sale_type")   or "").strip()

        parts = []
        if sale_type:
            parts.append(sale_type)
        if remarks and remarks.lower() not in sale_type.lower():
            parts.append(remarks)
        comment = "  ".join(parts)

        records.append({
            "raw_description": f"{name}\n{addr}" if addr else name,
            "property_name":   name,
            "address":         addr,
            "launch_date":     launch_dt,
            "land_zoning":     zoning,
            "tenure_raw":      str(tenure_raw or "").strip(),
            "tenure_yrs":      tenure_yrs,
            "site_area_sf":    int(site_area) if site_area else None,
            "max_gfa_sf":      int(max_gfa)   if max_gfa   else None,
            "price_sgd_m":     price_m,
            "price_psf_ppr":   price_psf,
            "comment":         comment,
            "location":        "",
            "quality":         "",
            "_source":         "pdf",
        })

    print(f"  → {len(records)} valid records after filtering")
    return records


# ─────────────────────────────────────────────────────────────────────────────
# IMAGE / SCREENSHOT INPUT PARSING
# ─────────────────────────────────────────────────────────────────────────────

# Ollama model names that support vision (image input).
# If the configured model is not in this list, image parsing requires OpenAI.
_IMAGE_EXTRACT_PROMPT = """\
You are extracting real estate land sale comparable data from a table screenshot.

IMPORTANT: Count every data row visible in the table and extract ALL of them — do not stop after the first row.

Extract ALL rows from the table in the image. For each row return these fields:
{field_list}

Rules:
- Return ONLY a single valid JSON array containing ALL rows. No preamble, no explanation, no markdown fences.
- ALL rows must be in ONE array: [{{}}, {{}}, ...] — do NOT return separate arrays per row.
- If a field is not present for a row, use null.
- Numbers as numbers (not strings). Dates as strings.
- Do not invent or estimate values — only extract what is visible in the image.
- Skip header rows and average/total rows.
- Price ranges like "600-630" should be returned as the string "600-630" (not split).
- SCOPE — this is a LAND SALES / GLS / SITE TENDER analysis. Extract ONLY land / site /
  government-land-tender transactions. Do NOT extract:
    * ASSET / INVESTMENT SALES of standing buildings (e.g. columns Buyer, Vendor,
      Purchaser, NPI/Cap Rate, NLA, Tenant), and
    * RENTAL / LEASING tables (e.g. columns Rent, $ psf/month, Lease, Tenant, Occupancy).
  If a table is an asset-sales or leasing table, ignore it entirely. If the image has no
  land / GLS table, return an empty array [].
"""


def _parse_image_records(image_path: str, llm_cfg: dict, openai_key: str = "",
                         subject_name: str = "") -> list:
    """
    Extract land comp records from a table screenshot using a vision LLM.
    Returns the same record format as parse_input_excel().
    """
    print(f"  Reading image: {Path(image_path).name} ...")

    field_list = "\n".join(
        f'  "{key}": {desc}'
        for _, key, desc in _OUTPUT_FIELDS
    ) + '\n  "address": Street address of the land site'

    prompt = _IMAGE_EXTRACT_PROMPT.format(field_list=field_list)

    print(f"  Extracting via vision LLM ...")
    try:
        raw = _call_vision_llm(image_path, prompt, llm_cfg, openai_key)
    except Exception as e:
        print(f"  [warning] Vision LLM call failed: {e}")
        return []

    print(f"  [debug]   Raw vision response ({len(raw)} chars):\n{raw[:1500]}")

    # Parse JSON array from response
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
    if m:
        raw = m.group(1).strip()
    m2 = re.search(r"\[[\s\S]*\]", raw)
    if not m2:
        print("  [warning] No JSON array found in vision LLM response.")
        print(f"  [debug]   Raw response: {raw[:400]}")
        return []
    raw_json = m2.group(0)

    def _try_parse(s: str):
        """Try JSON → fixed JSON → Python literal eval, return list or None."""
        import ast as _ast
        # Pre-convert JSON keywords → Python equivalents for ast.literal_eval.
        # minicpm-v often mixes single-quoted strings (Python) with null/true/false
        # (JSON), so neither parser handles it as-is without this normalisation.
        _py = re.sub(r'\bnull\b', 'None',
               re.sub(r'\btrue\b', 'True',
               re.sub(r'\bfalse\b', 'False', s)))
        for attempt in (
            lambda: json.loads(s),
            lambda: json.loads(_fix_json(s)),
            lambda: _ast.literal_eval(_py),  # single-quoted strings + null/true/false
        ):
            try:
                r = attempt()
                if isinstance(r, list):
                    return r
            except Exception:
                pass
        return None

    extracted = _try_parse(raw_json)
    if extracted is None:
        # Last resort — vision LLM may have returned one array per record on
        # separate lines ("Extra data" error).  Find every top-level array and merge.
        all_arrays = _split_json_arrays(raw)
        if len(all_arrays) > 1:
            print(f"  [info]    Found {len(all_arrays)} separate JSON arrays — merging …")
            merged = []
            for arr_str in all_arrays:
                arr = _try_parse(arr_str)
                if arr:
                    merged.extend(arr)
            if merged:
                extracted = merged
            else:
                print(f"  [warning] Could not parse vision LLM JSON (all strategies failed).")
                print(f"  [debug]   Raw response (first 500 chars):\n{raw[:500]}")
                return []
        else:
            print(f"  [warning] Could not parse vision LLM JSON (all strategies failed).")
            print(f"  [debug]   Raw response (first 500 chars):\n{raw[:500]}")
            return []
    if not isinstance(extracted, list):
        print("  [warning] Vision LLM did not return a JSON array.")
        return []

    print(f"  → {len(extracted)} raw records extracted from image")

    subj_tokens = set(re.sub(r"\W+", " ", subject_name.lower()).split()) if subject_name else set()

    records = []
    for item in extracted:
        if not isinstance(item, dict):
            continue

        name = str(item.get("site_name") or "").strip()
        if not name:
            continue

        if re.search(r"\b(total|average|avg|summary)\b", name, re.I):
            continue

        if subj_tokens:
            name_tokens = set(re.sub(r"\W+", " ", name.lower()).split())
            if len(name_tokens & subj_tokens) >= max(2, len(subj_tokens) * 0.75):
                continue

        price_m   = _num(item.get("price_sgd_m"))
        if price_m is not None and price_m > 100_000:
            price_m = round(price_m / 1_000_000, 3)
        price_psf = _num(item.get("price_psf_ppr"))
        site_area = _num(item.get("site_area_sf"))
        max_gfa   = _num(item.get("max_gfa_sf"))

        if price_m is None and price_psf is None:
            continue

        if price_psf is None and price_m is not None and max_gfa:
            price_psf = round(price_m * 1_000_000 / max_gfa)

        tenure_raw = item.get("tenure")
        tenure_yrs = _parse_tenure_yrs(tenure_raw)
        addr       = str(item.get("address")      or "").strip()
        zoning     = str(item.get("land_zoning")   or "").strip()
        launch_dt  = str(item.get("launch_date")   or "").strip()
        remarks    = str(item.get("remarks")        or "").strip()
        sale_type  = str(item.get("sale_type")      or "").strip()

        parts = []
        if sale_type:
            parts.append(sale_type)
        if remarks and remarks.lower() not in sale_type.lower():
            parts.append(remarks)
        comment = "  ".join(parts)

        records.append({
            "raw_description": f"{name}\n{addr}" if addr else name,
            "property_name":   name,
            "address":         addr,
            "launch_date":     launch_dt,
            "land_zoning":     zoning,
            "tenure_raw":      str(tenure_raw or "").strip(),
            "tenure_yrs":      tenure_yrs,
            "site_area_sf":    int(site_area) if site_area else None,
            "max_gfa_sf":      int(max_gfa)   if max_gfa   else None,
            "price_sgd_m":     price_m,
            "price_psf_ppr":   price_psf,
            "comment":         comment,
            "location":        "",
            "quality":         "",
            "_source":         "image",
        })

    print(f"  → {len(records)} valid records after filtering")
    return records


# ─────────────────────────────────────────────────────────────────────────────
# OLLAMA CLASSIFICATION  (location + quality)
# ─────────────────────────────────────────────────────────────────────────────

def _classify_system(country_name: str) -> str:
    return (
        f"You are a senior {country_name} commercial real estate analyst. "
        "Classify land sale comparables by location tier and site quality. "
        'Return ONLY {"comparables": [...]} — no markdown fences.'
    )


_CLASSIFY_PROMPT = """\
SUBJECT PROPERTY:
{subject_json}

LAND SALE COMPARABLES (0-indexed):
{comps_json}

Classify EVERY comparable ({n} entries). For each assign:

  address — the geocodable address to use for map plotting. Rules:
            • If the site_name or address field contains a street address
              (has a street number or road name), return the address portion.
            • If the text is a specific named site or development with no street
              address (e.g. "Jurong Lake District Site A"), use the site name as-is.
            • If the text is a generic description with no identifiable address
              or site name, return null — it cannot be reliably plotted.

Return ONLY:
{{"comparables": [{{"index":<int>,"address":<str or null>}},...]}}\
"""


def classify_land_comps(records: list, subject_cfg: dict,
                         max_comps: int, llm_cfg: dict) -> list:
    """Use LLM (GPT or Ollama) to clean up geocodable addresses. Falls back to no-op on error.
    Location and Quality are never set here — they are always left blank and
    filled in manually by the analyst.
    """
    provider = (llm_cfg or {}).get("provider", "ollama")
    country  = subject_cfg.get("country_name", "Singapore")

    try:
        slim = [{"index":      i,
                 "site_name":  r["property_name"],
                 "address":    r["address"],
                 "zoning":     r["land_zoning"],
                 "tenure_yrs": r["tenure_yrs"],
                 "price_sgd_m":r["price_sgd_m"]}
                for i, r in enumerate(records)]
        prompt = _CLASSIFY_PROMPT.format(
            subject_json = json.dumps(subject_cfg, indent=2),
            comps_json   = json.dumps(slim, indent=2),
            n            = len(slim),
        )
        messages = [
            {"role": "system", "content": _classify_system(country)},
            {"role": "user",   "content": prompt},
        ]

        if provider == "openai":
            from tools.llm_client import openai_chat as _openai_chat
            raw = _openai_chat(llm_cfg, messages, json_mode=True)
            model_label = (llm_cfg or {}).get("openai_model", "gpt-4o-mini")
        else:
            ocfg     = llm_cfg.get("ollama", {})
            base_url = ocfg.get("base_url", "http://localhost:11434")
            model_label = ocfg.get("model", "qwen2.5:3b")
            raw = _ollama_post(base_url, model_label, messages, timeout=120)

        raw    = re.sub(r"^```[a-z]*\n?", "", raw.strip())
        raw    = re.sub(r"\n?```$",       "", raw)
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            parsed = parsed.get("comparables") or next(
                (v for v in parsed.values() if isinstance(v, list)), [])

        for cls in parsed:
            idx = cls.get("index")
            if idx is not None and 0 <= idx < len(records):
                llm_addr = cls.get("address")
                if llm_addr:
                    records[idx]["address"] = str(llm_addr).strip()
                elif llm_addr is None and "address" in cls:
                    records[idx]["address"] = ""

        provider_label = "GPT" if provider == "openai" else "Ollama"
        print(f"      {provider_label} ({model_label}) processed {len(parsed)} records.")

    except Exception as exc:
        provider_label = "GPT" if provider == "openai" else "Ollama"
        print(f"      {provider_label} classify failed ({exc.__class__.__name__}: {exc}). "
              "Addresses kept as-is from source.")

    return records


def _classify_rules(records: list, subject_cfg: dict = None):
    """No-op fallback — location/quality are never inferred from keywords."""
    for r in records:
        r.setdefault("location", "")
        r.setdefault("quality",  "")


# ─────────────────────────────────────────────────────────────────────────────
# GEOCODING
# ─────────────────────────────────────────────────────────────────────────────

def _geocode_comps(records: list, mapbox_tok: str,
                   country_code: str, country_name: str,
                   s_lon: float, s_lat: float) -> list:
    suffix = f", {country_name}" if country_name else ""
    for r in records:
        name  = str(r.get("property_name") or r.get("site_name") or "").strip()
        addr  = str(r.get("address") or "").strip()
        name = name.replace("\n", " ").replace("\r", " ").strip()
        addr = addr.replace("\n", " ").replace("\r", " ").strip()

        if not name and not addr:
            r["lon"], r["lat"], r["distance_km"] = None, None, 9999.0
            r["_geo_provider"] = "failed"
            r["_geo_note"]     = "no address or name"
            print(f"      {'(no name)':<46}  NOT PLOTTED — no address or name")
            continue

        _STREET_TYPES = {"street", "st", "road", "rd", "avenue", "ave",
                         "crescent", "drive", "dr", "lane", "ln",
                         "place", "pl", "way", "boulevard", "blvd",
                         "terrace", "court", "ct", "close", "circle",
                         "jalan", "lorong", "tanjong"}
        _addr_words = set(re.sub(r"[^\w\s]", " ", addr.lower()).split())
        _real_addr  = addr if (re.search(r"\d", addr) and _addr_words & _STREET_TYPES) else ""

        queries = _build_geocode_queries(name, _real_addr, "")
        source  = "address" if _real_addr else "name"

        try:
            lon, lat, geo_note = geocode_with_fallbacks(queries, mapbox_tok, country_code)
            r["lon"], r["lat"] = lon, lat
            r["distance_km"]   = _haversine_km(lon, lat, s_lon, s_lat)
            r["_geo_provider"] = geo_note
            r["_geo_note"]     = geo_note
            tag = " (by name)" if source == "name" else ""
            if _near_country_centroid(lon, lat, country_code):
                r["_geo_suspect"] = True
                print(f"      {name[:46]:<46}  {r['distance_km']:>5.2f} km{tag}  "
                      f"⚠ ON COUNTRY CENTROID — likely invalid, VERIFY (flagged for review)")
            else:
                print(f"      {name[:46]:<46}  {r['distance_km']:>5.2f} km{tag}")
        except Exception as exc:
            r["lon"], r["lat"], r["distance_km"] = None, None, 9999.0
            r["_geo_provider"] = "failed"
            r["_geo_note"]     = str(exc)
            print(f"      {name[:46]:<46}  FAILED — {exc}")
    return records


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def run(config_path: str = "configs/deal_config.json",
        generate_map: bool = False,
        from_records: str = None,
        refinement_file: str = None):

    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)

    subject_cfg  = cfg["subject_property"]
    mb_cfg       = cfg.get("mapbox", {})
    mapbox_tok   = mb_cfg.get("token", "")
    llm_cfg      = cfg.get("llm", {"provider": "ollama",
                                    "ollama": {"base_url": "http://localhost:11434",
                                               "model": "qwen2.5:3b"}})
    ollama_cfg   = llm_cfg.get("ollama", {})
    base_url     = ollama_cfg.get("base_url", "http://localhost:11434")
    model        = ollama_cfg.get("model",    "qwen2.5:3b")
    country_code = cfg.get("country_code",
                           subject_cfg.get("country_code", "SG"))
    country_name = subject_cfg.get("country_name", "")
    prop_name    = subject_cfg["property_name"]
    deal_name    = subject_cfg.get("deal_name", prop_name)
    # Strip characters Windows forbids in file paths (< > : " / \ | ? * and control
    # chars) — an unsanitised name causes [Errno 22] Invalid argument on Windows.
    deal_slug    = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", deal_name).strip(" .").replace(" ", "_") or "deal"
    max_comps    = cfg.get("parameters", {}).get("max_comps",  10)
    bala_yield   = cfg.get("parameters", {}).get("bala_yield", 0.06)

    # ── Paths ─────────────────────────────────────────────────────────────────
    # Supports Excel (.xlsx), PDF (.pdf), and/or Image (.png/.jpg) for one deal.
    # land_input_file       → Excel source  (config key: "land_input_file")
    # land_input_pdf_file   → PDF source    (config key: "land_input_pdf_file")
    # land_input_image_file → Image/screenshot (config key: "land_input_image_file")
    _xl_cfg  = cfg.get("land_input_file")
    input_excel_files = ([_xl_cfg] if isinstance(_xl_cfg, str) else list(_xl_cfg)) if _xl_cfg else []
    input_file        = input_excel_files[0] if input_excel_files else None
    _pdf_cfg          = cfg.get("land_input_pdf_file", [])
    input_pdf_files   = [_pdf_cfg] if isinstance(_pdf_cfg, str) else list(_pdf_cfg)
    input_pdf_file    = input_pdf_files[0] if input_pdf_files else None
    _img_cfg  = cfg.get("land_input_image_file")
    input_image_files = ([_img_cfg] if isinstance(_img_cfg, str) else list(_img_cfg)) if _img_cfg else []
    input_image_file  = input_image_files[0] if input_image_files else None

    if not from_records:
        if not input_excel_files and not input_pdf_files and not input_image_files:
            raise ValueError(
                "No input file found in config.\n"
                "Add:  \"land_input_file\": \"Input_files/your_land_comps.xlsx\"  (Excel)\n"
                "  or  \"land_input_pdf_file\": \"Input_files/your_report.pdf\"   (PDF)\n"
                "  or  \"land_input_image_file\": \"Input_files/table_screenshot.png\"  (Image)\n"
                "  or any combination for a mixed-source deal."
            )
        for _xf in input_excel_files:
            if not Path(_xf).exists():
                raise FileNotFoundError(f"Input Excel not found: {_xf}")
        for _pf in input_pdf_files:
            if not Path(_pf).exists():
                raise FileNotFoundError(f"Input PDF not found: {_pf}")
        for _imf in input_image_files:
            if not Path(_imf).exists():
                raise FileNotFoundError(f"Input image not found: {_imf}")

    output_file = cfg.get("output_file",
                          f"output/{deal_slug}/{deal_slug}.xlsx")
    out_dir   = Path(output_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    out_excel = str(out_dir / f"Land_Sale_Comps_{deal_slug}.xlsx")
    out_geo   = str(out_dir / f"Land_Sale_Comps_{deal_slug}_geo.json")
    out_map   = str(out_dir / f"Land_Sale_Comps_{deal_slug}_map.png")

    print(f"\n{'='*64}")
    print(f"  Land Sale Comps : {deal_name}")
    print(f"{'='*64}")
    for _i, _xf in enumerate(input_excel_files, 1):
        _lbl = f" {_i}" if len(input_excel_files) > 1 else ""
        print(f"  Excel{_lbl} → {_xf}")
    for _i, _pf in enumerate(input_pdf_files, 1):
        _lbl = f" {_i}" if len(input_pdf_files) > 1 else ""
        print(f"  PDF{_lbl}   → {_pf}")
    for _i, _imf in enumerate(input_image_files, 1):
        _lbl = f" {_i}" if len(input_image_files) > 1 else ""
        print(f"  Image{_lbl} → {_imf}")
    print(f"  Output → {out_excel}")

    # ── 1. Geocode subject (only if mapbox token present) ─────────────────────
    s_lon = s_lat = None
    if mapbox_tok:
        address = subject_cfg.get("address", "")
        print(f"\n[1/5] Geocoding subject property")
        try:
            s_lon, s_lat, _ = geocode_with_fallbacks(
                [f"{prop_name}, {address}", address, prop_name],
                mapbox_tok, country_code,
            )
            print(f"      {prop_name}  →  ({s_lon:.5f}, {s_lat:.5f})")
        except Exception as e:
            print(f"      Geocoding failed: {e}  (distance sorting skipped)")
            s_lon = s_lat = None
    else:
        print(f"\n[1/5] No Mapbox token — geocoding skipped (comps ranked by order)")

    # ── 2. Parse input files OR load from saved records ──────────────────────
    out_records = str(out_dir / f"Land_Sale_Comps_{deal_slug}_records.json")

    if from_records:
        print(f"\n[2/5] Loading extracted records from  {from_records}")
        try:
            with open(from_records, encoding="utf-8") as rf:
                records = json.load(rf)
            if not isinstance(records, list):
                raise ValueError("Expected a JSON array of records.")
            print(f"      → {len(records)} records loaded from JSON")
        except Exception as e:
            raise ValueError(f"Cannot load records JSON ({from_records}): {e}") from e
    else:
        print(f"\n[2/5] Parsing input files  (keyword → fuzzy → LLM column mapping)")
        records = []
        openai_key = llm_cfg.get("openai_api_key", "")

        for _xl_i, _xl_path in enumerate(input_excel_files, 1):
            _src_label = f"excel_{_xl_i}" if len(input_excel_files) > 1 else "excel"
            _tag = f" {_xl_i}" if len(input_excel_files) > 1 else ""
            print(f"  [Excel{_tag}] {Path(_xl_path).name}")
            _excel_records, _xl_pu = parse_input_excel(_xl_path, base_url, model,
                                                        subject_name=prop_name, llm_cfg=llm_cfg)
            for _r in _excel_records:
                _r["_source"] = _src_label
            records += _excel_records
            print(f"      → {len(_excel_records)} records from Excel{_tag}")

        for _pdf_i, _pdf_path in enumerate(input_pdf_files, 1):
            _src_label = f"pdf_{_pdf_i}" if len(input_pdf_files) > 1 else "pdf"
            _tag = f" {_pdf_i}" if len(input_pdf_files) > 1 else ""
            print(f"  [PDF{_tag}] {Path(_pdf_path).name}")
            try:
                _pdf_records = _parse_pdf_records(
                    _pdf_path, llm_cfg, subject_name=prop_name)
            except Exception as _e:
                print(f"      [PDF{_tag}] extraction failed: {_e}  — skipping this file")
                _pdf_records = []
            for _r in _pdf_records:
                _r["_source"] = _src_label
            records += _pdf_records
            print(f"      → {len(_pdf_records)} records from PDF{_tag}")

        for _img_i, _img_path in enumerate(input_image_files, 1):
            _src_label = f"image_{_img_i}" if len(input_image_files) > 1 else "image"
            _tag = f" {_img_i}" if len(input_image_files) > 1 else ""
            print(f"  [Image{_tag}] {Path(_img_path).name}")
            _img_records = _parse_image_records(
                _img_path, llm_cfg, openai_key=openai_key, subject_name=prop_name)
            for _r in _img_records:
                _r["_source"] = _src_label
            records += _img_records
            print(f"      → {len(_img_records)} records from Image{_tag}")

        print(f"      → {len(records)} valid records extracted (combined)")

        if records:
            with open(out_records, "w", encoding="utf-8") as rf:
                json.dump(records, rf, indent=2, default=str)
            print(f"  Records → {out_records}")

    if not records:
        print("  No qualifying records found in input files.")
        return

    # ── Auto-detect price unit from input column headers ──────────────────────
    _detected_pu = next(
        (r["_price_unit"] for r in records if r.get("_price_unit") == "B"), "M"
    )
    if _detected_pu != subject_cfg.get("price_unit", "M"):
        subject_cfg["price_unit"] = _detected_pu
        print(f"  [Price unit] Auto-detected '{_detected_pu}' from input column headers")

    # ── Apply analyst refinement instructions (only when --refinement-file given) ──
    if refinement_file:
        print(f"\n  Applying refinement instructions from {refinement_file} ...")
        instructions = Path(refinement_file).read_text(encoding="utf-8").strip()
        _geo_path = Path(out_geo)
        if _geo_path.exists():
            try:
                _geo = json.loads(_geo_path.read_text(encoding="utf-8"))
                _marker_by_prop = {
                    c["property"].lower().strip(): c["map_marker"]
                    for c in _geo.get("comps", [])
                    if c.get("property") and c.get("map_marker")
                }
                for rec in records:
                    pn = str(rec.get("property_name") or rec.get("site_name") or "").lower().strip()
                    if pn in _marker_by_prop:
                        rec["_map_marker"] = _marker_by_prop[pn]
            except Exception:
                pass
        records = _apply_refinement(records, instructions, base_url, model)
        for rec in records:
            rec.pop("_map_marker", None)
        if not records:
            print("  Refinement resulted in 0 records — nothing to output.")
            return
        with open(out_records, "w", encoding="utf-8") as _rf:
            json.dump(records, _rf, indent=2, default=str)
        print(f"  Refined records saved → {out_records}  ({len(records)} kept)")

    # ── 3. Classify ───────────────────────────────────────────────────────────
    _provider_label = "GPT" if llm_cfg.get("provider") == "openai" else "Ollama"
    print(f"\n[3/5] Classifying via {_provider_label}  (location / quality)")
    records = classify_land_comps(records, subject_cfg,
                                   max_comps=len(records), llm_cfg=llm_cfg)
    print(f"      → {len(records)} records classified")

    # ── Post-classification field validation ─────────────────────────────────
    # Land comps have no Location or Quality column in the source data.
    # These fields are always left blank — the analyst fills them in manually
    # via the editable preview table.
    for r in records:
        r["location"] = ""
        r["quality"]  = ""

    # ── 4. Geocode + sort by distance (if token available) ────────────────────
    if mapbox_tok and s_lon is not None:
        print(f"\n[4/5] Geocoding comparables")
        records = _geocode_comps(records, mapbox_tok, country_code, country_name,
                                  s_lon, s_lat)
    else:
        print(f"\n[4/5] Geocoding skipped — comps kept in input order")

    for i, r in enumerate(records):
        r["map_marker"] = str(i + 1)

    # Console summary
    currency = subject_cfg.get("currency_symbol",
                                subject_cfg.get("currency", "S$"))
    print(f"\n  {'#':<3} {'Site':<42} {'km':>5}  {'Price M':>9}  "
          f"{'psf ppr':>8}  {'Yrs':>5}")
    print("  " + "─" * 76)
    for r in records:
        km    = r.get("distance_km", 0) if r.get("distance_km") is not None else 0
        pm    = float(r.get("price_sgd_m")   or 0)
        psf   = float(r.get("price_psf_ppr") or 0)
        yrs   = r.get("tenure_yrs")
        yrs_s = "FH" if (yrs is None or yrs == 0) else str(yrs)
        name  = str(r.get("property_name") or "")[:42]
        print(f"  {r['map_marker']:<3} {name:<42} {km:>5.2f}  "
              f"{currency}{pm:>7.1f}M  {psf:>8.0f}  {yrs_s:>5}")

    # ── 5. Render ─────────────────────────────────────────────────────────────
    print(f"\n[5/5] Rendering")
    _is_global = country_name.lower() not in ("", "singapore")
    if _is_global:
        subj_row  = _global_land_tbl.subject_to_row(subject_cfg, subject_cfg)
        comp_rows = [_global_land_tbl.comp_to_row(r, subject_cfg) for r in records]
        _global_land_tbl.build_workbook(subj_row, comp_rows, subject_cfg, out_excel)
    else:
        _pu       = subject_cfg.get("price_unit", "M")
        schema    = get_land_schema(subject_cfg)
        subj_row  = subject_to_row(subject_cfg)
        comp_rows = [comp_to_row(r, _pu) for r in records]
        build_workbook(subject_cfg, subj_row, comp_rows, out_excel, schema)
    print(f"  Excel → {out_excel}")

    if s_lon is not None and s_lat is not None:
        _geo_comps = [
            {"map_marker": r["map_marker"],
             "property":   str(r.get("property_name") or ""),
             "address":    str(r.get("address") or ""),
             "lon": r.get("lon"), "lat": r.get("lat")}
            for r in records
        ]
        write_geo_sidecar(out_geo, s_lon, s_lat, _geo_comps, mb_cfg)
        print(f"  Geo   → {out_geo}")

    # ── Write lon/lat + map_marker back into _records.json ───────────────────
    try:
        _rj_path = Path(out_records)
        if _rj_path.exists():
            with open(_rj_path, encoding="utf-8") as _rf:
                _saved_records = json.load(_rf)
            _meta_by_name = {
                str(r.get("property_name") or r.get("site_name") or ""): {
                    "lon":           r.get("lon"),
                    "lat":           r.get("lat"),
                    "map_marker":    r.get("map_marker"),
                    "_geo_provider": r.get("_geo_provider"),
                    "_geo_note":     r.get("_geo_note"),
                }
                for r in records
            }
            _any_updated = False
            for sr in _saved_records:
                _name = str(sr.get("property_name") or sr.get("site_name") or "")
                _meta = _meta_by_name.get(_name)
                if _meta:
                    for _fld in ("lon", "lat", "map_marker",
                                 "_geo_provider", "_geo_note"):
                        if _meta.get(_fld) is not None and sr.get(_fld) != _meta[_fld]:
                            sr[_fld] = _meta[_fld]
                            _any_updated = True
            if _any_updated:
                with open(_rj_path, "w", encoding="utf-8") as _rf:
                    json.dump(_saved_records, _rf, indent=2, default=str)
                print(f"  Records (+ coords + markers) → {_rj_path}")
    except Exception:
        pass

    if generate_map:
        if not mapbox_tok:
            print("  Map skipped  (no Mapbox token in config)")
        elif s_lon is None:
            print("  Map skipped  (subject geocoding failed)")
        else:
            comps_geo = [
                (r["map_marker"], r["lon"], r["lat"])
                for r in records if r.get("lon") is not None
            ]
            render_map(
                subject_lonlat = (s_lon, s_lat),
                comps          = comps_geo,
                token          = mapbox_tok,
                output_path    = out_map,
                style          = mb_cfg.get("style",    "streets-v12"),
                width          = mb_cfg.get("width",    1200),
                height         = mb_cfg.get("height",   900),
                padding        = mb_cfg.get("padding",  100),
                pin_size       = mb_cfg.get("pin_size", "l"),
            )
            print(f"  Map   → {out_map}")
    else:
        print("  Map skipped  (pass --map to generate)")

    print(f"\n  Done — {len(records)} comp(s) written to {out_excel}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate land sale comps from input Excel (no web search)"
    )
    parser.add_argument("--config", default="configs/deal_config.json",
                        help="Path to deal config JSON")
    parser.add_argument("--map", action="store_true",
                        help="Generate Mapbox map PNG")
    parser.add_argument("--from-records", metavar="JSON", default=None,
                        help="Skip extraction; load records from a previously saved "
                             "_records.json file.")
    parser.add_argument("--refinement-file", metavar="TXT", default=None,
                        help="Path to a text file with analyst instructions for "
                             "refining the records list (used with --from-records).")
    args = parser.parse_args()
    run(args.config, generate_map=args.map,
        from_records=args.from_records,
        refinement_file=args.refinement_file)
