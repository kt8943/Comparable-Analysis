#!/usr/bin/env python3
"""
scan_input_rent_comps.py
========================
Reads a user-provided Excel file of rental comparables, uses Ollama to
intelligently map columns and classify records, then produces the formatted
9-column output Excel and an optional Mapbox location map.

No web search — all data comes from the input file specified by
``rent_input_file`` in the deal config.

Pipeline
--------
  1  GEOCODE    Geocode the subject property via Mapbox
  2  PARSE      Read input Excel; Ollama auto-detects column mapping
  3  CLASSIFY   Ollama assigns location, quality, lease type per comp
  4  GEOCODE    Geocode each comp; sort by distance from subject
  5  RENDER     Formatted 9-column Excel; optional Mapbox map PNG

Usage
-----
    python3 scan_input_rent_comps.py --config configs/deal_config_88_Cecil.json
    python3 scan_input_rent_comps.py --config configs/deal_config_88_Cecil.json --map

Config keys used
----------------
    rent_input_file       : path to the Excel with rental comp data (required)
    output_file           : used to derive the output folder
    parameters.max_comps  : maximum comps to include (default 10)
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

from generate_rent_comps_table import (
    compute_eff_rent,
    build_workbook,
    get_rent_schema,
)
from generate_rent_comps_map import render_map
import generate_global_rent_comps_table as _global_rent_tbl
from generate_comps_map_base import geocode_any as geocode_with_fallbacks, build_geocode_queries as _build_geocode_queries
from tools.calculations import haversine_km as _haversine_km, parse_num as _num
from tools.json_utils import fix_json as _fix_json, split_json_arrays as _split_json_arrays
from tools.llm_client import ollama_post as _ollama_post, apply_refinement as _apply_refinement
from tools.excel_reader import find_best_sheet as _find_best_sheet, find_header_row as _find_header_row, sheet_keywords as _sheet_keywords
from tools.vision_llm import call_vision_llm as _call_vision_llm
from tools.column_mapper import map_columns as _map_cols_tool
from tools.geo_utils import write_geo_sidecar


# ─────────────────────────────────────────────────────────────────────────────
# FIELD REQUIREMENTS  —  flip the switch here, no other changes needed
# ─────────────────────────────────────────────────────────────────────────────

# Set True  → NLA (sf) must be present; records without it are dropped.
# Set False → NLA is optional; area-dependent metrics will be blank when absent.
_REQUIRE_NLA = False

# "Property name only" mode (set at runtime via run(name_only=...) / --name-only).
# When True, a comp qualifies on its property name alone — the area (NLA) gate is skipped.
_NAME_ONLY = False


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL PARSING  (Ollama-driven column mapping)
# ─────────────────────────────────────────────────────────────────────────────

# Output column names drive column detection — no hard-coded input synonyms.
# Ollama is shown the OUTPUT column display names and finds matching input columns.
_OUTPUT_FIELDS = [
    # (output_col_display_name,        internal_key,    description_for_ollama)
    ("Property",                       "building_name",  "Building or property name — e.g. 'Property', 'Building', 'Building Name', 'Development', 'Asset'"),
    ("Property Address",               "address",        "Street address — e.g. 'Address', 'Property Address', 'Street'"),
    ("Location",                       "district",       "District, submarket, or planning area — e.g. 'Location', 'District', 'Planning Area', 'Submarket', 'Area'"),
    ("Quality",                        "quality",        "Building grade — e.g. 'Quality', 'Grade', 'Building Grade', 'Class'"),
    ("Leased GLA (SF)",                "nla_sf",         "Net lettable or leased area — may be labelled NLA, GFA, GLA, Leased Area, Floor Area"),
    ("Gross Face Rents (SGD psf pm)",  "asking_rent",    "Monthly gross face rent per sq ft — e.g. 'Asking Rent', 'Gross Rent', 'Face Rent', 'Headline Rent', 'Passing Rent', 'Rent'"),
    ("Effective Rents (SGD psf pm)",   "eff_rent",       "Effective net rent per sq ft per month — e.g. 'Effective Rent', 'Net Rent', 'Net Effective Rent', 'Eff. Rent'"),
    ("Date of Lease Start",            "lease_date",     "Lease commencement date — e.g. 'Lease Date', 'Date of Lease Start', 'Commencement Date', 'Start Date'"),
    ("Lease Tenure (Yrs)",             "lease_term_yrs", "Lease term in years — e.g. 'Lease Term', 'Lease Tenure', 'Term (Yrs)', 'Duration'"),
    ("Rent-Free (Mths)",               "rent_free_mths", "Rent-free period in months — e.g. 'Rent-Free', 'Rent Free', 'RF', 'Incentive'"),
    ("Tenant",                         "tenant",         "Tenant / occupier in the lease deal — e.g. 'Tenant', 'Occupier', 'Lessee', 'Occupant'"),
    ("Type of Lease Area / Comments",  "lease_type",     "Lease/space type or remarks — e.g. 'Type' (Relocation/Expansion/New), 'Lease Type', 'Space Type', 'Remarks'"),
]
_OUTPUT_COL_TO_KEY = {col: key for col, key, _ in _OUTPUT_FIELDS}

_sheet_kws  = _sheet_keywords(_OUTPUT_FIELDS)
_best_sheet = lambda wb: _find_best_sheet(wb, _sheet_kws)


def _map_columns(headers: list, sample_rows: list,
                 base_url: str, model: str, llm_cfg: dict = None) -> tuple:
    return _map_cols_tool(headers, sample_rows, _OUTPUT_FIELDS,
                          _OUTPUT_COL_TO_KEY, base_url, model, llm_cfg=llm_cfg)


def parse_input_excel(input_file: str, base_url: str, model: str,
                      subject_name: str = "", llm_cfg: dict = None) -> list:
    """
    Read any Excel file of rent comps. Uses LLM (GPT or Ollama) to map columns.
    Returns list of dicts with standardised field names ready for
    classify_rent_comps() and geocoding.
    """
    wb   = openpyxl.load_workbook(input_file, data_only=True)
    best = _best_sheet(wb)
    ws   = wb[best]
    rows = [tuple(c.value for c in row) for row in ws.iter_rows()]

    hdr_idx = _find_header_row(rows)
    headers  = [str(c) if c is not None else "" for c in rows[hdr_idx]]
    data_rows = [r for r in rows[hdr_idx + 1:] if any(c not in (None, "") for c in r)]

    print(f"  Sheet: {best!r}  |  Header row: {hdr_idx + 1}"
          f"  |  Data rows: {len(data_rows)}")
    print(f"  Headers: {[h for h in headers if h]}")

    # Tiered column mapping: exact → keyword → fuzzy → Ollama (last resort)
    print(f"  Mapping columns …")
    col_map, unit_map = _map_columns(headers, list(data_rows[:3]), base_url, model, llm_cfg=llm_cfg)

    # Extract year from sheet name first, then fall back to lease date column header.
    _date_header_year = None
    _sheet_ym = re.search(r"\b(20\d{2})\b", best)
    if _sheet_ym:
        _date_header_year = _sheet_ym.group(1)
        print(f"  [date-year] {_date_header_year!r} extracted from sheet name {best!r}")
    else:
        _date_col_idx = col_map.get("lease_date")
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

    subj_tokens = set(re.sub(r"\W+", " ", subject_name.lower()).split()) if subject_name else set()

    records = []
    for row in data_rows:
        name = str(_get(row, "building_name") or "").strip()
        if not name:
            continue

        # Skip subject property
        if subj_tokens:
            name_tokens = set(re.sub(r"\W+", " ", name.lower()).split())
            overlap = len(name_tokens & subj_tokens)
            if overlap >= max(2, len(subj_tokens) * 0.75):
                continue

        nla    = _get_num(row, "nla_sf")
        asking = _get_num(row, "asking_rent")
        eff    = _get_num(row, "eff_rent")

        # Area (leased GLA) is the key field for a rent/lease comp. The rent rate
        # is OPTIONAL — many "key lease transactions" tables list the space leased
        # but no rate (the analyst fills the rent in the preview). So keep a comp
        # on area alone; drop only when there is no area.
        if not nla and not _NAME_ONLY:
            continue

        addr     = str(_get(row, "address")      or "").strip()
        district = str(_get(row, "district")      or "").strip()
        quality  = str(_get(row, "quality")       or "").strip()
        l_date   = str(_get(row, "lease_date")    or "").strip()
        if l_date and _date_header_year and not re.search(r"\b(?:19|20)\d{2}\b", l_date):
            l_date = f"{l_date} {_date_header_year}"
        l_term   = _get_num(row, "lease_term_yrs")
        rf_mths  = _get_num(row, "rent_free_mths")
        tnt      = str(_get(row, "tenant")        or "").strip()
        l_type   = str(_get(row, "lease_type")    or "").strip()

        records.append({
            # classify_rent_comps expects raw_description
            "raw_description": f"{name}\n{addr}" if addr else name,
            # keep individual fields for geocoding + enrichment
            "property_name":   name,
            "address":         addr,
            "location":        district,
            "quality":         quality,
            "nla_sf":          int(nla) if nla else None,
            "asking_rent":     asking,
            "eff_rent":        eff,
            "lease_date":      l_date,
            "lease_term_yrs":  l_term,
            "rent_free_mths":  rf_mths,
            "tenant":          tnt,
            "lease_type":      l_type,
            "_source":         "excel",
        })

    return records


# ─────────────────────────────────────────────────────────────────────────────
# PDF INPUT PARSING
# ─────────────────────────────────────────────────────────────────────────────

_PDF_SECTION_KEYWORDS = [
    "Leasing Comparables", "Additional Leasing Comparables",
    "Rental Comparables", "Additional Rental Comparables",
    "Comparable Leases", "Tenancy Schedule", "Rental Evidence",
    "Lease Transactions", "Recent Leasing", "Recent Leases", "Leasing Activity",
]


def _parse_pdf_records(pdf_path: str, llm_cfg: dict,
                       subject_name: str = "") -> list:
    """
    Extract rent comp records from a PDF using the shared 4-stage
    pdf_extractor pipeline (pdfplumber page discovery → table detection →
    field mapping → record assembly).

    Returns the same record format as parse_input_excel() so downstream
    classify_rent_comps() and geocoding work unchanged.
    Location and Quality are read directly from source columns (district /
    quality) when present in the PDF table — never inferred.
    """
    from pdf_extractor import extract_pdf_records

    # Reject tables that share a lease page but aren't lease comps:
    #   • Market-statistics tables (inventory / vacancy / absorption by submarket)
    #   • Key SALES transactions (seller/buyer, sale price) — those are asset sales
    # Lease-transaction tables (Property/Submarket/Tenant/SF/Type) have none of
    # these markers, so they pass through.
    _NON_RENT_TABLE_MARKERS = [
        "inventory", "vacancy", "vacant", "absorption", "under cnstr",
        "seller/buyer", "seller / buyer", "price (s$", "price(s$",
    ]

    # _OUTPUT_FIELDS already includes address, district, quality, tenant, lease_type
    # dedup=False: a building can have several lease deals (different tenants /
    # floors / areas) — each is a distinct rent comp, so do NOT merge by name.
    raw_records = extract_pdf_records(
        pdf_path, _PDF_SECTION_KEYWORDS, _OUTPUT_FIELDS,
        llm_cfg, subject_name=subject_name,
        reject_table_headers=_NON_RENT_TABLE_MARKERS,
        dedup=False,
    )
    if not raw_records:
        return []

    subj_tokens = set(re.sub(r"\W+", " ", subject_name.lower()).split()) if subject_name else set()

    records = []
    for item in raw_records:
        if not isinstance(item, dict):
            continue

        name = str(item.get("building_name") or "").strip()
        if not name:
            continue

        # Skip subject (extractor already filters, but double-check)
        if subj_tokens:
            name_tokens = set(re.sub(r"\W+", " ", name.lower()).split())
            if len(name_tokens & subj_tokens) >= max(2, len(subj_tokens) * 0.75):
                continue

        nla    = _num(item.get("nla_sf"))
        asking = _num(item.get("asking_rent"))
        eff    = _num(item.get("eff_rent"))

        # Area (leased GLA) is the key field for a rent/lease comp. The rent rate
        # is OPTIONAL — many "key lease transactions" tables list the space leased
        # but no rate (the analyst fills the rent in the preview). So keep a comp
        # on area alone; drop only when there is no area.
        if not nla and not _NAME_ONLY:
            continue

        # Date sanity: non-empty date with no 4-digit year is a sentence fragment.
        l_date = str(item.get("lease_date") or "").strip()
        if l_date and not re.search(r"\b(?:19|20)\d{2}\b", l_date):
            continue

        addr     = str(item.get("address")       or "").strip()
        district = str(item.get("district")       or "").strip()
        quality  = str(item.get("quality")        or "").strip()
        l_term   = _num(item.get("lease_term_yrs"))
        rf_mths  = _num(item.get("rent_free_mths"))
        tnt      = str(item.get("tenant")         or "").strip()
        l_type   = str(item.get("lease_type")     or "").strip()

        records.append({
            "raw_description": f"{name}\n{addr}" if addr else name,
            "property_name":   name,
            "address":         addr,
            "location":        district,
            "quality":         quality,
            "nla_sf":          int(nla) if nla else None,
            "asking_rent":     asking,
            "eff_rent":        eff,
            "lease_date":      l_date,
            "lease_term_yrs":  l_term,
            "rent_free_mths":  rf_mths,
            "tenant":          tnt,
            "lease_type":      l_type,
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
You are extracting real estate rental comparable data from a table screenshot.

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
- Rent ranges like "12-14" should be returned as the string "12-14" (not split).
- SCOPE — this is a LEASING / RENT analysis. Extract ONLY lease / rental transactions
  (a tenant leasing space at a rent). Do NOT extract:
    * ASSET / INVESTMENT SALES tables (e.g. columns Buyer, Vendor, Sale Price, Cap Rate,
      $ psf capital value), and
    * GLS / land tender tables (e.g. columns Tender, Tenderer, psf ppr, Site Area).
  Also ignore market-statistics tables (island-wide rent indices, vacancy, supply). If a
  table is a sales/land/statistics table, ignore it. If the image has no leasing table,
  return an empty array [].
"""


def _parse_image_records(image_path: str, llm_cfg: dict, openai_key: str = "",
                         subject_name: str = "") -> list:
    """
    Extract rent comp records from a table screenshot using a vision LLM.
    Returns the same record format as parse_input_excel().
    """
    print(f"  Reading image: {Path(image_path).name} ...")

    field_list = "\n".join(
        f'  "{key}": {desc}'
        for _, key, desc in _OUTPUT_FIELDS
    ) + '\n  "address": Street address of the property'

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

        name = str(item.get("building_name") or "").strip()
        if not name:
            continue

        if subj_tokens:
            name_tokens = set(re.sub(r"\W+", " ", name.lower()).split())
            if len(name_tokens & subj_tokens) >= max(2, len(subj_tokens) * 0.75):
                continue

        nla    = _num(item.get("nla_sf"))
        asking = _num(item.get("asking_rent"))
        eff    = _num(item.get("eff_rent"))

        # Area (leased GLA) is the key field for a rent/lease comp. The rent rate
        # is OPTIONAL — many "key lease transactions" tables list the space leased
        # but no rate (the analyst fills the rent in the preview). So keep a comp
        # on area alone; drop only when there is no area.
        if not nla and not _NAME_ONLY:
            continue

        addr     = str(item.get("address")       or "").strip()
        district = str(item.get("district")       or "").strip()
        quality  = str(item.get("quality")        or "").strip()
        l_date   = str(item.get("lease_date")     or "").strip()
        l_term   = _num(item.get("lease_term_yrs"))
        rf_mths  = _num(item.get("rent_free_mths"))
        tnt      = str(item.get("tenant")         or "").strip()
        l_type   = str(item.get("lease_type")     or "").strip()

        records.append({
            "raw_description": f"{name}\n{addr}" if addr else name,
            "property_name":   name,
            "address":         addr,
            "location":        district,
            "quality":         quality,
            "nla_sf":          int(nla) if nla else None,
            "asking_rent":     asking,
            "eff_rent":        eff,
            "lease_date":      l_date,
            "lease_term_yrs":  l_term,
            "rent_free_mths":  rf_mths,
            "tenant":          tnt,
            "lease_type":      l_type,
            "_source":         "image",
        })

    print(f"  → {len(records)} valid records after filtering")
    return records


# ─────────────────────────────────────────────────────────────────────────────
# GEOCODING
# ─────────────────────────────────────────────────────────────────────────────

def _geocode_comps(records: list, mapbox_tok: str,
                   country_code: str, country_name: str,
                   s_lon: float, s_lat: float) -> list:
    """Geocode each comp; attach lon / lat / distance_km.

    Uses the same strategy as sales comps:
      - If the address field looks like a real street address (has a digit
        AND a street-type keyword), use it as the primary geocoding query.
      - Otherwise (district/submarket labels like "Raffles Place", "CBD"),
        fall back to the building name — which Mapbox resolves accurately
        for well-known commercial buildings.
    """
    suffix = f", {country_name}" if country_name else ""
    for r in records:
        name = str(r.get("property") or r.get("property_name") or "").strip()
        addr = str(r.get("address") or "").strip()
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
        _has_digit  = bool(re.search(r"\d", addr))
        _is_foreign = bool(country_name) and country_name.strip().lower() != "singapore"
        if _is_foreign:
            # Foreign building/hostel names (e.g. Korean) are often un-findable, so
            # geocoding by name returns a locality centroid — many comps then stack
            # on one point. A real street address (has a number) resolves precisely.
            # District-only labels like "Gangnam" have no number → fall back to name.
            _real_addr = addr if _has_digit else ""
        else:
            _real_addr = addr if (_has_digit and _addr_words & _STREET_TYPES) else ""

        queries = _build_geocode_queries(name, _real_addr, "")
        source  = "address" if _real_addr else "name"

        try:
            lon, lat, geo_note = geocode_with_fallbacks(queries, mapbox_tok, country_code)
            r["lon"], r["lat"] = lon, lat
            r["distance_km"]   = _haversine_km(lon, lat, s_lon, s_lat)
            r["_geo_provider"] = geo_note
            r["_geo_note"]     = geo_note
            tag = " (by name)" if source == "name" else ""
            print(f"      {name[:46]:<46}  {r['distance_km']:>5.2f} km{tag}")
        except Exception as exc:
            r["lon"], r["lat"], r["distance_km"] = None, None, 9999.0
            r["_geo_provider"] = "failed"
            r["_geo_note"]     = str(exc)
            print(f"      {name[:46]:<46}  FAILED — {exc}")
    return records


# ─────────────────────────────────────────────────────────────────────────────
# ROW CONVERSION
# ─────────────────────────────────────────────────────────────────────────────

def _to_comp_row(r: dict) -> dict:
    """Normalise a classified record into the Excel row dict."""
    bldg      = str(r.get("property") or r.get("property_name") or "").strip()
    addr      = str(r.get("address") or "").strip()
    prop_text = f"{bldg}\n{addr}" if addr else bldg
    _src = r.get("_source", "")
    _src_map = {"excel": "Excel", "pdf": "PDF", "image": "Image", "manual": "Manual"}
    if _src.startswith("pdf_"):
        _src_map[_src] = "PDF " + _src[4:]
    if _src.startswith("excel_"):
        _src_map[_src] = "Excel " + _src[6:]
    if _src.startswith("image_"):
        _src_map[_src] = "Image " + _src[6:]
    return {
        "source":         _src_map.get(_src, ""),
        "property":       prop_text,
        "map_marker":     str(r.get("map_marker", "")),
        "lease_date":     str(r.get("lease_date") or ""),
        "nla_sf":         r.get("nla_sf"),
        "lease_term_yrs": r.get("lease_term_yrs"),
        "asking_rent":    r.get("asking_rent"),
        "eff_rent":       r.get("eff_rent"),
        "location":       str(r.get("location") or ""),
        "quality":        str(r.get("quality") or ""),
        "tenant":         str(r.get("tenant") or ""),
        # classify_rent_comps returns asset_type; fall back to lease_type
        "lease_type":     str(r.get("asset_type") or r.get("lease_type") or ""),
    }


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def run(config_path: str = "configs/deal_config.json",
        generate_map: bool = False,
        from_records: str = None,
        refinement_file: str = None,
        name_only: bool = False):

    global _NAME_ONLY
    _NAME_ONLY = name_only

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
    country_code = cfg.get("country_code", "")
    country_name = subject_cfg.get("country_name", "")
    prop_name    = subject_cfg["property_name"]
    deal_name    = subject_cfg.get("deal_name", prop_name)
    # Strip characters Windows forbids in file paths (< > : " / \ | ? * and control
    # chars) — an unsanitised name causes [Errno 22] Invalid argument on Windows.
    deal_slug    = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", deal_name).strip(" .").replace(" ", "_") or "deal"
    max_comps    = cfg.get("parameters", {}).get("max_comps", 10)

    # ── Paths ─────────────────────────────────────────────────────────────────
    # Supports Excel (.xlsx), PDF (.pdf), and/or Image (.png/.jpg) for one deal.
    # rent_input_file       → Excel source  (config key: "rent_input_file")
    # rent_input_pdf_file   → PDF source    (config key: "rent_input_pdf_file")
    # rent_input_image_file → Image/screenshot (config key: "rent_input_image_file")
    _xl_raw  = cfg.get("rent_input_file") or cfg.get("input_file")
    input_excel_files = ([_xl_raw] if isinstance(_xl_raw, str) else list(_xl_raw)) if _xl_raw else []
    input_file        = input_excel_files[0] if input_excel_files else None
    _pdf_cfg          = cfg.get("rent_input_pdf_file", [])
    input_pdf_files   = [_pdf_cfg] if isinstance(_pdf_cfg, str) else list(_pdf_cfg)
    input_pdf_file    = input_pdf_files[0] if input_pdf_files else None
    _img_cfg  = cfg.get("rent_input_image_file")
    input_image_files = ([_img_cfg] if isinstance(_img_cfg, str) else list(_img_cfg)) if _img_cfg else []
    input_image_file  = input_image_files[0] if input_image_files else None

    if not from_records:
        if not input_excel_files and not input_pdf_files and not input_image_files:
            raise ValueError(
                "No input file found in config.\n"
                "Add:  \"rent_input_file\": \"Input_files/your_comps.xlsx\"  (Excel)\n"
                "  or  \"rent_input_pdf_file\": \"Input_files/your_report.pdf\"  (PDF)\n"
                "  or  \"rent_input_image_file\": \"Input_files/table_screenshot.png\"  (Image)\n"
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
                          f"output/{deal_slug}/Transaction_Comparables_{deal_slug}.xlsx")
    out_dir   = Path(output_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    out_excel = str(out_dir / f"Rent_Comps_{deal_slug}.xlsx")
    out_geo   = str(out_dir / f"Rent_Comps_{deal_slug}_geo.json")
    out_map   = str(out_dir / f"Rent_Comps_{deal_slug}_map.png")

    print(f"\n{'='*64}")
    print(f"  Rent Comps from Input : {deal_name}")
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
        print(f"\n[1/5] No Mapbox token — geocoding skipped (comps ranked by relevance)")

    # ── 2. Parse input files OR load from saved records ──────────────────────
    out_records = str(out_dir / f"Rent_Comps_{deal_slug}_records.json")

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
            _excel_records = parse_input_excel(_xl_path, base_url, model,
                                               subject_name=prop_name, llm_cfg=llm_cfg)
            for _r in _excel_records:
                _r["_source"] = _src_label
            records += _excel_records
            print(f"      → {len(_excel_records)} records from Excel{_tag}")

        for _pdf_i, _pdf_path in enumerate(input_pdf_files, 1):
            _src_label = f"pdf_{_pdf_i}" if len(input_pdf_files) > 1 else "pdf"
            _tag = f" {_pdf_i}" if len(input_pdf_files) > 1 else ""
            print(f"  [PDF{_tag}] {Path(_pdf_path).name}")
            _pdf_records = _parse_pdf_records(
                _pdf_path, llm_cfg, subject_name=prop_name)
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
        print("  No qualifying records (need building name + GLA; rent optional).")
        return

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
                    pn = str(rec.get("property_name") or "").lower().strip()
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

    # ── 3. Pass records through without LLM classification ───────────────────
    print(f"\n[3/5] Skipping LLM classification — using parsed values as-is")
    classified = []
    for i, r in enumerate(records):
        c = dict(r)
        c.setdefault("property",      c.get("property_name", ""))
        c.setdefault("address",       c.get("address", ""))
        c.setdefault("location",      c.get("location", ""))
        c.setdefault("quality",       c.get("quality", ""))
        c.setdefault("lease_type",    c.get("lease_type", ""))
        c["asset_type"] = ""
        c["type"]       = "Comparable"
        c.setdefault("_source", "excel")
        classified.append(c)
    print(f"      → {len(classified)} records")

    # ── 4. Geocode + sort by distance ─────────────────────────────────────────
    print(f"\n[4/5] Geocoding comparables")
    if mapbox_tok and s_lon is not None and s_lat is not None:
        classified = _geocode_comps(classified, mapbox_tok,
                                    country_code, country_name, s_lon, s_lat)
    else:
        print("  Geocoding skipped — no Mapbox token or subject coordinates unavailable.")

    for i, r in enumerate(classified):
        r["map_marker"] = str(i + 1)

    # ── Location competitiveness (SG, URA proximity vs subject) ───────────────
    # OneMap-geocoded comps only; others left blank. SG-only (URA Master Plan).
    if s_lon is not None and s_lat is not None:
        try:
            from tools.location_score import apply_location as _apply_loc
            classified = _apply_loc(classified,
                                    subject_cfg.get("property_name", ""),
                                    subject_cfg.get("address", ""),
                                    subject_cfg.get("asset_class", ""),
                                    subj_lonlat=(s_lon, s_lat))
        except Exception as _le:
            print(f"  [location] skipped: {_le}")

    classified = compute_eff_rent(classified)

    # Print summary
    print(f"\n  {'#':<3} {'Property':<44} {'km':>5}  {'Asking':>8}  {'Eff.Rent':>8}")
    print("  " + "─" * 74)
    for r in classified:
        km     = r.get("distance_km", 0)
        asking = float(r.get("asking_rent") or 0)
        eff    = float(r.get("eff_rent")    or 0)
        name   = str(r.get("property") or r.get("property_name") or "")[:44]
        print(f"  {r['map_marker']:<3} {name:<44} {km:>5.2f}  {asking:>8.2f}  {eff:>8.2f}")

    # ── 5. Render Excel + Map ─────────────────────────────────────────────────
    print(f"\n[5/5] Rendering")
    _is_global = country_name.lower() not in ("", "singapore")
    if _is_global:
        subj_row  = _global_rent_tbl.subject_to_row(subject_cfg, subject_cfg)
        comp_rows = [_global_rent_tbl.comp_to_row(r, subject_cfg) for r in classified]
        _global_rent_tbl.build_workbook(subj_row, comp_rows, subject_cfg, out_excel)
    else:
        comp_rows = [_to_comp_row(r) for r in classified]
        schema    = get_rent_schema(subject_cfg)
        build_workbook(subject_cfg, comp_rows, out_excel, schema)

    if s_lon is not None and s_lat is not None:
        _geo_comps = [
            {"map_marker": r["map_marker"],
             "property":   str(r.get("property") or r.get("property_name") or ""),
             "address":    str(r.get("address") or ""),
             "lon": r.get("lon"), "lat": r.get("lat")}
            for r in classified
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
                str(r.get("property_name") or r.get("property") or ""): {
                    "lon":           r.get("lon"),
                    "lat":           r.get("lat"),
                    "map_marker":    r.get("map_marker"),
                    "_geo_provider": r.get("_geo_provider"),
                    "_geo_note":     r.get("_geo_note"),
                    # location competitiveness label (Superior/Comparable/Inferior,
                    # or "" when not scored) — so the preview reflects it, not the
                    # raw extracted district text.
                    "location":      r.get("location"),
                }
                for r in classified
            }
            _any_updated = False
            for sr in _saved_records:
                _name = str(sr.get("property_name") or sr.get("property") or "")
                _meta = _meta_by_name.get(_name)
                if _meta:
                    for _fld in ("lon", "lat", "map_marker",
                                 "_geo_provider", "_geo_note", "location"):
                        if _meta.get(_fld) is not None and sr.get(_fld) != _meta[_fld]:
                            sr[_fld] = _meta[_fld]
                            _any_updated = True
            if _any_updated:
                with open(_rj_path, "w", encoding="utf-8") as _rf:
                    json.dump(_saved_records, _rf, indent=2, default=str)
                print(f"  Records (+ coords + markers) → {_rj_path}")
    except Exception:
        pass

    if generate_map and s_lon is not None and s_lat is not None:
        comps_geo = [
            (r["map_marker"], r["lon"], r["lat"])
            for r in classified if r.get("lon") is not None
        ]
        render_map(
            subject_lonlat=(s_lon, s_lat),
            comps=comps_geo,
            token=mapbox_tok,
            output_path=out_map,
            style=mb_cfg.get("style",     "streets-v12"),
            width=mb_cfg.get("width",     1200),
            height=mb_cfg.get("height",   900),
            padding=mb_cfg.get("padding", 100),
            pin_size=mb_cfg.get("pin_size", "l"),
        )
        print(f"  Map   → {out_map}")
    else:
        print("  Map skipped  (pass --map to generate)")

    print(f"\n  Done — {len(classified)} comp(s) written to {out_excel}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate rent comps from a provided input Excel (no web search)"
    )
    parser.add_argument("--config", default="configs/deal_config.json",
                        help="Path to deal config JSON")
    parser.add_argument("--name-only", action="store_true",
                        help="Qualify a comp on its property name alone (skip the area gate)")
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
        refinement_file=args.refinement_file,
        name_only=args.name_only)
