#!/usr/bin/env python3
"""
generate_land_comps_map.py
==========================
Plots the subject property (★) and land sale comparables (1–N) on a
Mapbox static map image for use in deal presentation materials.

Mirrors generate_rent_comps_map.py — the core geocoding and rendering
logic lives in generate_sales_comps_map.py and is shared across all
three comp types (sales / rent / land).

Usage
-----
    python3 generate_land_comps_map.py --config configs/deal_config_88_Cecil.json

Output
------
    output/<DealName>/Land_Sale_Comps_<DealName>_map.png
"""

import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).parent))

import argparse
import json
from pathlib import Path

import openpyxl

# All geocoding + rendering comes from the shared base — independent of
# sales and rent map modules.
from generate_comps_map_base import (
    geocode_with_fallbacks,
    render_map,
    _parse_property_text,
)
from generate_comps_map_base import shared_mapbox_token as _shared_mapbox_token


# ─────────────────────────────────────────────────────────────────────────────
# READ COMPS FROM LAND EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def read_land_comps_from_excel(path: str) -> list:
    """
    Return [(map_marker, property_text), ...] from a Land Sale Comparables Excel.
    Looks at the active sheet for rows where Map Marker is a number (1, 2, 3 …)
    and skips the subject (★) and blank rows.
    """
    wb   = openpyxl.load_workbook(path, data_only=True)
    ws   = wb.active
    rows = [tuple(c.value for c in row) for row in ws.iter_rows()]

    # Locate header row — find "marker" and "property" columns
    hdr_row = marker_col = prop_col = None
    for i, row in enumerate(rows):
        row_strs = [str(c or "").lower() for c in row]
        if any("marker" in s for s in row_strs) and any("property" in s for s in row_strs):
            hdr_row    = i
            marker_col = next(j for j, s in enumerate(row_strs) if "marker" in s)
            prop_col   = next(j for j, s in enumerate(row_strs) if "property" in s)
            break

    if hdr_row is None:
        return []

    comps = []
    for row in rows[hdr_row + 1:]:
        if not row or all(c in (None, "") for c in row):
            continue
        marker = str(row[marker_col] if marker_col < len(row) else "").strip()
        prop   = str(row[prop_col]   if prop_col   < len(row) else "").strip()
        if marker and prop and marker not in ("", "★", "—"):
            comps.append((marker, prop))

    return comps


# ─────────────────────────────────────────────────────────────────────────────
# STANDALONE RUNNER
# ─────────────────────────────────────────────────────────────────────────────

def run(config_path: str = "configs/deal_config.json"):
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)

    subject_cfg  = cfg["subject_property"]
    map_cfg       = cfg.get("map", {})
    token        = map_cfg.get("token")
    style        = map_cfg.get("style",    "streets-v12")
    width        = map_cfg.get("width",    1200)
    height       = map_cfg.get("height",   900)
    padding      = map_cfg.get("padding",  100)
    pin_size     = map_cfg.get("pin_size", "l")
    country_code = cfg.get("country_code", "")
    bounds_raw   = map_cfg.get("geocode_bounds")
    bounds_tuple = tuple(bounds_raw) if bounds_raw else None

    deal_name    = subject_cfg.get("deal_name", subject_cfg["property_name"])
    country_name = subject_cfg.get("country_name", "")
    prop_name    = subject_cfg["property_name"]
    address      = subject_cfg.get("address", "")
    suffix       = (f", {country_name}"
                    if country_name and country_name.lower() not in address.lower()
                    else "")

    # Derive output path
    output_file = cfg.get("output_file", f"output/{deal_name.replace(' ','_')}")
    out_dir     = Path(output_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    deal_slug   = deal_name.replace(" ", "_")

    # Prefer an already-generated land comps Excel; fall back to config output_file
    excel_candidates = [
        out_dir / f"Land_Sale_Comps_{deal_slug}.xlsx",
        out_dir / f"Online_Land_Comps_{deal_slug}.xlsx",
    ]
    excel_path = next((str(p) for p in excel_candidates if p.exists()), None)
    if excel_path is None:
        print("  No land comps Excel found. Run scan_input_land_comps.py first.")
        return

    map_output = str(out_dir / f"Land_Sale_Comps_{deal_slug}_map.png")

    print(f"\n{'='*60}\n  Land Comps Map : {deal_name}\n{'='*60}")

    # Geocode subject
    print(f"\n[1/4] Geocoding subject")
    s_lon, s_lat, _ = geocode_with_fallbacks(
        [f"{prop_name}, {address}", address, prop_name],
        token, country_code, bounds=bounds_tuple,
    )
    print(f"      {prop_name} → ({s_lon}, {s_lat})")

    # Read comps from Excel
    print(f"\n[2/4] Reading comps from {excel_path}")
    comps_raw = read_land_comps_from_excel(excel_path)
    print(f"      → {len(comps_raw)} comps found")

    # Geocode each comp
    print(f"\n[3/4] Geocoding comparables")
    comps_geo = []
    for marker, prop in comps_raw:
        name, addr_line = _parse_property_text(prop)
        queries = []
        if addr_line:
            queries.append(
                f"{addr_line}{suffix}" if suffix not in addr_line else addr_line)
        queries.append(f"{name}{suffix}")
        if addr_line:
            area = addr_line.split(",")[-1].strip()
            queries.append(f"{name}, {area}{suffix}")
        try:
            lon, lat, _ = geocode_with_fallbacks(
                queries, token, country_code, bounds=bounds_tuple)
            label = f"{addr_line or name}"[:45]
            print(f"      {marker:>2}. {label:<45} ({lon}, {lat})")
            comps_geo.append((marker, lon, lat))
        except Exception as exc:
            print(f"      {marker:>2}. {name:<45} FAILED — {exc}")

    # Render
    print(f"\n[4/4] Rendering map  (pin_size={pin_size!r})")
    render_map(
        subject_lonlat = (s_lon, s_lat),
        comps          = comps_geo,
        token          = _shared_mapbox_token(),
        output_path    = map_output,
        style          = style,
        width          = width,
        height         = height,
        padding        = padding,
        pin_size       = pin_size,
    )
    print(f"  Map → {map_output}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Land Comparables map")
    parser.add_argument("--config", default="configs/deal_config.json",
                        help="Path to deal config JSON")
    args = parser.parse_args()
    run(args.config)
