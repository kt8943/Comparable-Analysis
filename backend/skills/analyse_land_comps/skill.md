---
name: analyse_land_comps
description: Run the end-to-end land sale comparables pipeline from an Excel, PDF, or image input file
type: pipeline
requires:
  config_keys:
    - land_input_file
    - output_file
    - llm.ollama.base_url
    - llm.ollama.model
    - subject_property.lon
    - subject_property.lat
    - subject_property.country_name
    - subject_property.country_code
    - subject_property.remaining_leasehold_yrs
    - mapbox_token         # optional — needed only for map output
  skills: []
allowed_tools:
  - tools.calculations.parse_num
  - tools.calculations.haversine_km
  - tools.calculations.bala_factor
  - tools.calculations.bala_expr
  - tools.llm_client.ollama_post
  - tools.excel_reader.find_best_sheet
  - tools.excel_reader.find_header_row
  - tools.excel_reader.sheet_keywords
  - tools.column_mapper.map_columns_ollama
  - tools.vision_llm.call_vision_llm
  - tools.json_utils.fix_json
  - tools.json_utils.split_json_arrays
---

## When to use
The analyst has an Excel sheet, PDF table, or screenshot of land sale comparables and wants the formatted output workbook with Bala-adjusted price, map, and sidecar JSON files.

## Instructions
1. Load deal config from `--config` path
2. Detect input file type from `land_input_file` extension
3. Parse records:
   - Excel → `parse_input_excel` → `_map_columns_ollama` (via `tools.column_mapper`) → list of raw dicts
   - PDF   → `parse_comps_from_pdf` → column mapping → list of raw dicts
   - Image → `_parse_image_records` (calls `tools.vision_llm.call_vision_llm`)
4. Enrich geocodable addresses via `classify_land_comps` (Ollama; falls back to as-is)
5. Geocode each comp via Mapbox using address (not site name); sort by `haversine_km` distance
6. Assign map markers; write `*_records.json` and `*_geo.json`
7. Build Excel workbook via `build_workbook` (generate_land_comps_table)
   - Column I: `Price psf ppr` = Price × 1M / Max GFA (live Excel formula)
   - Column J: `Adj. Price psf ppr` = Price × Bala(subject) / Bala(comp) (live Excel formula)
8. If `--map` flag: render Mapbox static PNG

## Output format
| File | Description |
|---|---|
| `Land_Sale_Comps_<Deal>.xlsx` | Formatted 13-column land comp workbook with live Bala formulas |
| `Land_Sale_Comps_<Deal>_map.png` | Mapbox static map (optional) |
| `Land_Sale_Comps_<Deal>_records.json` | Raw parsed records before geocoding |
| `Land_Sale_Comps_<Deal>_geo.json` | Geocoded lon/lat per comp + hidden flags |

## Examples
```bash
python3 scan_input_land_comps.py --config configs/deal_config_88_Cecil.json
python3 scan_input_land_comps.py --config configs/deal_config_88_Cecil.json --map
```

## Notes
- Land comps are NOT filtered by location/quality — analysts fill those in manually
- Geocoding uses extracted street address only; if no address, comp is plotted at origin (0,0) and should be hidden in `_geo.json`
- Bala adjustment uses `Input_files/bala_table.xlsx`; run `generate_bala_table` once if missing
