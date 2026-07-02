---
name: analyse_sales_comps
description: Run the end-to-end asset sales comparables pipeline from an Excel, PDF, or image input file
type: pipeline
requires:
  config_keys:
    - input_file
    - output_file
    - llm.ollama.base_url
    - llm.ollama.model
    - subject_property.lon
    - subject_property.lat
    - subject_property.country_name
    - subject_property.country_code
    - mapbox_token         # optional — needed only for map output
  skills: []
allowed_tools:
  - tools.calculations.parse_num
  - tools.calculations.parse_remaining_yrs
  - tools.calculations.parse_sale_date
  - tools.calculations.haversine_km
  - tools.calculations.bala_factor
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
The analyst has an Excel sheet, PDF table, or screenshot of asset sales transaction comparables and wants the formatted output workbook, map, and sidecar JSON files.

## Instructions
1. Load deal config from `--config` path
2. Detect input file type from `input_file` extension (`.xlsx` / `.pdf` / image)
3. Parse records:
   - Excel → `parse_input_excel` → `_map_columns_ollama` → list of raw dicts
   - PDF   → `parse_comps_from_pdf` → column mapping → list of raw dicts
   - Image → `_parse_image_records` (calls `tools.vision_llm.call_vision_llm`)
4. Classify comps via Ollama: assigns `location`, `quality`, `asset_type`, `relevance_score`; falls back to `_classify_rules` on error
5. Compute metrics: `price_psf_gfa` (price × 1M / GFA), `adj_cap_rate` (Bala-adjusted)
6. Geocode each comp via Mapbox; sort ascending by `haversine_km` distance from subject
7. Assign `map_marker` (1, 2, 3, …) in distance order; write `*_records.json` and `*_geo.json`
8. Build Excel workbook via `build_workbook` (generate_sales_comps_table)
9. If `--map` flag: render Mapbox static PNG via `render_map`

## Output format
| File | Description |
|---|---|
| `Transaction_Comparables_<Deal>.xlsx` | Formatted 15-column workbook (two tables + Bala sheet) |
| `Transaction_Comparables_<Deal>_map.png` | Mapbox static map with numbered markers (optional) |
| `Transaction_Comparables_<Deal>_records.json` | Raw parsed records before geocoding |
| `Transaction_Comparables_<Deal>_geo.json` | Geocoded lon/lat per comp + hidden flags |

## Examples
```bash
python3 scan_input_sales_comps.py --config configs/deal_config_88_Cecil.json
python3 scan_input_sales_comps.py --config configs/deal_config_88_Cecil.json --map
```

## Notes
- `_REQUIRE_GFA = True` at top of script drops records with no GFA (default: False)
- Stake % is parsed from `sale_type` (e.g. "50.1% interest") and stored in a hidden column P
- Bala adjustment requires `Input_files/bala_table.xlsx` — run `generate_bala_table` once if missing
- Image input requires a vision-capable Ollama model (e.g. minicpm-v) or an OpenAI key
