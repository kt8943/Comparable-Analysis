---
name: analyse_rent_comps
description: Run the end-to-end rental comparables pipeline from an Excel, PDF, or image input file
type: pipeline
requires:
  config_keys:
    - rent_input_file
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
  - tools.calculations.haversine_km
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
The analyst has an Excel sheet, PDF table, or screenshot of rental comparable data and wants the formatted output workbook, map, and sidecar JSON files.

## Instructions
1. Load deal config from `--config` path
2. Detect input file type from `rent_input_file` extension
3. Parse records:
   - Excel → `parse_input_excel` → `_map_columns_ollama` (via `tools.column_mapper`) → list of raw dicts
   - PDF   → `parse_comps_from_pdf` → column mapping → list of raw dicts
   - Image → `_parse_image_records` (calls `tools.vision_llm.call_vision_llm`)
4. Classify and compute metrics via `classify_rent_comps` + `compute_eff_rent` (generate_rent_comps_table)
5. Geocode each comp; sort by `haversine_km` distance; assign map markers
6. Write `*_records.json` and `*_geo.json`
7. Build Excel workbook via `build_workbook` (generate_rent_comps_table)
8. If `--map` flag: render Mapbox static PNG

## Output format
| File | Description |
|---|---|
| `Rent_Comps_<Deal>.xlsx` | Formatted 9-column rent comp workbook |
| `Rent_Comps_<Deal>_map.png` | Mapbox static map (optional) |
| `Rent_Comps_<Deal>_records.json` | Raw parsed records before geocoding |
| `Rent_Comps_<Deal>_geo.json` | Geocoded lon/lat per comp + hidden flags |

## Examples
```bash
python3 scan_input_rent_comps.py --config configs/deal_config_88_Cecil.json
python3 scan_input_rent_comps.py --config configs/deal_config_88_Cecil.json --map
```

## Notes
- `_REQUIRE_NLA = True` at top of script drops records with no NLA (default: False)
- Effective rent = asking rent minus rent-free amortisation; null if rent-free months absent
- Image input requires a vision-capable Ollama model or an OpenAI key
