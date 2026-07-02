---
name: search_online_sales_comps
description: Search the web for asset sales comparables using OpenAI web search, classify and map results
type: pipeline
requires:
  config_keys:
    - output_file
    - subject_property.property_name
    - subject_property.lon
    - subject_property.lat
    - subject_property.country_name
    - subject_property.country_code
    - subject_property.asset_class
    - llm.openai_api_key
    - mapbox_token         # optional
  skills: []
allowed_tools:
  - tools.calculations.haversine_km
  - tools.calculations.bala_factor
  - tools.json_utils.fix_json
---

## When to use
No input file is available and the analyst wants to source comparables automatically from the web. Requires an OpenAI API key. Results are less reliable than analyst-curated data; use as a starting point only.

## Instructions
1. Load deal config; extract subject property details
2. Run 3-level proximity search via OpenAI web search:
   - Level 1: CBD / immediate submarket
   - Level 2: island-wide
   - Level 3: regional (SEA / Asia-Pacific)
3. Deduplicate results by property name
4. Classify each result (location, quality, asset type) via GPT
5. Geocode; sort by `haversine_km`; assign map markers
6. Write `*_records.json` and `*_geo.json`
7. Build Excel workbook; optionally render map

## Output format
| File | Description |
|---|---|
| `Online_Comparables_<Deal>.xlsx` | Formatted output workbook (same schema as analyse_sales_comps) |
| `Online_Comparables_<Deal>_map.png` | Mapbox static map (optional) |
| `Online_Comparables_<Deal>_records.json` | Raw search results |
| `Online_Comparables_<Deal>_geo.json` | Geocoded results |

## Examples
```bash
python3 search_online_sales_comps.py --config configs/deal_config_88_Cecil.json
python3 search_online_sales_comps.py --config configs/deal_config_88_Cecil.json --map
```

## Notes
- Results depend on OpenAI's web search; accuracy varies — always review before use
- No `input_file` key required in config; the pipeline sources data itself
