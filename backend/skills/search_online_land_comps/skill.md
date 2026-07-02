---
name: search_online_land_comps
description: Search the web for land sale comparables using OpenAI web search, classify and map results
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
No land comp file is available and the analyst wants to source GLS / en-bloc land transactions from the web. Requires an OpenAI API key.

## Instructions
1. Load deal config; extract subject property details
2. Run proximity search via OpenAI web search
3. Clean geocodable addresses via Ollama
4. Geocode; sort by `haversine_km`; assign map markers
5. Write `*_records.json` and `*_geo.json`
6. Build Excel workbook with Bala-adjusted price formulas; optionally render map

## Output format
| File | Description |
|---|---|
| `Online_Land_Comps_<Deal>.xlsx` | Formatted land comp workbook |
| `Online_Land_Comps_<Deal>_map.png` | Mapbox static map (optional) |
| `Online_Land_Comps_<Deal>_records.json` | Raw search results |
| `Online_Land_Comps_<Deal>_geo.json` | Geocoded results |

## Examples
```bash
python3 search_online_land_comps.py --config configs/deal_config_88_Cecil.json
```

## Notes
- GLS results (government land sales) tend to have more reliable pricing data than en-bloc estimates
