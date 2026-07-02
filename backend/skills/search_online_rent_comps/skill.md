---
name: search_online_rent_comps
description: Search the web for rental comparables using OpenAI web search, classify and map results
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
  - tools.json_utils.fix_json
---

## When to use
No rent comp file is available and the analyst wants to source rental data from the web. Requires an OpenAI API key.

## Instructions
1. Load deal config; extract subject property details
2. Run proximity search via OpenAI web search (submarket → island-wide)
3. Classify results; compute effective rent where rent-free data is present
4. Geocode; sort by `haversine_km`; assign map markers
5. Write `*_records.json` and `*_geo.json`
6. Build Excel workbook; optionally render map

## Output format
| File | Description |
|---|---|
| `Online_Rent_Comps_<Deal>.xlsx` | Formatted rent comp workbook |
| `Online_Rent_Comps_<Deal>_map.png` | Mapbox static map (optional) |
| `Online_Rent_Comps_<Deal>_records.json` | Raw search results |
| `Online_Rent_Comps_<Deal>_geo.json` | Geocoded results |

## Examples
```bash
python3 search_online_rent_comps.py --config configs/deal_config_88_Cecil.json
```

## Notes
- Always review web-sourced rent data — asking rents from listings may not reflect transacted rates
