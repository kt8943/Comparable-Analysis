---
name: setup_new_deal
description: Create a new deal configuration JSON from a deal brief file or manual inputs
type: pipeline
requires:
  config_keys: []    # creates the config; no prior config needed
  skills: []
allowed_tools:
  - tools.llm_client.ollama_post
---

## When to use
Starting a new deal analysis. Creates `configs/deal_config_<DealName>.json` which every other skill depends on.

## Instructions
1. Collect deal details (name, address, asset class, GFA, price, tenure) from user input or a deal brief file
2. If a deal brief PDF/Excel is provided: extract text/data using `pypdf` or `openpyxl`
3. Call Ollama to derive: country, currency, currency symbol, area unit, land zoning, leasehold years, keywords for market report matching
4. Populate the deal config template with all derived and user-supplied values
5. Write `configs/deal_config_<DealName>.json`

## Output format
```json
{
  "subject_property": {
    "property_name": "...",
    "address": "...",
    "country_name": "Singapore",
    "country_code": "SG",
    "asset_class": "office",
    "gfa_sf": 123456,
    "price_sgd_m": 500.0,
    "remaining_leasehold_yrs": 77,
    "lon": 103.849074,
    "lat": 1.280914
  },
  "llm": { "ollama": { "base_url": "...", "model": "..." } },
  "mapbox_token": "pk...."
}
```

## Examples
```bash
# Via frontend — fill in the "New Deal" form in the Streamlit dashboard
# Or directly:
python3 create_deal_config.py --name "88 Cecil" --address "88 Cecil Street, Singapore"
```

## Notes
- `lon` / `lat` are geocoded from the address via Mapbox when a token is available
- All other pipeline skills require this config to exist first
