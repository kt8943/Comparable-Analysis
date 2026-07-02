---
name: classify_land_comps
description: Enrich land comp records with Ollama — resolves a geocodable address for each site; location and quality are never inferred and are left blank for the analyst
type: atomic
requires:
  config_keys:
    - llm.ollama.base_url
    - llm.ollama.model
    - subject_property.country_name
  skills: []
allowed_tools:
  - tools.llm_client.ollama_post
  - tools.json_utils.fix_json
---

## When to use
Called in Stage 3 of `analyse_land_comps` after records have been parsed from Excel, PDF, or image. The only purpose is to produce a reliable geocodable `address` for map plotting — land sale descriptions often mix site names with generic descriptions that cannot be geocoded.

## Instructions
1. Build a slim representation of each record: `index`, `site_name` (property_name), `address`, `zoning`, `tenure_yrs`, `price_sgd_m`
2. Call `tools.llm_client.ollama_post` with a prompt that asks for: `index` and `address` (geocodable street address or specific site name; `null` if no identifiable location)
3. For each response item: if `address` is non-null, replace `records[idx]["address"]` with the cleaned value; if explicitly `null`, clear the address to prevent bad geocoding
4. On Ollama failure: keep all addresses as-is from the source (no-op fallback)

## Output format
Same list as input with `address` fields updated. Location and quality fields remain blank — they must be filled in manually by the analyst.

## Examples
```python
# In scan_input_land_comps.py
records = classify_land_comps(records, subject_cfg,
                               max_comps=len(records), llm_cfg=llm_cfg)
```

## Notes
- This is deliberately minimal — unlike `classify_sales_comps` and `classify_rent_comps`, no location tier, quality grade, or relevance score is assigned by the LLM
- Land site descriptions such as "Jurong Lake District Site A" have no street address — passing them through geocoding returns wrong coordinates; setting `address = null` causes `_geocode_comps` to skip that record cleanly (distance_km = 9999)
- `classify_land_comps` lives in `scan_input_land_comps.py`, not in `tools/`
