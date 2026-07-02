---
name: classify_rent_comps
description: Classify rental comp records with Ollama — cleans property name, resolves geocodable address, and assigns a relevance score; location/quality/lease_type are always kept from source data
type: atomic
requires:
  config_keys:
    - llm.ollama.base_url
    - llm.ollama.model
    - subject_property.asset_class
    - subject_property.country_name
  skills: []
allowed_tools:
  - tools.llm_client.ollama_post
  - tools.json_utils.fix_json
---

## When to use
Called in Stage 3 of `analyse_rent_comps` after records have been parsed from Excel, PDF, or image. Returns a limited enrichment — only property name cleaning, geocodable address resolution, and relevance scoring. Does NOT assign location/quality/lease_type from LLM world knowledge.

## Instructions
1. Build a numbered list of `raw_description` strings from the comp records
2. Call `tools.llm_client.ollama_post` with a prompt that asks only for: clean `property` name, geocodable `address`, and `relevance` score (0–10)
3. Parse the JSON array response; merge `property`, `address`, `relevance` back into original comp dicts by 1-based index
4. Sort by `relevance` descending; assign sequential `map_marker` values
5. On Ollama failure: keep raw order, use first line of `raw_description` as property name, assign blank address

## Output format
Same list as input with these fields updated per record:
| Field | Type | Description |
|---|---|---|
| `property` | `str` | Cleaned property name (no floor/unit info) |
| `address` | `str` | Geocodable street address or building name; blank if unresolvable |
| `relevance` | `int` | 0–10 relevance to subject |
| `map_marker` | `str` | Sequential integer string (`"1"`, `"2"`, …) |

## Examples
```python
# In scan_input_rent_comps.py
from generate_rent_comps_table import classify_rent_comps

classified = classify_rent_comps(records, subject_cfg,
                                  max_comps=len(records), llm_cfg=llm_cfg)
# IMPORTANT: overwrite location, quality, lease_type with original parsed values after this call
```

## Notes
- Location, quality, and lease_type are **always** restored from the original parsed source data after classification — the LLM infers these from world knowledge of building names, which may diverge from the actual source data
- `classify_rent_comps` lives in `generate_rent_comps_table.py`, not in `tools/` — it is tightly coupled to the rent table schema
- Effective rent calculation (`compute_eff_rent`) is a separate step called after classification in `scan_input_rent_comps.py`
