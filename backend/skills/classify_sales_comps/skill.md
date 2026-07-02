---
name: classify_sales_comps
description: Classify asset sales comp records with Ollama — assigns location (Core/Fringe/Suburban), quality (Grade A/B/C), asset type, and relevance score; falls back to keyword rules if Ollama is unavailable
type: atomic
requires:
  config_keys:
    - llm.ollama.base_url
    - llm.ollama.model
    - subject_property.country_name
    - subject_property.asset_class
  skills: []
allowed_tools:
  - tools.llm_client.ollama_post
  - tools.json_utils.fix_json
---

## When to use
Called in Stage 3 of `analyse_sales_comps` after records have been parsed from Excel, PDF, or image. Enriches each raw comp record with classification fields that the analyst uses to filter and weight comparables.

## Instructions
1. Build a slim version of the raw comps (keep only: `raw_description`, `sale_date`, `gfa_sf`, `price_sgd_m`, `remaining_yrs`, `land_zoning`, `sale_type`, `stake_pct`)
2. Call `tools.llm_client.ollama_post` with a classification prompt that includes the subject property JSON and the slim comps JSON
3. Parse the JSON array response; for each item extract: `location`, `quality`, `asset_type`, `relevance_score`, `include` flag
4. Merge classification fields back into the original full comp dicts
5. On any Ollama failure, fall back to `_classify_rules`: location and quality are left blank; asset_type is derived from `sale_type` + `asset_class`
6. Assign sequential `map_marker` values (1, 2, 3 …)

## Output format
Same list as input, with these fields added to each dict:
| Field | Type | Example |
|---|---|---|
| `location` | `str` | `"Core CBD"` / `"Fringe"` / `"Suburban"` |
| `quality` | `str` | `"Grade A"` / `"Grade B"` |
| `asset_type` | `str` | `"Whole Building (Office)"` |
| `relevance_score` | `int` | `1–10` |
| `include` | `bool` | `true` |
| `map_marker` | `str` | `"1"` |

## Examples
```python
# In scan_input_sales_comps.py
classified = classify_comps(raw_comps, subject_cfg, base_url, model)
```

## Notes
- Larger models produce more consistent classification (qwen2.5:14b recommended)
- The rules fallback intentionally leaves `location` and `quality` blank — these require human judgement and cannot be reliably inferred from address strings alone
- `relevance_score` is LLM-assigned on a 1–10 scale; the dashboard lets analysts filter by minimum score
- `include` flag from LLM allows the model to flag obviously off-topic records; analyst can override in the dashboard
