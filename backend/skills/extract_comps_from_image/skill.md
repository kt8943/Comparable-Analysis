---
name: extract_comps_from_image
description: Extract comp records from a table screenshot using a vision LLM (OpenAI or Ollama), with up to 3 retries keeping the best response
type: atomic
requires:
  config_keys:
    - llm.ollama.base_url
    - llm.ollama.model
  skills: []
allowed_tools:
  - tools.vision_llm.call_vision_llm
  - tools.json_utils.fix_json
  - tools.json_utils.split_json_arrays
  - tools.calculations.parse_num
  - tools.calculations.parse_remaining_yrs
---

## When to use
When the input source for a comp pipeline is an image file (`.png`, `.jpg`, `.jpeg`). Called inside `analyse_sales_comps`, `analyse_rent_comps`, and `analyse_land_comps` when the `input_file` extension is an image format.

## Instructions
1. Build the extraction prompt: field list derived from the comp-type `_OUTPUT_FIELDS` schema
2. Call `tools.vision_llm.call_vision_llm(image_path, prompt, llm_cfg, openai_key)` up to 3 times
3. Keep the attempt that returns the most JSON objects (counted by `"property_name"` occurrences)
4. Extract the JSON array from the best response using `re.search(r"\[[\s\S]*\]", raw)`
5. Try `json.loads` → `tools.json_utils.fix_json` → `ast.literal_eval` in order until one succeeds
6. Apply `parse_num`, `parse_remaining_yrs` to convert raw cell values to typed Python values
7. Return list of raw comp dicts (same format as `parse_input_excel`)

## Output format
List of raw comp record dicts. Keys match the `internal_key` values from `_OUTPUT_FIELDS` for the relevant comp type, plus `_source: "image"`.

## Examples
```python
from tools.vision_llm import call_vision_llm
from tools.json_utils import fix_json
import json, re

raw = call_vision_llm(image_path, prompt, llm_cfg, openai_key="")
m = re.search(r"\[[\s\S]*\]", raw)
records = json.loads(fix_json(m.group(0))) if m else []
```

## Notes
- OpenAI (gpt-4o, gpt-4o-mini) gives better accuracy than local Ollama vision models for dense tables
- Ollama requires a dedicated vision model (llava, minicpm-v, llama3.2-vision) — set in the sidebar Vision model selector; the main analysis model is used as fallback only if no vision model is selected
- `num_predict: 4096` is sent for Ollama to prevent mid-array truncation on large tables
- Up to 3 retry attempts always run; whichever returns the most records is kept
- Scanned PDFs should be converted to images first, then passed through this skill
