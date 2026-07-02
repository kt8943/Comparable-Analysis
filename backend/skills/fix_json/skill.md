---
name: fix_json
description: Repair common LLM JSON formatting errors so the string can be parsed by json.loads()
type: atomic
requires:
  config_keys: []
  skills: []
allowed_tools:
  - tools.json_utils.fix_json
  - tools.json_utils.split_json_arrays
---

## When to use
After receiving raw text from any LLM call that is expected to return JSON — the model may wrap the response in markdown fences, use Python literals, or add trailing commas. Call this before `json.loads()`.

## Instructions
1. Call `tools.json_utils.fix_json(text)` on any LLM-returned string before parsing
2. If the LLM returned multiple arrays (one per row, common from vision models), call `tools.json_utils.split_json_arrays(text)` to extract each array and process them separately
3. After fixing, call `json.loads(fixed_text)` — if it still fails, log the raw text for debugging

## Output format
| Call | Return type | Description |
|---|---|---|
| `fix_json(text)` | `str` | Cleaned JSON string, ready for `json.loads()` |
| `split_json_arrays(text)` | `list[str]` | List of raw array strings extracted from the text |

## Examples
```python
from tools.json_utils import fix_json, split_json_arrays
import json

# Single array with common LLM errors
raw = """```json
[{"name": "ABC", "value": None,}]
```"""
cleaned = fix_json(raw)
records = json.loads(cleaned)

# Multiple arrays (vision model output)
raw_multi = '[{"a":1}]\n[{"a":2}]'
arrays = split_json_arrays(raw_multi)
records = []
for arr_str in arrays:
    records.extend(json.loads(fix_json(arr_str)))
```

## Notes
- Fixes applied (in order): strip markdown fences, remove `//` comments, remove trailing commas before `}` or `]`, replace Python literals (`None→null`, `True→true`, `False→false`)
- Does **not** fix structural errors like mismatched brackets — if `json.loads()` still fails after `fix_json`, the LLM produced malformed output that requires a retry
- `split_json_arrays` respects nesting and string escapes — safe to call on any LLM text even if it contains only one array
