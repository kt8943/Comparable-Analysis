---
name: bala_factor
description: Look up the Singapore Bala Table to return a leasehold fraction (0.0–1.0) for a given number of remaining years
type: atomic
requires:
  config_keys: []
  skills:
    - generate_bala_table    # Input_files/bala_table.xlsx must exist
allowed_tools:
  - tools.calculations.bala_factor
  - tools.calculations.bala_expr
---

## When to use
Whenever a pipeline needs to convert remaining leasehold years into a fraction of freehold value — used for Bala-adjusted cap rate computation and for writing the Excel formula equivalent.

## Instructions
1. Call `tools.calculations.bala_factor(n)` with the remaining years integer
2. For freehold or 999-year tenure pass `n ≥ 999` → returns `1.0`
3. For `n = 1…99` — exact lookup in the loaded `_BALA_TABLE` singleton
4. For `n = 100…998` — linear interpolation between 96 % at 99 yrs and 100 % at 999 yrs
5. To embed an Excel formula string call `tools.calculations.bala_expr(x_ref)` with the cell reference string (e.g. `"F3"`)

## Output format
| Call | Return type | Example |
|---|---|---|
| `bala_factor(77)` | `float` (0.0–1.0) | `0.87` |
| `bala_factor(999)` | `float` | `1.0` |
| `bala_expr("F3")` | `str` (Excel formula) | `"IF(OR(F3<=0,...),1,IF(F3<=99,VLOOKUP(...),...))"` |

## Examples
```python
from tools.calculations import bala_factor, bala_expr

factor = bala_factor(77)          # → 0.87 (approx)
formula = bala_expr("F3")         # → Excel IF/VLOOKUP string
adj_cap_rate = ftm_cap_rate * bala_factor(comp_yrs) / bala_factor(subject_yrs)
```

## Notes
- `_BALA_TABLE` is loaded once at import time from `Input_files/bala_table.xlsx`; raises `FileNotFoundError` if missing — run `generate_bala_table` first
- Source: Appendix 2 of the SLA/SISV leasehold valuation guidelines
- `bala_expr` generates a self-contained Excel formula — the workbook must contain a sheet named `'Bala Tbl'` with the table in columns A:B rows 2–100
