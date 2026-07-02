---
name: generate_bala_table
description: Convert the Singapore Bala Table PDF into the Excel lookup file used by all comp pipelines
type: pipeline
requires:
  config_keys: []
  skills: []
allowed_tools: []    # uses pypdf + openpyxl directly; no tools/ functions needed
---

## When to use
One-time setup task. Run once when deploying the platform on a new machine, or when `Input_files/bala_table.xlsx` is missing. All comp pipelines (sales, land) will fail without it.

## Instructions
1. Read `Input_files/bala table.pdf` using `pypdf`
2. Parse the two-column table: Column A = Remaining Leasehold Years, Column B = % of Freehold Value
3. Write parsed rows to `Input_files/bala_table.xlsx` using `openpyxl`

## Output format
`Input_files/bala_table.xlsx` — two columns, rows 1–100:
- Column A: Remaining years (1–99)
- Column B: Leasehold value as % of freehold (e.g. 60.0 for 60%)

## Examples
```bash
python3 generate_bala_table.py
```

## Notes
- Source: Appendix 2 of the Singapore Land Authority (SLA) / SISV leasehold valuation guidelines
- Only needs to be run once; the output file is committed to the repo
- `tools.calculations.bala_factor` loads this file at import time
