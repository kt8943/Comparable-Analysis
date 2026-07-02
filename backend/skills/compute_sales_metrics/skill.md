---
name: compute_sales_metrics
description: Compute price PSF GFA and Bala-adjusted cap rate for each asset sales comp record
type: atomic
requires:
  config_keys:
    - subject_property.remaining_leasehold_yrs
  skills:
    - bala_factor     # bala_table.xlsx must exist
allowed_tools:
  - tools.calculations.bala_factor
  - tools.calculations.bala_expr
---

## When to use
Called in Stage 4 of `analyse_sales_comps` after classification and geocoding. Adds computed metric fields to each comp dict that will be written to the output Excel.

## Instructions
1. Read `subject_cfg["remaining_leasehold_yrs"]` as the reference tenure for Bala adjustment
2. For each comp record:
   - `price_psf_gfa = (price_sgd_m / stake_pct) × 1,000,000 / gfa_sf` (round to integer; `None` if `gfa_sf` is missing)
   - `ftm_cap_rate = npi_yield` (forward-to-market NOI cap rate, as-is from source)
   - `adj_cap_rate = ftm_cap_rate × bala_factor(comp_remaining_yrs) / bala_factor(subject_remaining_yrs)`
3. Return the updated list

## Output format
Same list as input, with these fields added to each dict:
| Field | Type | Description |
|---|---|---|
| `price_psf_gfa` | `int \| None` | Price per SF of GFA (SGD) |
| `ftm_cap_rate` | `float` | Forward-to-market NOI cap rate (as extracted) |
| `adj_cap_rate` | `float` | Bala-adjusted cap rate relative to subject tenure |

## Examples
```python
# In scan_input_sales_comps.py
from tools.calculations import bala_factor

subj_yrs = subject_cfg["remaining_leasehold_yrs"]
for c in comps:
    price_m = float(c.get("price_sgd_m") or 0)
    stake   = float(c.get("stake_pct") or 1.0)
    gfa     = c.get("gfa_sf")
    rem_yrs = int(c.get("remaining_yrs") or 0)
    ftm_cr  = float(c.get("npi_yield") or 0)

    c["price_psf_gfa"] = round((price_m / stake) * 1e6 / gfa) if gfa else None
    c["ftm_cap_rate"]  = ftm_cr
    c["adj_cap_rate"]  = ftm_cr * bala_factor(rem_yrs) / bala_factor(subj_yrs)
```

## Notes
- `stake_pct` defaults to `1.0` (whole building) when not present — partial stake purchases divide through by stake fraction before computing per-SF metrics
- `adj_cap_rate` is only meaningful when both comp and subject are leasehold; freehold comps (remaining_yrs ≥ 999) return `bala_factor = 1.0`, so adj_cap_rate equals ftm_cap_rate
- The Excel output also includes a formula version of the Bala adjustment via `bala_expr()` for the workbook's formula sheet — that is written by `generate_sales_comps_table.py`, not here
