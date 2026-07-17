# PGIM Deal Analysis Platform

**▶ Live app:** [comparable-analysis.streamlit.app](https://comparable-analysis-sqlmhiffglczqwfz6zjc6t.streamlit.app)

A Streamlit application that produces two institutional deliverables for a real estate
deal: **comparable-analysis tables** (asset sales, land sales, rents) and an
**investment rationale** memo with a per-claim source audit.

Inputs are broker PDFs, offering memoranda, Excel comp sheets, screenshots, or a live
web search. Outputs are formatted Excel workbooks, Mapbox location maps, and a single
combined Word document.

This document is the reference for reviewers. It states, for every output column,
whether the value was **mapped** from a source, **calculated** by a rule, or
**generated** by a model — and which rule or model produced it.

---

## Table of contents

1. [Quick start](#1-quick-start)
2. [Architecture](#2-architecture)
3. [Repository layout](#3-repository-layout)
4. [The pipeline, end to end](#4-the-pipeline-end-to-end)
5. [Column reference — how every cell is produced](#5-column-reference--how-every-cell-is-produced)
6. [Calculation rules](#6-calculation-rules)
7. [The Location column — how it is generated](#7-the-location-column--how-it-is-generated)
8. [Map generation](#8-map-generation)
9. [Online search rules](#9-online-search-rules)
10. [Investment rationale rules](#10-investment-rationale-rules)
11. [Source audit](#11-source-audit)
12. [Word output format](#12-word-output-format)
13. [Configuration reference](#13-configuration-reference)
14. [Deployment](#14-deployment)
15. [Known limits and review notes](#15-known-limits-and-review-notes)
16. [Technical limitations — where we need help](#16-technical-limitations--where-we-need-help)

---

## 1. Quick start

```bash
# 1. Dependencies
pip install -r requirements.txt

# 2. One-time: build the Singapore Bala Table lookup from the source PDF
python3 backend/generate_bala_table_excel.py

# 3. Local LLM (only for the Ollama path; the GPT path needs no local model)
ollama serve
ollama pull qwen2.5:3b

# 4. Launch
streamlit run frontend/app.py
```

Credentials live in `configs/shared_settings.json` (git-ignored). On Streamlit Cloud
they come from Streamlit Secrets and are written into that file at startup by
`_bootstrap_cloud_secrets()` in `frontend/app.py`.

---

## 2. Architecture

```
Streamlit UI (frontend/app.py)
    │  subprocess per task — no shared memory, no partial state
    ▼
backend/scan_input_*.py        ← file-based comps  (Excel / PDF / image)
backend/search_online_*.py     ← web-sourced comps (OpenAI web search + connectors)
backend/generate_*_table.py    ← Excel writers (shared by both paths)
backend/generate_*_map.py      ← Mapbox static maps
backend/generate_investment_rationale.py
    │
    ▼
output/<Deal>/  *.xlsx, *_map.png, Source_Audit.xlsx  →  combined .docx
```

**The design decision everything else follows from:** the model decides *classification
and prose*; deterministic Python decides *arithmetic, formatting, and precedence*. No
number that reaches a client document is computed by an LLM. Cap rates, unit prices,
and lease adjustments are Python values or live Excel formulas.

Each backend task runs as its own subprocess. The UI streams its stdout as a run log,
so any run is reproducible from the command line with the same config file.

---

## 3. Repository layout

```text
PGIM-CompAnalysis/
├── frontend/
│   └── app.py                            Streamlit dashboard (entry point, ~5.5k lines)
│
├── backend/
│   ├── new_deal.py                       New-deal wizard (LLM-assisted field derivation)
│   ├── pdf_extractor.py                  Shared PDF comp extraction (pdfplumber + GPT vision)
│   ├── comp_classifier.py                Uploaded file → sales / rent / land
│   ├── comp_acquisition_agent.py         Bounded acquire→verify→evaluate→reflect loop
│   ├── orchestrator.py                   Deterministic task ordering
│   │
│   ├── scan_input_sales_comps.py         File-based comps → records  (one per comp type)
│   ├── scan_input_rent_comps.py
│   ├── scan_input_land_comps.py
│   │
│   ├── search_online_sales_comps.py      Web-sourced comps → records (one per comp type)
│   ├── search_online_rent_comps.py
│   ├── search_online_land_comps.py
│   │
│   ├── generate_sales_comps_table.py     Records → formatted Excel (schema lives here)
│   ├── generate_rent_comps_table.py
│   ├── generate_land_comps_table.py
│   ├── generate_global_*_table.py        Non-SG variants (no Bala adjustment)
│   │
│   ├── generate_comps_map_base.py        Geocoding providers + Mapbox static render
│   ├── generate_*_comps_map.py           Per-type map wrappers
│   │
│   ├── generate_investment_rationale.py  Rationale + RAG source audit (~2.6k lines)
│   │
│   ├── tools/                            Shared library — import from here, do not fork
│   │   ├── house_rules.py                Comp-search policy (single source of truth)
│   │   ├── calculations.py               bala_factor, parse_cap_rate, haversine, dedup
│   │   ├── column_mapper.py              3-tier input-column → field mapping
│   │   ├── location_score.py             Location column for SG comps (URA proximity)
│   │   ├── ura_landuse.py                URA Master Plan land-use lookups (local)
│   │   ├── excel_reader.py               Sheet/header detection, cell parsing
│   │   ├── llm_client.py                 Ollama wrappers, agent loop
│   │   ├── vision_llm.py                 Image → comp records
│   │   ├── json_utils.py                 JSON repair
│   │   └── onemap_auth.py                OneMap token handling
│   │
│   └── sources/                          Grounded (non-web-search) connectors
│       ├── registry.py                   get_grounded(country, comp_type, cfg, params)
│       └── sg/ura_pmi.py, ura_gls.py     data.gov.sg registries
│
├── configs/
│   ├── shared_settings.json              SECRETS — git-ignored, never distribute
│   └── deal_config_<Deal>.json           Per-deal inputs
│
└── output/<Deal>/                        Generated workbooks, maps, audit, docx
```

---

## 4. The pipeline, end to end

### Stage 0 — New deal setup

`backend/new_deal.py`. The user supplies deal name, address, asset class, GFA, price,
cap rate, tenure. An LLM derives `country_name`, `currency`, `gfa_unit`,
`land_zoning`, `location`, `submarket_keywords`, and `asset_search_keyword`. Written
to `configs/deal_config_<Deal>.json`.

Deal configs **do not** carry comp-search settings — those are house rules (§9). A deal
may override any of them by adding the key to its own `online_search` / `rent_search` /
`land_search` block.

### Stage 1 — Acquisition

Two independent paths produce the same record shape.

**A. File path** (`scan_input_*.py`)

| Step | Mechanism |
|---|---|
| 1. Sheet detection | Score every sheet by how many output-field keywords appear in any header row; highest wins |
| 2. Header detection | First row with ≥ 3 text cells |
| 3. Column mapping | `tools/column_mapper.py`, three tiers (below) |
| 4. Row qualification | Must have a non-empty name **and** at least one price value |
| 5. Unit normalisation | `detect_unit_multiplier()` reads the header text (sqm→SF, S$000→S$M, psm→psf) and returns a per-field multiplier |

Column mapping is deliberately tiered so the LLM is the last resort, not the first:

1. **Exact** — normalised header equals a known synonym. Zero false positives.
2. **Embedding** — cosine similarity against a field-synonym corpus. Offline, deterministic.
3. **LLM** — GPT or Ollama, called only for headers still unresolved.

A name-match post-correction pass then overrides the LLM where a header unambiguously
matches one field's keywords (a column literally named "Address" always maps to the
address field).

**PDF inputs** route through `pdf_extractor.py`:

- **Ollama path** — keyword page discovery → `pdfplumber` table detection → rule-based
  column mapping → row assembly + dedup.
- **GPT-4o vision path** — `pymupdf` renders every page to an image; all pages go in one
  call. The model locates the table visually and returns JSON records.

The vision path exists because some PDFs render property names as floating text that
sits visually inside a cell but outside its boundary box. `pdfplumber.extract_tables()`
returns empty strings for those cells; the vision model reads what is on the page.

Records extracted from **prose** rather than a detected table grid are tagged
`_llm_parsed` and surfaced in the UI as an AI-judgment notice, because a table grid is
evidence and a sentence is an inference.

**B. Online path** (`search_online_*.py`) — see §9.

### Stage 2 — Classification

One LLM call scores every comp at once and assigns `location`, `quality`,
`asset_type`, and `relevance_score` (1–10). Sales and rent then sort by relevance;
land sorts by distance from the subject.

### Stage 3 — Geocoding

`geocode_any()` in `generate_comps_map_base.py`. Provider is chosen by
`shared_settings.geocoding_provider`, falling back to Mapbox on any failure.

Only records with a **real address** are geocoded. Property names are not used as
geocoding queries — they resolve unreliably. A comp that geocodes to the country
centroid is flagged `ON COUNTRY CENTROID` for the analyst rather than silently plotted.

### Stage 4 — Trim to `max_results`

Applied **after** classification, so the cap keeps the most relevant comps (nearest, for
land) rather than whichever query happened to run first. Map markers are renumbered
after the trim.

### Stage 5 — Excel render

`generate_*_table.py` writes the workbook: subject block, comps block, average row,
`Params` sheet, `Bala Tbl` sheet, and (online path only) a `Sources` sheet.

### Stage 6 — Verification (LLM, guard-railed)

An optional LLM pass may **blank** a field it cannot support, and records every change
in `_verify_edits` for display. It may **not** rewrite a property name, swap a number,
or invent a value. Placeholder cleanup (`"N/A"` → `None`) is filtered out of the edit
log as noise.

---

## 5. Column reference — how every cell is produced

Legend: **Mapped** = taken from the source. **Calculated** = deterministic Python or a
live Excel formula. **Generated** = produced by a model or a scoring rule.

### 5.1 Transaction Comparables — Asset Sales

Schema: `backend/generate_sales_comps_table.py` → `OUTPUT_SCHEMA`

| # | Column | Field | Origin | Rule |
|---|---|---|---|---|
| 1 | Type | `type` | Calculated | Literal `"Subject"` or `"Comparable"` |
| 2 | Source | `source` | Generated | File path: `PDF <name>` / `Excel <name>` / `Image <name>` / `Manual`. Online path: origin label (`Web search`, `URA PMI`, `Web search + URA GLS`). Cell hyperlinks to the first verification URL; all URLs on the `Sources` sheet |
| 3 | Property | `property` | Mapped | `property_name` + newline + `address`. Falls back to first line of `raw_description` |
| 4 | Map Marker | `map_marker` | Generated | 1-based index after the relevance sort; subject renders `★`. Matches the pin on the map PNG |
| 5 | Sale Date | `sale_date` | Mapped | Free text as reported (`"Q1 2024"`, `"Mar 2024"`) — not reformatted |
| 6 | Land Zoning | `land_zoning` | Mapped | Falls back to the subject's zoning when the source omits it |
| 7 | Remaining Leasehold (Y) | `remaining_yrs` | **Calculated** | `parse_remaining_yrs()`: a number passes through; `"99 years from 2004"` derives to `77`; `"Freehold"` → `999` → displays **`FH`**. `0` means *unknown*, not freehold. Unparseable → `—` |
| 8 | GFA (SF) | `gfa_sf` | Mapped | Blank when the source reports none — **never `0`** |
| 9 | Price (SGD M) | `price_sgd_m` | Mapped | A reported range (`"600-630"`) displays as the original string; the numeric midpoint is used for psf |
| 10 | Price (SGD psf GFA) | `price_psf_gfa` | **Calculated** | **Reported first.** Only if the source gives no unit price: `round(price_m / stake × 1e6 / gfa)`. Missing input → `—` |
| 11 | FTM NOI Cap Rate | `ftm_cap_rate` | Mapped | `parse_cap_rate()` normalises to a **fraction** (`4.5%` → `0.045`) because the cell format is `0.00%`. Displays `4.50%` |
| 12 | Adj. Cap Rate | `adj_cap_rate` | **Calculated** | **Reported first.** Else a live Excel formula: `=IFERROR(I×Bala(comp_yrs)/Bala(subj_yrs),"—")`. Global (non-SG) deals skip Bala and carry FTM through |
| 13 | Location | `location` | **Generated** | SG: URA proximity score → `Superior` / `Comparable` / `Inferior` (§7). Non-SG: LLM classification |
| 14 | Quality | `quality` | Generated | LLM. Office: `Grade A+/A/B`. Logistics: `Grade A (Modern)` / `Cold Storage` / … |
| 15 | Asset Type | `asset_type` | Generated | LLM, `"<Sale Structure> (<Use>)"` e.g. `Block Sale (Office)` |
| — | Stake % | `stake_pct` | Mapped | Hidden column, one past the last visible. Feeds the psf calculation |

### 5.2 Rent Comparables

Schema: `backend/generate_rent_comps_table.py` → `RENT_SCHEMA_BASE`

| # | Column | Field | Origin | Rule |
|---|---|---|---|---|
| 1–4 | Type / Source / Property / Map Marker | — | — | As §5.1 |
| 5 | Date of Lease Start | `lease_date` | Mapped | Free text as reported |
| 6 | Leased GLA (SF) | `nla_sf` | Mapped | Required — a record without it is dropped at extraction |
| 7 | Lease Tenure (Yrs) | `lease_term_yrs` | Mapped | Drives the effective-rent calculation |
| 8 | Gross Face Rents (SGD psf pm) | `asking_rent` | Mapped | Required (or `eff_rent`) |
| 9 | Effective Rents (SGD psf pm) | `eff_rent` | **Calculated** | **Reported first.** Else `compute_eff_rent()` amortises the rent-free period over the lease term |
| 10 | Location | `location` | Generated | §7 |
| 11 | Quality | `quality` | Generated | LLM |
| 12 | Tenant | `tenant` | Mapped | As reported |
| 13 | Type of Lease Area / Comments | `lease_type` | Mapped/Generated | Source text, else LLM label (`Whole Floor (Office)`) |

### 5.3 Land Sale Comparables

Schema: `backend/generate_land_comps_table.py` → `LAND_SCHEMA_BASE`

| # | Column | Field | Origin | Rule |
|---|---|---|---|---|
| 1–3 | Type / Source / Property | — | — | As §5.1 |
| 4 | Map Marker | `map_marker` | Generated | 1-based index after the **distance** sort (not relevance) — land comps rank by proximity |
| 5 | Date of Launch | `launch_date` | Mapped | Tender launch, not award |
| 6 | Land Zoning | `land_zoning` | Mapped | As reported; falls back to the subject's zoning |
| 7 | Land Tenure (Y) | `tenure_yrs` | Calculated | As §5.1 #7; `999` → `FH` |
| 8 | Site Area (SF) | `site_area_sf` | Mapped | Land area, not GFA |
| 9 | Max GFA (SF) | `max_gfa_sf` | Mapped | Permissible GFA, i.e. site area × plot ratio when reported that way |
| 10 | Price (SGD M) | `price_sgd_m` | Mapped | Tender/award price |
| 11 | Price (SGD psf ppr) | `price_psf_ppr` | **Calculated** | **Reported first.** Else `price ÷ max_gfa` (per plot ratio, not site area) |
| 12 | Adj. Price (SGD psf ppr) | `adj_price_psf` | **Calculated** | Bala-adjusted to the subject's tenure |
| 13–15 | Location / Quality / Comment | | Generated | §7 + LLM |

**Land tables are excluded from asset-sales extraction.** A GLS or land table appearing
in a broker PDF must not populate the asset-sales comp set; the extractor filters on
table semantics, not just keywords.

---

## 6. Calculation rules

### 6.1 Precedence — applies to every computed cell

```
1. Reported directly by the source   → use it as-is
2. Not reported, inputs present      → calculate
3. Inputs missing                    → "—"
```

**Never `0`.** `0` is a measurement; a blank is an absence. Writing `0` for "not
reported" corrupts the average row and silently understates a comp.

This rule caused two real defects, both fixed and both worth understanding:

- `compute_metrics()` overwrote a source-reported psf with a calculated one whenever GFA
  existed, so the reported-first rule downstream never saw it.
- `adj_cap_rate` was always recomputed from Bala, discarding a source's own adjusted cap
  rate — twice, once in Python and again in `_write_formulas`. The `_adj_reported` flag
  now short-circuits both.

### 6.2 Cap rates

Stored as **fractions** (`0.045`), never percentages, because Excel cells use the
`0.00%` format. `parse_cap_rate()` normalises: a value `≥ 1` is divided by 100.

`parse_num()` alone strips `%` **without** rescaling — `"4.5%"` would become `4.5` and
display as `450.00%`. Always use `parse_cap_rate()` for a rate.

### 6.3 Tenure and the freehold convention

| Value | Meaning | Displays |
|---|---|---|
| `999` | Freehold | `FH` |
| `1`–`998` | Remaining years | `77 yrs` |
| `0` | **Unknown / not reported** | `—` |
| `None` | Not reported | `—` |

`0` is *not* freehold. An online-search prompt once stated `0 = freehold`, contradicting
the code; that would have converted every unknown tenure into a freehold comp.
Corrected.

### 6.4 Bala Table (Singapore leasehold adjustment)

`bala_factor(n)` in `tools/calculations.py`, from the SLA/SISV table:

| n | Factor |
|---|---|
| `n ≤ 0` or `n ≥ 999` | `1.0` (freehold) |
| `1 ≤ n ≤ 99` | SLA/SISV lookup table |
| `100 ≤ n ≤ 998` | Linear interpolation, 96% @ 99 yrs → 100% @ 999 yrs |

In Excel it is a live `VLOOKUP` against the `Bala Tbl` sheet (`bala_expr()`), so a
reviewer can change a tenure and see the adjusted cap rate move.

Singapore only. Global deals carry the FTM cap rate through unadjusted.

### 6.5 Preview and export number formatting

One shared formatter, `_fmt_grid_val(cell, header)`:

- Map Marker → plain integer (`1`, not `1.0`); subject `★` passes through
- Percentage cells (detected from the cell's own `number_format`) → `xx.00%`
- Every other number → `xxx,xxx.0` (thousands separator, exactly one decimal)
- Text → unchanged

There are **two** grid readers — `_read_excel_preview` (detail page) and
`_read_pgim_grid` (Overview page + Word export). Any new column, notice, or format must
be wired into **both**, or it will appear in one place and not the other.

---

## 7. The Location column — how it is generated

`backend/tools/location_score.py`. For Singapore comps this column is a **computed
score**, not an LLM opinion.

**Coordinates.** Geocoded via **OneMap** specifically — separate from the map pin, which
uses the sidebar provider. OneMap gives the best SG precision and keeps scoring
consistent regardless of which provider draws the map. A comp OneMap cannot resolve gets
a blank Location.

**Data.** URA Master Plan land-use polygons, held locally. No network call and no token
at scoring time. The full 181 MB GeoJSON is git-ignored; only the 1.9 MB
`_landuse_buckets.pkl` cache is deployed, so cloud still scores SG locations.

**Factors** depend on the asset class (`_sector_key()` maps `asset_class` → sector):

| Sector | Factor 1 | Factor 2 |
|---|---|---|
| office | Distance to nearest CBD node (lower better) | Commercial land-use coverage within 1 km |
| industrial / data_centre | Business land-use coverage within 1 km | Distance to nearest port/airport |
| retail | Commercial coverage within 1 km | Tier-weighted retail-centre attractiveness |
| hospitality | Tourist-attraction count within 1 km | Commercial coverage within 1 km |
| mixed | Distance to nearest CBD node | Residential + commercial coverage |

"Density" factors are **area-coverage fractions** — the share of the 1 km circle covered
by that land use — not raw parcel counts, so large estates outweigh small lots.

**Scoring.** Each factor is scored comp-vs-subject into `[-1, 1]` (subject = 0):

- Higher-is-better factors: smoothed relative difference, `(c − s) / (c + s + k)`
- Distance factors: `(s − c) / 5 km`, clamped. The fixed 5 km reference stops a subject
  sitting on a landmark (distance ≈ 0) from forcing every comp to −1.

The factor scores are averaged, then labelled:

| Score | Label |
|---|---|
| `> 0.3` | Superior |
| `−0.3 … 0.3` | Comparable |
| `< −0.3` | Inferior |

**Non-SG comps** get their Location from the classification LLM instead, because the URA
data is Singapore-only.

---

## 8. Map generation

`backend/generate_comps_map_base.py` → `render_map()`, wrapped per comp type.

**Geocoding and rendering are separate concerns with separate providers.** Google
resolves the coordinates; Mapbox draws the PNG. They share no credential and neither
falls back to the other.

### 8.1 Geocoding providers — Google

`geocode_any()` selects on `shared_settings.geocoding_provider`, **falling back to
Google** if the chosen provider fails or is misconfigured:

| Provider | Use | Requires |
|---|---|---|
| `google` | **Default**, global, best rooftop accuracy | `google_maps_key` |
| `onemap` | Singapore, free, best SG accuracy | none |
| `kakao` | Korea | `kakao_api_key` |

Deal configs carry **no** geocoding token; the key lives in Shared Settings only.

`geocode_with_fallbacks(queries, …)` tries each query in order and returns the first
hit, so callers pass a descending-specificity list:
`["<name>, <address>", "<address>", "<name>"]`.

`country_code` is applied as a component filter and must be set explicitly in the
config. There is deliberately **no address-sniffing heuristic** — a wrong country guess
silently geocodes a comp onto the wrong continent.

### 8.2 Rendering — Mapbox

The map is a **Mapbox Static Images API** PNG — not an interactive widget — so it drops
straight into Word and Excel. The token comes from `shared_settings.mapbox_token` /
the `MAPBOX_TOKEN` secret, never from a deal config.

**Why Mapbox and not Google for the image.** Google Static Maps was evaluated and
rejected on three measured grounds:

| | Mapbox | Google Static Maps |
|---|---|---|
| Max image | 1200×900 @2x = **2400×1800 px** | `size` capped at 640 per axis → **1280×960 px** |
| Zoom | fractional, exact fit to the comp extent | **integer only** |
| Marker label | full text — `10`…`15`, `★` | **one character** (A–Z, 0–9) |

Two of those fail *silently*, which is the real argument: Google returns a **square**
image for a 1200×900 request (clamping each axis independently, which changes the
aspect ratio and misplaces every pin), and a fractional zoom returns a **whole-world
map at zoom 0**. Neither raises an error. A Google port is viable — draw the pins
locally with Pillow, clamp the size keeping aspect, floor the zoom — but it costs ~65%
of the image resolution in a client-facing Word document for no benefit.

- Subject pin: red, labelled `★`, plotted first
- Comp pins: navy, labelled with the Map Marker index — the same number as the table's
  Map Marker column, which is what ties the two together
- `pin_size: "l"` uses Mapbox's built-in `pin-l` (no Pillow needed); `"xl"` / `"xxl"`
  draw oversized pins locally with Pillow
- Bounds auto-fit to all plotted points, with `padding` px of margin
- When the subject is hidden (`plot_subject=False`), comps render red — there is no
  subject to contrast against
- A comp may carry a per-pin colour override via the geo sidecar's `color` field

Comps that failed to geocode are simply absent from the PNG. The number of plotted pins
can therefore be lower than the row count in the table; the run log states which comps
were dropped.

---

## 9. Online search rules

Policy lives in **`backend/tools/house_rules.py`** — one file, applied to every deal,
existing and new, local and cloud. Deal configs do not carry these settings.

```
Precedence:  HOUSE_RULES  →  BY_ASSET_CLASS  →  the deal's own search block
```

A deal that genuinely needs different numbers sets the key in its own `online_search` /
`rent_search` / `land_search` block, and that wins. Config always beats code: nothing in
the module overrides a value a deal explicitly states.

### 9.1 The location ladder

| Tier | Radius | Escalates when |
|---|---|---|
| 1 · Proximity | `proximity_km` = **3 km** | — |
| 2 · City | `city_km` = **25 km** | tier 1 returns < `min_results` (3) |
| 3 · Country | **no distance cap** | tier 2 returns < `min_results` |

Logistics and industrial use **5 km / 50 km** (`BY_ASSET_CLASS`) — they trade over a
wider catchment. That is a property of the asset class, not of one deal.

**Tier 2 is a radius, not a municipal boundary.** The geocoder returns lon/lat only,
with no locality field to test containment against, so a true boundary test is not
available without new geo data. 25 km covers Singapore and most metros; widen `city_km`
for a larger one.

**Tier 3 has no radius.** Country containment comes from the country-scoped geocode and
country-scoped queries, not from a distance test.

### 9.2 Recency — independent of the ladder

| Comp type | `recency_months` |
|---|---|
| Sales | 60 (5 years) |
| Land | 60 (5 years) |
| Rent | **36 (3 years)** — rental evidence dates faster than capital evidence |

Applied identically to web search and grounded connectors — one cap per run.
Unparseable dates are **kept**, not dropped; every drop is logged.

Widening the search *area* never widens the *date window*.

### 9.3 `years_back` vs `recency_months` — different things

They act at opposite ends of the pipeline and nothing links them:

- **`years_back` shapes the query.** It builds the query string — `_year_window(2)` →
  `(2026 OR 2025 OR 2024)`. It is what the search is **asked for**.
- **`recency_months` filters the results.** Anything older is dropped after extraction.
  It is what is **kept**.

Setting `years_back_max` past `recency_months / 12` therefore buys rows that are then
discarded. `warn_window_vs_recency()` reports that conflict in the run log and
deliberately does **not** silently change either number.

### 9.4 Cost budget

`max_queries` = **5** per category. One query = **1 web search + 1 extract call**.

| | Calls |
|---|---|
| 5 web searches (`gpt-4o-mini-search-preview`) | 5 |
| 5 extractions (`gpt-4o-mini`) | 5 |
| 1 classification (`gpt-4o-mini`, all comps at once) | 1 |
| **Total per category, worst case** | **11** |

Search and extract are different jobs done by different models: the search model browses
the live web and returns prose plus `url_citation` annotations (the Source URLs); the
extract model never touches the web and only turns that prose into JSON.

A healthy deal costs far less — the ladder stops as soon as a tier returns
`min_results`. Results are cached by config hash; a re-run costs **zero** unless
`--refresh` is passed. Note that on a thin deal the budget, not the ladder, is the
binding constraint: 5 queries may be spent before tier 3 is reached.

### 9.5 Result limits

`max_results` = **15** per category, applied **after** classification so the cap keeps
the most relevant comps (nearest, for land).

### 9.6 Cross-source dedup

Two mechanisms:

1. **Name key** — normalised name (24 chars) + price.
2. **Same-building test** — `find_same_building()` in `tools/calculations.py`. Merges
   only when **both** hold: within **75 m** *and* within **5%** on price/rent. Catches
   one deal reported by two sources (e.g. in two languages, or under a translated name)
   without fusing two different buildings. URLs from every matching source accumulate on
   the surviving record.

75 m is deliberately tight. Two sources quoting the same building normally quote the
same canonical street address and geocode to nearly the same point, so the tolerance
only has to absorb provider jitter. The failure modes are not symmetric: a missed
duplicate shows up as two similar rows an analyst can see and merge, whereas a false
merge silently deletes evidence.

> **Historical note for reviewers.** The previous key was
> `(round(lon,2), round(lat,2), round(price / max(price*0.05, f)))`. The price term is
> algebraically `round(1/0.05) = 20` for any price above the floor — a **constant**. The
> key therefore degenerated into a bare 2-decimal coordinate cell of **~1.1 km**, which
> in a CBD is dozens of distinct towers, all merged into whichever was found first. This
> capped comp counts well below `max_results`.

### 9.7 Grounded connectors

Beyond web search, `sources/registry.py` supplies keyless registries — SG URA PMI and
URA GLS via data.gov.sg, plus broker reports. Enable per deal with
`online_search.sources: ["web_search", "ura_pmi"]`. They flow through the same
dedup → geocode → recency pipeline, capped at the city tier.

---

## 10. Investment rationale rules

`backend/generate_investment_rationale.py`. Two LLM calls: prose, then audit.

### 10.1 Structure — exactly four sections

| # | Theme | Content |
|---|---|---|
| 1 | Market fundamentals | Supply/demand balance, vacancy, absorption, completions; rental and capital-value momentum |
| 2 | Location / market preference | Why occupiers and investors prefer this submarket; demand drivers; connectivity |
| 3 | Asset & deal particulars | Asset quality, asset-class-specific angle, pricing vs comparables, risks and mitigants |
| 4 | Capital markets | Transaction volumes, investor appetite, yield trend, capital-value outlook |

Always four — never three, never five. A distinct angle the research supports (a
quantified ESG premium, a named supply moratorium) is folded into whichever section best
evidences it rather than appended as a fifth.

Section 3 is **asset-class aware**: `_ASSET_CLASS_SECTION3_ANGLES` supplies a different
lead angle for office / industrial / data centre / retail / hospitality / mixed.

Titles are model-written, 6–9 words, derived from the data actually found. Each section
is 2 paragraphs (3 only if the data supports a third distinct point), 80–130 words each.

### 10.2 Integrity rules

- **Evidence discipline** — every conclusion anchors to a specific figure or named fact
  from the research. Plain assertions are not permitted. No paragraph may be all numbers
  with no reasoning, or all reasoning with no numbers.
- **No data invention at the writing stage** — the model first lists its data points
  (STEP 1), then may use only those. A figure not listed cannot appear in the prose.
- **Repetition cap** — no statistic appears more than 3 times across all sections.
- **English only** — non-Latin source text is translated before use.
- **Source anonymisation** — reports are labelled "Research Report 1 / 2 / …" so no PDF
  filename can be echoed into body text.
- **No internal labels** — the four theme names are scaffolding and must never appear in
  the output. Enforced twice: the prompt says so, and `_strip_theme_labels()` removes
  them deterministically, because a prompt is not a guarantee with an LLM. The strip only
  fires at the start of a paragraph, so prose that legitimately says "capital markets
  remain liquid" is untouched.

### 10.3 Location context

One `gpt-4o-mini-search-preview` call per run asks what published sources say about the
subject's connectivity and precinct. It is **qualitative by construction**: the prompt
forbids stating a distance or walking time unless a source explicitly gives that figure.
"Directly connected to Raffles Place MRT, in the prime CBD" is allowed; "0.4 km from the
station" is not, unless cited.

If nothing is found, the block is omitted and section 2 falls back to demand drivers
rather than asserting anything unsourced. Claims that match this block are cited to
their source URL in the audit with citation type `Web Search`.

### 10.4 Extraction and caching

`pypdf` reads each page with a `[PAGE N]` marker. Text is truncated at 14,000 chars
keeping the **first 75% + last 25%** — executive summary and conclusions, dropping the
middle. Results cache on `filename + size + mtime`, so unchanged reports re-run
instantly.

---

## 11. Source audit

Every specific claim in the finished prose is extracted and matched to a source. Written
to `Source_Audit.xlsx` (12 columns), rows needing manual verification highlighted red.

**RAG path** (OpenAI key + PDF paths available):

1. Build a page-chunk embedding index (`text-embedding-3-small`).
2. Embed each claim, retrieve the best-matching page.
3. **Gate on the claim's single most distinctive number** — not "any number it mentions"
   — with scale-bridged fuzzy matching, so "S$1.38 billion" in the memo matches
   "1,377.8" reported in millions on the source page.
4. Reduce the page to a short sentence-level quote plus a separate full-passage
   `Context` column — never the whole page.
5. A row with no real number or keyword signal is **force-flagged low-confidence** rather
   than accepted on page-level similarity alone.

**LLM-fallback path** (no key or no PDF): the audit model proposes `source_file`,
`page_ref`, `supporting_text`, cross-checked against the cached Stage-1 *summary* via a
4-word sliding-window match.

**Citation types:** `Report` (PDF-matched), `Web Search` (location context — cross-check
reads `🌐 Web source — open the URL and verify`), `Deal Config` (from the subject's own
inputs), `Comparable Evidence` (from the comp tables).

The sheet's banner states which path produced it, because the cross-check column means
different things in each.

---

## 12. Word output format

`_build_combined_docx()` in `frontend/app.py` produces one document per deal.

### 12.1 Page setup

| Property | Value |
|---|---|
| Orientation | **Landscape** |
| Page size | US Letter, 27.94 cm × 21.59 cm (width/height swapped manually — python-docx does not swap them for you) |
| Margins | 0.5" all sides |
| Body font | **Arial Narrow 10 pt**, forced across the whole document |
| Section headings | Arial 11 pt, navy |

### 12.2 Document order

1. Deal name (Heading 0) + address
2. For each comp type present — **Rent → Sales → Land** (`_COMP_TYPES` order):
   - Section heading
   - PGIM-standard comp table
   - Location map PNG, scaled to fit the usable page box
3. Investment rationale prose

Comp types with no generated workbook are skipped silently; the document is built from
whatever exists.

### 12.3 The PGIM comp table

Built natively as a Word table (not an image), so reviewers can edit it:

```
┌──────────────────────────────────────────────┐
│ Subject Sales                    ← navy banner, merged, white bold, LEFT
├──────────────────────────────────────────────┤
│ Type │ Source │ Property │ …     ← column names, bold, centered
│ Subject │ … │ Frasers Tower │ …  ← subject row
│                                   ← blank separator row
│ Comparable Asset Sales           ← navy banner, merged
│ Comparable │ PDF … │ …           ← comp rows
│ Average │ … │ S$3,050.0 │ 3.60%  ← grey #D6DCE4, bold
└──────────────────────────────────────────────┘
```

- **Horizontal rules only** — no vertical lines (`_table_horizontal_borders`)
- Every cell centered **except** the navy banners, which stay left-aligned
- Average row shaded `#D6DCE4`, bold; averages computed only over columns whose header
  matches that type's keywords (`psf`, `cap rate` / `tenure`)
- Cell values come from `_read_pgim_grid` → `_fmt_grid_val`, so Word matches the
  on-screen preview exactly (§6.5)

---

## 13. Configuration reference

### 13.1 `configs/shared_settings.json` — **secrets, never distribute**

Git-ignored. Contains `mapbox_token`, `google_maps_key`, `openai_api_key`,
`kakao_api_key`, `onemap_email`, `onemap_password`, `ura_access_key`, and
`geocoding_provider`.

On Streamlit Cloud these come from Streamlit Secrets; `_bootstrap_cloud_secrets()` merges
them into this file at startup.

### 13.2 `configs/deal_config_<Deal>.json`

| Block | Purpose |
|---|---|
| `subject_property` | Name, address, asset class, GFA, price, cap rate, tenure, `country_name`, `currency`, `location`, `submarket_keywords`, `asset_search_keyword` |
| `country_code` | Explicit ISO code for geocoding. No heuristic fallback |
| `parameters` | `bala_yield` (default 0.06), `max_comps` |
| `openai` | `search_model`, `extract_model` |
| `mapbox` | `style`, `width`, `height`, `padding`, `pin_size` |
| `online_search` / `rent_search` / `land_search` | **Normally empty.** Only for per-deal overrides of a house rule, or `sources: [...]` to enable grounded connectors |
| `output_file` | Drives the output directory |

### 13.3 `backend/tools/house_rules.py` — comp-search policy

`HOUSE_RULES` (radii, `min_results`, `max_results`, `max_queries`, `years_back*`,
`max_level`), `RECENCY_MONTHS` per comp type, `BY_ASSET_CLASS` radius overrides. Change a
number here and every deal picks it up on the next run.

---

## 14. Deployment

Streamlit Cloud, from GitHub `kt8943/Comparable-Analysis`. Push to `main`; Cloud
auto-redeploys.

**The local working copy and the cloud repo share no git history.** Deploy by copying
`backend/` and `frontend/app.py` into a fresh clone of the cloud repo and pushing — never
by repointing `origin` and pushing the local branch.

**Never copy `configs/` to the deploy repo.** It holds live credentials. The cloud
`.gitignore` blocks the whole directory; leave it that way.

Cloud cannot reach: local network drives, the 181 MB `MasterPlan2025.geojson` (only the
1.9 MB `.pkl` cache ships), or any on-prem file. Build cloud-compatible features only.

---

## 15. Known limits and review notes

Ranked by what a reviewer should look at first.

1. **Credentials are embedded in `configs/`.** Every `deal_config_*.json` carries a live
   Mapbox token, and `deal_config_*.json` is **not** git-ignored (only
   `shared_settings.json` and `tmp*.json` are). The token is already in the local git
   history. **Before sharing this folder in any form, strip the tokens or share via the
   cloud repo, which contains no `configs/` at all.**
2. **`configs/` accumulates temp files.** ~800 `tmp*.json` from a failed cleanup path,
   each carrying the Mapbox token. Safe to delete; the generating path should clean up
   after itself.
3. **The city tier is a radius, not a boundary** (§9.1). Documented, not fixed — fixing it
   needs a locality field the geocoder does not return.
4. **The query budget can bind before the ladder finishes** (§9.4). On a thin deal, 5
   queries may be spent before tier 3 runs, so a short comp set may reflect the budget
   rather than the market.
5. **LLM classification is nondeterministic.** Two identical runs have returned different
   property names from the same PDF. This is why extraction is table-first and
   prose-derived records are flagged.
6. **Rationale quality depends on the market reports supplied.** With no reports the memo
   has little to anchor to, and the integrity rules will suppress rather than invent.
7. **Bala adjustment is Singapore-only.** Global deals carry the FTM cap rate through
   unadjusted; confirm this is intended for non-SG reviews.
8. **Location scoring is Singapore-only** and depends on the URA cache being present.

### Principles a reviewer should hold the code to

- Trust **tables**, not prose. A grid is evidence; a sentence is an inference.
- Prefer **omission over fabrication**. Blank beats a plausible guess.
- **No unit conversion** and no dropping of qualifiers on extraction.
- Every computed cell: **reported → calculated → `—`**, never `0`.
- The model classifies and writes; **Python computes**.

---

## 16. Technical limitations — where we need help

The items below are the known ceilings of the current design. They are not bugs to be
patched; each needs a decision or capability the project does not have today. Input
from reviewers is specifically wanted on all four.

### 16.1 Extraction accuracy — can a model be trained for this?

**The limitation.** Comps arrive in every shape: broker PDFs with real table grids,
PDFs whose "tables" are floating text, Excel sheets with bespoke headers, and
screenshots. Today this is handled with a tiered mapper (exact synonym → embedding →
LLM) plus a GPT-4o vision path, all against a **general-purpose** model that has never
seen our schema. It works, but accuracy is bounded by prompt engineering, and the LLM
tier is **nondeterministic** — two identical runs on the same PDF have returned
different property names.

**Where we need help.** Could a model be fine-tuned on our own labelled comp
tables — header → field mappings, and page-image → record pairs — so that reading and
mapping from an unfamiliar source becomes reliable rather than best-effort? Open
questions for a reviewer:

- Is fine-tuning (or a smaller trained extraction model) justified versus continuing to
  improve prompts and the deterministic tiers?
- Do we have, or can we build, enough labelled examples? Every past deal's input file
  and its approved output table is a training pair we already own.
- How would we evaluate it? An eval harness exists (`eval/run_extract.py` with gold
  files) but covers only a few broker formats.
- What is the accuracy target that would justify the cost — and who signs off that a
  trained model is safe for IC-facing numbers?

### 16.2 Deployment and shared memory across users

**The limitation.** The app is single-user by construction. Streamlit state is
per-session, and everything durable is a **file on the server's disk** — deal configs,
caches, outputs. Two analysts using the cloud app do not see each other's work; there
is no shared store, no accounts, no concurrency control. Two people working the same
deal will silently overwrite each other's outputs, because the only "database" is the
`output/<Deal>/` folder.

**Where we need help.** Can this be deployed as a proper web app with a shared,
persistent store — so a team works one deal together and the analysis, edits, and
sign-off are visible to everyone? Open questions:

- What backs the shared state — a real database, or object storage plus a metadata
  layer? What happens to the per-deal folder convention?
- Authentication and entitlement: who may see which deal? This becomes essential the
  moment the data is real rather than demo.
- Locking or merge semantics for two analysts editing the same comp table.
- An audit trail of who changed which cell, which the IC process will want anyway.

### 16.3 Narrative generation

**The limitation.** The investment rationale is a single prose call with heavy
guard-rails (evidence discipline, no invention at the writing stage, repetition caps,
banned phrases). Quality is capped by what the market reports contain and by the fact
that the model writes each section independently, in one pass, with no revision loop.
It cannot weigh two conflicting sources, and it cannot reason about the deal the way an
analyst does — it can only assemble what the research states.

**Where we need help.** Can we get materially better narrative out of an LLM here?
Candidate directions a reviewer might have views on:

- A draft → critique → revise loop rather than one pass, with the critic checking the
  argument, not just the format.
- Giving the writer the comp tables as structured evidence rather than a text summary,
  so it can reason over the numbers instead of restating them.
- Letting a reviewer's edits feed back as few-shot examples of house style.
- A stronger model for the prose call specifically — the extraction path can stay cheap.
- Where the honest limit is: some judgement should stay with the analyst, and the memo
  should say so rather than manufacture confidence.

### 16.4 Personal LLM accounts — internal data cannot be uploaded

**The limitation, and the most restrictive one.** The app currently runs on a
**personal OpenAI account**. Under that arrangement internal or confidential deal
material **must not** be uploaded, so every cloud run is limited to demo or public
data. This is why the pipeline is built around public sources — web search, URA
registries, broker research — and why the on-prem paths (network-drive inboxes,
confidential IMs) exist separately.

The consequence is that the tool cannot currently be pointed at the material it would
be most useful on: real offering memoranda and internal underwriting.

**Where we need help.** Moving to an approved enterprise arrangement is a
prerequisite for real use, not an optimisation. Open questions:

- Which route does the firm sanction — an enterprise OpenAI/Azure OpenAI tenancy with
  a no-training guarantee, a Bedrock/Vertex deployment, or a self-hosted open model?
- What is the data-classification threshold above which material may not leave the
  network at all? The Ollama path exists for exactly this case, at some quality cost.
- Who owns the approval, and what evidence do they need (data flow, retention,
  region)?
- Until that exists, the cloud app should be treated as a **demonstrator on public
  data**, not a production tool.
