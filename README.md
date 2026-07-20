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
4. [The engine — orchestrator, agents, tools, skills](#4-the-engine--orchestrator-agents-tools-skills)
5. [The pipeline, end to end](#5-the-pipeline-end-to-end)
6. [Column reference — how every cell is produced](#6-column-reference--how-every-cell-is-produced)
7. [Table detection, extraction and mapping (PDF)](#7-table-detection-extraction-and-mapping-pdf)
8. [Calculation rules](#8-calculation-rules)
9. [The Location column — how it is generated](#9-the-location-column--how-it-is-generated)
10. [Map generation](#10-map-generation)
11. [Online search rules](#11-online-search-rules)
12. [Investment rationale rules](#12-investment-rationale-rules)
13. [Source audit](#13-source-audit)
14. [Word output format](#14-word-output-format)
15. [Configuration reference](#15-configuration-reference)
16. [Deployment](#16-deployment)
17. [Verifying a change](#17-verifying-a-change)
18. [Known limits and review notes](#18-known-limits-and-review-notes)
19. [Technical limitations — where we need help](#19-technical-limitations--where-we-need-help)

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

## 4. The engine — orchestrator, agents, tools, skills

Sections 4–12 describe what the pipeline *produces*. This section describes what
*drives* it. Four layers, with one rule deciding which is used where:

> **An agent where the path is uncertain; a deterministic tool where the path is
> known.**

That rule is why there is no general "AI agent" looping over the whole deal. Routing a
`.xlsx` to the Excel scanner is a known path — a rule does it. Deciding whether a
messy broker PDF actually yielded usable comps is uncertain — an agent does that.

| Layer | Owns | Decides |
|---|---|---|
| **Orchestrator** | `orchestrator.py` | Which agent runs which task with which tool, in what order |
| **Agents** | `comp_classifier.py`, `comp_acquisition_agent.py`, the rationale writer | The uncertain judgment inside one step |
| **Tools** | `backend/tools/*.py` | Deterministic work — maths, parsing, geo, I/O |
| **Skills** | `backend/skills/*/skill.md` | Written specs for each capability (see the caveat below) |

### 4.1 Orchestrator — `orchestrator.py` (148 lines)

Rule-based, not an LLM. Given a deal config and the inputs present, it returns an
explicit, auditable **plan**: which agent performs which task with which tool, in what
order. It is pure and plannable — no Streamlit, no subprocess — so the frontend calls
`build_plan()` both to *show* the plan and to *drive* execution.

It is deliberately not an LLM picking steps: the routing is a known path
(file type → scan tool, reports → rationale). The orchestrator names the agents and
tools; **the agents own the judgment**.

### 4.2 Agents — narrow, bounded, three of them

**`comp_classifier.py` (314 lines) — "what type(s) is this file?"**
Deterministic-first, and **multi-label**: one broker PDF may hold both a land table and
a sales table, and is routed to *each* matching scan (each type's reject-markers keep
the tables apart). A file that reads like market research rather than a comp table is
flagged `is_report` so the UI nudges it to the Market Reports box. This is what lets a
user drop every file into one box.

**`comp_acquisition_agent.py` (312 lines) — "did it work, and what next?"**
A bounded **acquire → verify → evaluate → reflect → fallback** loop over the
deterministic scan scripts. The frontend runs the scripts as tools; the agent verifies
the result, scores quality, and on failure reflects to pick a fallback (e.g. file scan
came back empty → try online search). Its rules:

- **verify/flag only** — it never invents or "corrects" a number
- **deterministic scoring**; the LLM is used *only* for the reflection step
- **bounded and auditable** — every decision returns a small typed dict

**Rationale writer — `generate_investment_rationale.py`** — the prose itself (§12).

### 4.3 Tools — `backend/tools/` (14 modules)

The deterministic library. Every scan and search script imports from here; nothing
forks its own copy.

| Tool | Responsibility |
|---|---|
| `house_rules.py` | Comp-search policy — the single source of truth (§11) |
| `calculations.py` | `bala_factor`, `parse_cap_rate`, `haversine_km`, `find_same_building` |
| `column_mapper.py` | 3-tier input-header → field mapping + unit multipliers |
| `location_score.py` | The Location column for SG comps (§9) |
| `ura_landuse.py`, `ura_zone.py` | URA Master Plan lookups, local — no network at scoring time |
| `excel_reader.py` | Sheet detection, header finding, cell parsing |
| `llm_client.py` | Ollama wrappers + the agent loop (`run_agent_loop`, `apply_refinement`) |
| `vision_llm.py` | Image → comp records |
| `json_utils.py` | JSON repair for imperfect model output |
| `geo_utils.py` | Geo sidecar writer (excludes credentials) |
| `onemap_auth.py` | OneMap token handling |
| `report_period.py` | Report period / fiscal-quarter parsing |
| `corp_ssl.py` | Corporate SSL interception handling |

### 4.4 Skills — `backend/skills/*/skill.md` (19 specs)

Each capability has a Markdown spec with YAML front-matter:

```yaml
name: classify_sales_comps
description: Classify asset sales comp records with Ollama — assigns location,
             quality, asset type, and relevance score; falls back to keyword rules
type: atomic
requires:
  config_keys: [llm.ollama.base_url, subject_property.asset_class, ...]
  skills: []
allowed_tools:
  - tools.llm_client.ollama_post
  - tools.json_utils.fix_json
```

...followed by **When to use** and numbered **Instructions**.

The 19 skills are `analyse_*_comps` and `search_online_*_comps` (composite, one per
comp type), `classify_*_comps`, `parse_input_excel`, `extract_comps_from_pdf`,
`extract_comps_from_image`, `compute_sales_metrics`, `bala_factor`,
`generate_bala_table`, `geocode_and_map`, `generate_investment_rationale`,
`setup_new_deal`, and `fix_json`.

> **Note:** `skills/` is a specification layer — no Python loads it at runtime. Each
> file states the contract a capability honours: its inputs, its permitted tools, its
> steps. The behaviour that ships lives in the `.py` modules, so read a `skill.md` as
> intent and the module as the implementation.

---

## 5. The pipeline, end to end

### Stage 0 — New deal setup

`backend/new_deal.py`. The user supplies deal name, address, asset class, GFA, price,
cap rate, tenure. An LLM derives `country_name`, `currency`, `gfa_unit`,
`land_zoning`, `location`, `submarket_keywords`, and `asset_search_keyword`. Written
to `configs/deal_config_<Deal>.json`.

Deal configs **do not** carry comp-search settings — those are house rules (§11). A deal
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

**B. Online path** (`search_online_*.py`) — see §11.

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

## 6. Column reference — how every cell is produced

Legend: **Mapped** = taken from the source. **Calculated** = deterministic Python or a
live Excel formula. **Generated** = produced by a model or a scoring rule.

### 6.1 Transaction Comparables — Asset Sales

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
| 13 | Location | `location` | **Generated** | SG: URA proximity score → `Superior` / `Comparable` / `Inferior` (§9). Non-SG: LLM classification |
| 14 | Quality | `quality` | Generated | LLM. Office: `Grade A+/A/B`. Logistics: `Grade A (Modern)` / `Cold Storage` / … |
| 15 | Asset Type | `asset_type` | Generated | LLM, `"<Sale Structure> (<Use>)"` e.g. `Block Sale (Office)` |
| — | Stake % | `stake_pct` | Mapped | Hidden column, one past the last visible. Feeds the psf calculation |

### 6.2 Rent Comparables

Schema: `backend/generate_rent_comps_table.py` → `RENT_SCHEMA_BASE`

| # | Column | Field | Origin | Rule |
|---|---|---|---|---|
| 1–4 | Type / Source / Property / Map Marker | — | — | As §6.1 |
| 5 | Date of Lease Start | `lease_date` | Mapped | Free text as reported |
| 6 | Leased GLA (SF) | `nla_sf` | Mapped | Required — a record without it is dropped at extraction |
| 7 | Lease Tenure (Yrs) | `lease_term_yrs` | Mapped | Drives the effective-rent calculation |
| 8 | Gross Face Rents (SGD psf pm) | `asking_rent` | Mapped | Required (or `eff_rent`) |
| 9 | Effective Rents (SGD psf pm) | `eff_rent` | **Calculated** | **Reported first.** Else `compute_eff_rent()` amortises the rent-free period over the lease term |
| 10 | Location | `location` | Generated | §9 |
| 11 | Quality | `quality` | Generated | LLM |
| 12 | Tenant | `tenant` | Mapped | As reported |
| 13 | Type of Lease Area / Comments | `lease_type` | Mapped/Generated | Source text, else LLM label (`Whole Floor (Office)`) |

### 6.3 Land Sale Comparables

Schema: `backend/generate_land_comps_table.py` → `LAND_SCHEMA_BASE`

| # | Column | Field | Origin | Rule |
|---|---|---|---|---|
| 1–3 | Type / Source / Property | — | — | As §6.1 |
| 4 | Map Marker | `map_marker` | Generated | 1-based index after the **distance** sort (not relevance) — land comps rank by proximity |
| 5 | Date of Launch | `launch_date` | Mapped | Tender launch, not award |
| 6 | Land Zoning | `land_zoning` | Mapped | As reported; falls back to the subject's zoning |
| 7 | Land Tenure (Y) | `tenure_yrs` | Calculated | As §6.1 #7; `999` → `FH` |
| 8 | Site Area (SF) | `site_area_sf` | Mapped | Land area, not GFA |
| 9 | Max GFA (SF) | `max_gfa_sf` | Mapped | Permissible GFA, i.e. site area × plot ratio when reported that way |
| 10 | Price (SGD M) | `price_sgd_m` | Mapped | Tender/award price |
| 11 | Price (SGD psf ppr) | `price_psf_ppr` | **Calculated** | **Reported first.** Else `price ÷ max_gfa` (per plot ratio, not site area) |
| 12 | Adj. Price (SGD psf ppr) | `adj_price_psf` | **Calculated** | Bala-adjusted to the subject's tenure |
| 13–15 | Location / Quality / Comment | | Generated | §9 + LLM |

**Land tables are excluded from asset-sales extraction.** A GLS or land table appearing
in a broker PDF must not populate the asset-sales comp set; the extractor filters on
table semantics, not just keywords.

---

## 7. Table detection, extraction and mapping (PDF)

Section 5 says which *column* each value lands in. This section says how a table is
found in a PDF at all — the stage to look at when an extraction returns fewer comps
than the report contains. All of it lives in `backend/pdf_extractor.py`.

### 7.1 Stage A — page discovery (`find_relevant_pages`)

Every page's text is scanned for section-heading phrases (`Key Sales Transactions`,
`Investment Sales`, `Notable Transactions`, …). Two tiers:

1. **Keyword** — the phrase appears in the page text.
2. **Embedding** — cosine similarity ≥ `0.60` against the keyword corpus, so a
   semantically equivalent heading ("Headline Deals" ≈ "Key Transactions") is caught
   without an exact substring. Applied **only to pages that already contain a table**,
   to limit false positives.

Returns `{page_num, section_title, matched_keywords, has_table, text_preview}`. Pages
that match nothing are never opened again — so a heading this stage misses is a table
the pipeline can never find.

### 7.2 Stage B — table detection engines

Tried in order, per page, until one yields tables:

| Engine | When | Notes |
|---|---|---|
| **camelot** | first | Region-based. Detects ruled sections and extracts each |
| **pdfplumber** line tables | camelot yields nothing usable | Line/border-based |
| **img2table + easyocr** | both above fail | OCR — for scanned or image-only pages |

**The multi-column case.** Camelot is asked for page *regions*. On a multi-column
layout a region can span the whole page, returning one blob containing the title, the
stats table, the lease table, the contacts column and the transaction table together.
It returns a plausible grid rather than an error, so Stage C repairs the shape.

### 7.3 Stage C — table repair

Raw engine output is rarely a clean grid. Applied in order:

| Repair | Fixes |
|---|---|
| `_split_at_internal_headers` | **A comp header buried inside a blob starts a new table.** Requires the row to look like column headers *and* name both a name-like and a price-like column, so ALL-CAPS data cannot split a healthy table |
| `_merge_h_fragments` | A table horizontally shredded into fragments |
| `_collapse_multirow_header` | A header spread over 2–6 rows collapsed into one |
| Title-row promotion | A single-cell title row (`RECENT KEY LEASE TRANSACTIONS`) is dropped and the next row promoted to headers |
| `_orphaned_hdr` | A header row separated from its data rows |
| `_is_unit_subtitle_row` | A units row (`S$M`, `psf`) under the header, not data |
| `_split_collapsed_price_cells` | Two prices collapsed into one cell |
| `_merge_transaction_cont_rows` | A property name wrapped across rows re-joined to its transaction |

> **Example — a two-column MarketBeat page.** Camelot can return such a page as a
> single blob: title, statistics, leases, the contacts column and the transaction table
> in one grid. Its first row is page furniture (`['MARKET STATISTICS','','','OFFICE Q1
> 2025']`), while the real header — `['PROPERTY','SUBMARKET','SELLER/BUYER','PRICE
> (S$M)']` — sits far down the blob with its transactions beneath it. The table is
> present and correctly columnised, but only `_split_at_internal_headers` recovers it as
> a table.

### 7.4 Stage D — filtering (what is refused)

| Filter | Refuses |
|---|---|
| `reject_table_headers` | Out-of-scope tables by header marker. Asset-sales runs pass the GLS/land markers (`successful tender`, `psf ppr`, `per plot ratio`, …) so a land table is refused |
| `_is_prose_table` | A block of prose the engine mistook for a grid |
| `_has_header_row` | A grid with no identifiable header |
| Summary-table filtering | Market-statistics and total rows (`CBD GRADE A TOTAL`) |

A **deliberately rejected** table sets `_had_rejected_table`, which **suppresses the
text fallback** for that page. Without it, a land table correctly refused as a grid
would be re-mined straight out of the page's prose and leak back in — undoing the
rejection silently.

### 7.5 Stage E — the fallback gate

Per page: `found_any` flips true as soon as **any** table is appended. If it is true,
neither img2table nor the LLM text path runs.

The gate is a count, so a page yielding only unusable grids reads as handled and the
text fallback does not run. Stage C's repairs (§7.3) are what ensure a real comp table
is recognised before the gate is reached — which is why detection quality, not the
gate, is where extraction work belongs.

### 7.6 Stage F — column mapping (`_map_cols`)

Table headers → schema fields via `tools/column_mapper.py`, three tiers, LLM last:
**exact synonym → embedding similarity → LLM**. `detect_unit_multiplier()` reads the
header text for units (sqm→SF, S$000→S$M, psm→psf) and returns a per-field multiplier,
so values normalise without a second pass.

### 7.7 Stage G — record assembly and provenance

Two paths, and **the difference is the audit trail**:

| Path | When | Provenance |
|---|---|---|
| `_from_table` | a grid survived | **`_prov`** per field — table, row, col, header, cell. Every value was *read out of a cell* |
| `_from_text` | no grid on this page | **`_prov: null`** + `_llm_parsed`. The LLM decided where each value starts and ends, reading unstructured text |

`_from_text` records surface in the UI as the AI-judgment notice. That is the honest
signal: a grid is evidence, a sentence is an inference. `_from_text` applies the same
`reject_table_headers` check, so an out-of-scope table cannot re-enter through prose.

### 7.8 Stage H — record qualification

A row must survive all of:

| Check | Drops |
|---|---|
| `_is_real_candidate` | Rows with no name, or no price value |
| `_is_category_label` | `Office`, `Retail` — a section label, not a property |
| `_is_sentence_fragment` | Prose caught as a name |
| `_skip_subject` | The subject property itself |

### 7.9 Debugging a table that will not extract

In order:

1. **Did Stage A find the page?** The run log prints `Page N: '<title>' keywords=[…]`.
   No line ⇒ a heading problem, not a table problem.
2. **What did Stage B return?** `Page N: X table(s) found`, then per table
   `table i: N data rows, headers=[…]`. Headers that read like page furniture
   (`MARKET STATISTICS`, `OFFICE Q1 2025`) mean a blob was mistaken for a table.
3. **Was it refused?** Look for `skipping table (header matched reject list)` or
   `only out-of-scope table(s) found`.
4. **Did the fallback fire?** `camelot found no tables — trying img2table`. If absent
   while your table is missing, the gate was satisfied by a junk grid (§7.5).
5. **Reproduce without the UI or an LLM:**

```python
from scan_input_sales_comps import _PDF_SECTION_KEYWORDS
from pdf_extractor import find_relevant_pages, extract_page_tables
pages = find_relevant_pages("report.pdf", _PDF_SECTION_KEYWORDS)
for t in extract_page_tables("report.pdf", pages):
    print(t["page_num"], t["headers"], len(t["rows"]))
```

Because `extract_page_tables` makes no LLM call, this is free and deterministic — the
fastest way to tell a *detection* bug from a *model* bug.

---

## 8. Calculation rules

### 8.1 Precedence — applies to every computed cell

```
1. Reported directly by the source   → use it as-is
2. Not reported, inputs present      → calculate
3. Inputs missing                    → "—"
```

**Never `0`.** `0` is a measurement; a blank is an absence. Writing `0` for "not
reported" corrupts the average row and silently understates a comp.

Both halves matter. A source's own reported figure must never be overwritten by a
calculated one, and a calculated figure must never be invented from a missing input.

### 8.2 Cap rates

Stored as **fractions** (`0.045`), never percentages, because Excel cells use the
`0.00%` format. `parse_cap_rate()` normalises: a value `≥ 1` is divided by 100.

`parse_num()` alone strips `%` **without** rescaling — `"4.5%"` would become `4.5` and
display as `450.00%`. Always use `parse_cap_rate()` for a rate.

### 8.3 Tenure and the freehold convention

| Value | Meaning | Displays |
|---|---|---|
| `999` | Freehold | `FH` |
| `1`–`998` | Remaining years | `77 yrs` |
| `0` | **Unknown / not reported** | `—` |
| `None` | Not reported | `—` |

`0` is *not* freehold — treating it as freehold would silently convert every unknown
tenure into a freehold comp and flatter the adjusted cap rate.

### 8.4 Bala Table (Singapore leasehold adjustment)

`bala_factor(n)` in `tools/calculations.py`, from the SLA/SISV table:

| n | Factor |
|---|---|
| `n ≤ 0` or `n ≥ 999` | `1.0` (freehold) |
| `1 ≤ n ≤ 99` | SLA/SISV lookup table |
| `100 ≤ n ≤ 998` | Linear interpolation, 96% @ 99 yrs → 100% @ 999 yrs |

In Excel it is a live `VLOOKUP` against the `Bala Tbl` sheet (`bala_expr()`), so a
reviewer can change a tenure and see the adjusted cap rate move.

Singapore only. Global deals carry the FTM cap rate through unadjusted.

### 8.5 Preview and export number formatting

One shared formatter, `_fmt_grid_val(cell, header)`:

- Map Marker → plain integer (`1`, not `1.0`); subject `★` passes through
- Percentage cells (detected from the cell's own `number_format`) → `xx.00%`
- Every other number → `xxx,xxx.0` (thousands separator, exactly one decimal)
- Text → unchanged

There are **two** grid readers — `_read_excel_preview` (detail page) and
`_read_pgim_grid` (Overview page + Word export). Any new column, notice, or format must
be wired into **both**, or it will appear in one place and not the other.

---

## 9. The Location column — how it is generated

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

## 10. Map generation

`backend/generate_comps_map_base.py` → `render_map()`, wrapped per comp type.

**Geocoding and rendering are separate concerns with separate providers.** Google
resolves the coordinates; Mapbox draws the PNG. They share no credential and neither
falls back to the other.

### 10.1 Geocoding providers — Google

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

### 10.2 Rendering — Mapbox

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

## 11. Online search rules

Policy lives in **`backend/tools/house_rules.py`** — one file, applied to every deal,
existing and new, local and cloud. Deal configs do not carry these settings.

```
Precedence:  HOUSE_RULES  →  BY_ASSET_CLASS  →  the deal's own search block
```

A deal that genuinely needs different numbers sets the key in its own `online_search` /
`rent_search` / `land_search` block, and that wins. Config always beats code: nothing in
the module overrides a value a deal explicitly states.

### 11.1 The location ladder

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

### 11.2 Recency — independent of the ladder

| Comp type | `recency_months` |
|---|---|
| Sales | 60 (5 years) |
| Land | 60 (5 years) |
| Rent | **36 (3 years)** — rental evidence dates faster than capital evidence |

Applied identically to web search and grounded connectors — one cap per run.
Unparseable dates are **kept**, not dropped; every drop is logged.

Widening the search *area* never widens the *date window*.

### 11.3 `years_back` vs `recency_months` — different things

They act at opposite ends of the pipeline and nothing links them:

- **`years_back` shapes the query.** It builds the query string — `_year_window(2)` →
  `(2026 OR 2025 OR 2024)`. It is what the search is **asked for**.
- **`recency_months` filters the results.** Anything older is dropped after extraction.
  It is what is **kept**.

Setting `years_back_max` past `recency_months / 12` therefore buys rows that are then
discarded. `warn_window_vs_recency()` reports that conflict in the run log and
deliberately does **not** silently change either number.

### 11.4 Cost budget

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

### 11.5 Result limits

`max_results` = **15** per category, applied **after** classification so the cap keeps
the most relevant comps (nearest, for land).

### 11.6 Cross-source dedup

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

### 11.7 Grounded connectors

Beyond web search, `sources/registry.py` supplies keyless registries — SG URA PMI and
URA GLS via data.gov.sg, plus broker reports. Enable per deal with
`online_search.sources: ["web_search", "ura_pmi"]`. They flow through the same
dedup → geocode → recency pipeline, capped at the city tier.

---

## 12. Investment rationale rules

`backend/generate_investment_rationale.py`. Two LLM calls: prose, then audit.

### 12.1 Structure — exactly four sections

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

### 12.2 Integrity rules

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

### 12.3 Location context

One `gpt-4o-mini-search-preview` call per run asks what published sources say about the
subject's connectivity and precinct. It is **qualitative by construction**: the prompt
forbids stating a distance or walking time unless a source explicitly gives that figure.
"Directly connected to Raffles Place MRT, in the prime CBD" is allowed; "0.4 km from the
station" is not, unless cited.

If nothing is found, the block is omitted and section 2 falls back to demand drivers
rather than asserting anything unsourced. Claims that match this block are cited to
their source URL in the audit with citation type `Web Search`.

### 12.4 Extraction and caching

`pypdf` reads each page with a `[PAGE N]` marker. Text is truncated at 14,000 chars
keeping the **first 75% + last 25%** — executive summary and conclusions, dropping the
middle. Results cache on `filename + size + mtime`, so unchanged reports re-run
instantly.

---

## 13. Source audit

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

## 14. Word output format

`_build_combined_docx()` in `frontend/app.py` produces one document per deal.

### 14.1 Page setup

| Property | Value |
|---|---|
| Orientation | **Landscape** |
| Page size | US Letter, 27.94 cm × 21.59 cm (width/height swapped manually — python-docx does not swap them for you) |
| Margins | 0.5" all sides |
| Body font | **Arial Narrow 10 pt**, forced across the whole document |
| Section headings | Arial 11 pt, navy |

### 14.2 Document order

1. Deal name (Heading 0) + address
2. For each comp type present — **Rent → Sales → Land** (`_COMP_TYPES` order):
   - Section heading
   - PGIM-standard comp table
   - Location map PNG, scaled to fit the usable page box
3. Investment rationale prose

Comp types with no generated workbook are skipped silently; the document is built from
whatever exists.

### 14.3 The PGIM comp table

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
  on-screen preview exactly (§8.5)

---

## 15. Configuration reference

### 15.1 `configs/shared_settings.json` — **secrets, never distribute**

Git-ignored. Contains `mapbox_token`, `google_maps_key`, `openai_api_key`,
`kakao_api_key`, `onemap_email`, `onemap_password`, `ura_access_key`, and
`geocoding_provider`.

On Streamlit Cloud these come from Streamlit Secrets; `_bootstrap_cloud_secrets()` merges
them into this file at startup.

### 15.2 `configs/deal_config_<Deal>.json`

| Block | Purpose |
|---|---|
| `subject_property` | Name, address, asset class, GFA, price, cap rate, tenure, `country_name`, `currency`, `location`, `submarket_keywords`, `asset_search_keyword` |
| `country_code` | Explicit ISO code for geocoding. No heuristic fallback |
| `parameters` | `bala_yield` (default 0.06), `max_comps` |
| `openai` | `search_model`, `extract_model` |
| `mapbox` | `style`, `width`, `height`, `padding`, `pin_size` |
| `online_search` / `rent_search` / `land_search` | **Normally empty.** Only for per-deal overrides of a house rule, or `sources: [...]` to enable grounded connectors |
| `output_file` | Drives the output directory |

### 15.3 `backend/tools/house_rules.py` — comp-search policy

`HOUSE_RULES` (radii, `min_results`, `max_results`, `max_queries`, `years_back*`,
`max_level`), `RECENCY_MONTHS` per comp type, `BY_ASSET_CLASS` radius overrides. Change a
number here and every deal picks it up on the next run.

---

## 16. Deployment

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

## 17. Verifying a change

Read this before changing anything in `pdf_extractor.py`, the scan modules, or the
table writers. Extraction changes do not raise when they go wrong — they simply return
fewer comps — so a change is not "done" until its output has been diffed against a
known-good run.

### 17.1 Reports where a comp table is detected

**Read this table as "a grid was found," not "the values are correct."** It comes from
the no-LLM detection probe (§17.2) — the report's comp table is located and its rows
land in a grid shape. It does **not** mean the values were mapped to the right fields,
survived qualification, or were checked against the source page. Detection is a
necessary condition for a correct comp, not a sufficient one.

Only **one** report in this corpus has been checked at that deeper level — its rows
compared by hand against the source page: `singapore-office-mb-1q2025.pdf` (§7.3's
worked example). Everything else below is detection-only.

| Publisher | Series | Quarters — grid detected |
|---|---|---|
| **Cushman & Wakefield** MarketBeat | Capital Markets | 4Q2023, 1Q–2Q2024, 1Q–3Q2025, 4Q2025, 1Q2026 |
| | Office | 1Q2025, 2Q2025, 4Q2025 |
| | Industrial | 2Q2025, 3Q2025, 4Q2025 |
| **Savills** Sales & Investment Briefing | — | Q3–Q4 2023, Q2–Q4 2024, Q1–Q4 2025, Q1 2026 |
| **Colliers** Investment Report / Outlook | — | Q3–Q4 2023, Q1/Q3/Q4 2024, Q1–Q2 2025 |
| | Industrial Insights | Q1 2026 |
| **CBRE** Figures | — | Q1 2026 |

Coverage is by report, not by publisher: a series whose grid is detected in one quarter
can change layout in the next, so a new quarter needs its own check rather than an
assumption from the publisher name. Detection changing without warning is exactly what
happened to `singapore-office-mb-1q2025.pdf` before the fix in §7.3 — that report would
have appeared in a table like this one while silently returning zero real transactions.

What is outside this surface, and should be treated as unproven:

- **Market.** Singapore only. Korea and Japan deal configs exist and the search rules
  are country-agnostic, but no KR/JP broker PDF has been run — their layouts are
  untested.
- **Comp type.** The surface above is the **sales** path. Rent and land run the same
  four stages with far less coverage across formats.
- **Publisher.** JLL, Knight Frank and Edmund Tie have not been run. A new publisher
  means a new page layout, which is where Stage B/C behaviour differs (§7.2, §7.3).
- **Language.** English reports only.

A new publisher, market or quarter is therefore where work is most likely needed. Start
with the detection probe below to see what a page actually yields before assuming a
mapping problem.

### 17.2 Detection without an LLM (free, deterministic)

`extract_page_tables` makes no model call, so table *detection* can be checked
instantly and repeatably (see §7.9):

```python
from scan_input_sales_comps import _PDF_SECTION_KEYWORDS
from pdf_extractor import find_relevant_pages, extract_page_tables
pages = find_relevant_pages("report.pdf", _PDF_SECTION_KEYWORDS)
for t in extract_page_tables("report.pdf", pages):
    print(t["page_num"], t["headers"][:4], len(t["rows"]))
```

This is the fastest way to tell a **detection** bug from a **model** bug.

### 17.3 The baseline diff — the regression check that matters

Before changing detection, snapshot every PDF;
afterwards, re-run and diff. **Any file you did not intend to change must be
byte-identical.**

```python
# capture: headers + row counts per table, per PDF, for every sales-keyword PDF
# (no LLM calls — extract_page_tables only), then re-run after the change and diff.
```

A detection change should alter only the files it targets. One that also moves
unrelated files is matching too broadly and needs a tighter condition.

### 17.4 House rules a change must not break

- Every computed cell: **reported → calculated → `—`**, never `0` (§8.1)
- Cap rates are **fractions**; use `parse_cap_rate()`, never bare `parse_num()` (§8.2)
- Tenure: `999` = freehold → `FH`; `0` = **unknown**, not freehold (§8.3)
- **Two grid readers** exist (`_read_excel_preview`, `_read_pgim_grid`). Any new column,
  notice or format must be wired into **both**, or it appears in the preview and not the
  Word export — or vice versa (§8.5)
- Trust **tables**, not prose. `_from_text` records must keep `_prov: null` and their
  `_llm_parsed` flag — never launder an inference into a fact (§7.7)
- `skills/*.md` are specs, not code. Change the module; the spec can drift (§4.4)

---

## 18. Known limits and review notes

Ranked by what a reviewer should look at first.

1. **Credentials in `configs/`.** Deal configs no longer carry a map credential (the
   Mapbox token moved to Shared Settings when geocoding split to Google), but ~800
   `tmp*.json` left by a failed cleanup path still each contain the old token, and the
   token remains in the local git history. `shared_settings.json` and `tmp*.json` are
   git-ignored and the cloud repo holds no `configs/` at all — so **GitHub is clean**.
   The exposure is only if this folder is copied or zipped. **Share via the cloud repo,
   or clear `configs/tmp*.json` and rotate the token first.** The temp files are safe to
   delete; the generating path should clean up after itself.
2. **The city tier is a radius, not a boundary** (§11.1). Documented, not fixed — fixing it
   needs a locality field the geocoder does not return.
3. **The query budget can bind before the ladder finishes** (§11.4). On a thin deal, 5
   queries may be spent before tier 3 runs, so a short comp set may reflect the budget
   rather than the market.
4. **LLM classification is nondeterministic.** Two identical runs have returned different
   property names from the same PDF. This is why extraction is table-first and
   prose-derived records are flagged.
5. **Rationale quality depends on the market reports supplied.** With no reports the memo
   has little to anchor to, and the integrity rules will suppress rather than invent.
6. **Bala adjustment is Singapore-only.** Global deals carry the FTM cap rate through
   unadjusted; confirm this is intended for non-SG reviews.
7. **Location scoring is Singapore-only** and depends on the URA cache being present.

### Principles a reviewer should hold the code to

- Trust **tables**, not prose. A grid is evidence; a sentence is an inference.
- Prefer **omission over fabrication**. Blank beats a plausible guess.
- **No unit conversion** and no dropping of qualifiers on extraction.
- Every computed cell: **reported → calculated → `—`**, never `0`.
- The model classifies and writes; **Python computes**.

---

## 19. Technical limitations

Four ceilings of the current design. Each needs a decision or a capability the project
does not have today. Reviewer input wanted on all four.

**1. Extraction accuracy — can a model be trained for this?**
Comps arrive as broker PDFs, headerless "tables" of floating text, bespoke Excel
sheets, and screenshots. Today a general-purpose model that has never seen our schema
reads them, so accuracy is capped by prompt engineering and the LLM tier is
nondeterministic — identical runs have returned different property names. Could a model
be fine-tuned on our own labelled comp tables (every past deal's input file plus its
approved output is a training pair we already own) to make reading and mapping an
unfamiliar source reliable rather than best-effort? Open: is it worth the cost versus
better prompts, how do we evaluate it, and who signs off that it is safe for IC-facing
numbers?

**2. Deployment and shared memory across users**
The app is single-user by construction: Streamlit state is per-session and the only
durable store is the server's filesystem. Two analysts on the same deal silently
overwrite each other's outputs. Can this be a proper web app with a shared persistent
store, so a team works one deal together? Open: what backs the state, authentication
and who may see which deal, locking or merge semantics, and an audit trail of who
changed which cell.

**3. Narrative generation**
The rationale is one prose call with guard-rails — no revision loop, and each section
written independently. It cannot weigh conflicting sources or reason about the deal;
it assembles what the research states. Can an LLM do materially better here — a
draft → critique → revise loop, the comp tables handed over as structured evidence
rather than prose, reviewer edits fed back as house-style examples, or a stronger model
for the prose call only? And where is the honest limit at which judgement should stay
with the analyst?

**4. Personal LLM account — internal data cannot be uploaded**
The most restrictive limit. The app runs on a **personal OpenAI account**, so internal
or confidential deal material must not be uploaded and every cloud run is limited to
demo or public data. The tool cannot be pointed at the material it would be most useful
on: real offering memoranda and internal underwriting. An approved enterprise
arrangement (enterprise/Azure OpenAI with a no-training guarantee, Bedrock/Vertex, or a
self-hosted model) is a **prerequisite for real use, not an optimisation**. Until then,
treat the cloud app as a demonstrator on public data.
