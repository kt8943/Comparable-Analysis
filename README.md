# PGIM Deal Analysis Platform

An AI-powered deal analysis platform for institutional real estate underwriting. Covers three workflows — **Comparable Analysis** (asset sales, land sales, rent comps), **Investment Rationale** generation, and **New Deal** setup — all driven by a per-deal JSON config and a Streamlit dashboard.

---

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Generate the Bala Table Excel (run once per machine)
python backend/generate_bala_table_excel.py

# 3. Start Ollama (local LLM — required for analysis)
ollama serve

# 4. Pull recommended models (first time only)
ollama pull deepseek-r1:7b    # best for investment reasoning
ollama pull qwen2.5:3b        # fast, lightweight tasks

# 5. Launch the dashboard
streamlit run frontend/app.py

# 6. Or use the interactive CLI launcher
python3 run.py
```

> **OpenAI (optional):** Set your API key in **⚙️ Shared Settings** in the sidebar, or via `export OPENAI_API_KEY="sk-..."` before launching. Required for GPT models and AI-powered online comp search. The GPT-4o vision path also requires `pip install pymupdf`.
>
> **Windows note:** If Streamlit does not open automatically, create `.streamlit/config.toml` in the project root with `[browser]\ngatherUsageStats = false` and `headless = true`, then open `http://localhost:8501` manually.

---

## Repository Layout

```text
PGIM/
├── frontend/
│   └── app.py                           # Streamlit dashboard (entry point)
│
├── backend/
│   ├── new_deal.py                      # New deal wizard (LLM-assisted)
│   ├── generate_investment_rationale.py # Investment rationale pipeline
│   ├── generate_bala_table_excel.py     # One-time: PDF → Excel (Bala Table)
│   ├── pdf_extractor.py                 # Shared PDF comp extraction (pdfplumber multi-page)
│   │
│   ├── tools/                           # Shared utility library (all scan scripts import from here)
│   │   ├── calculations.py              # Pure math: haversine, bala_factor, parse_num, parse_date
│   │   ├── llm_client.py                # Ollama wrappers + agent loop (run_agent_loop, apply_refinement)
│   │   ├── excel_reader.py              # Sheet detection, header finding, cell parsing
│   │   ├── column_mapper.py             # Ollama column mapping + name-match post-correction
│   │   ├── json_utils.py                # JSON repair, array extraction
│   │   ├── vision_llm.py                # Image → comp records (wraps llm_client)
│   │   └── geo_utils.py                 # Geo sidecar writer
│   │
│   ├── scan_input_sales_comps.py        # Asset sales comps from input Excel / PDF / image
│   ├── scan_input_rent_comps.py         # Rent comps from input Excel / PDF / image
│   ├── scan_input_land_comps.py         # Land sales comps from input Excel / PDF / image
│   │
│   ├── search_online_sales_comps.py     # Asset sales comps — AI web search
│   ├── search_online_rent_comps.py      # Rent comps — AI web search
│   ├── search_online_land_comps.py      # Land sales comps — AI web search
│   │
│   ├── generate_sales_comps_table.py    # Excel schema + formatter — asset sales
│   ├── generate_rent_comps_table.py     # Excel schema + formatter — rent comps
│   ├── generate_land_comps_table.py     # Excel schema + formatter — land sales
│   │
│   ├── generate_sales_comps_map.py      # Mapbox map builder — asset sales
│   ├── generate_rent_comps_map.py       # Mapbox map builder — rent comps
│   ├── generate_land_comps_map.py       # Mapbox map builder — land sales
│   └── generate_comps_map_base.py       # Shared geocoding + map rendering engine
│
├── configs/
│   └── deal_config_<DealName>.json      # One config per deal
│
├── Input_files/
│   ├── bala_table.xlsx                  # Singapore SLA/SISV Bala Table (generated)
│   ├── bala table.pdf                   # Source PDF for Bala Table
│   ├── *.xlsx                           # Manually curated comps input files
│   └── market_reports/
│       ├── *.pdf                        # Market research PDFs for rationale
│       └── cache/                       # LLM extraction cache (auto-managed)
│
├── output/
│   └── <DealName>/                      # All outputs per deal
│       ├── Transaction_Comparables_<DealName>.xlsx
│       ├── Transaction_Comparables_<DealName>_records.json
│       ├── Transaction_Comparables_<DealName>_geo.json
│       ├── Transaction_Comparables_<DealName>_map.png
│       ├── Land_Sale_Comps_<DealName>.xlsx
│       ├── Rent_Comps_<DealName>.xlsx
│       ├── Investment_Rationale.md
│       ├── Investment_Rationale_meta.json
│       └── Source_Audit.xlsx
│
├── run.py                               # Interactive CLI launcher
└── requirements.txt
```

---

## Code File Interaction Map

### How files call each other

There are three calling patterns in this codebase:

| Pattern | When used |
|---------|-----------|
| `subprocess.run()` | `app.py` launching any backend script — each analysis run is a separate Python process |
| `import` | Backend scan/search scripts pulling in table builders and map generators |
| Direct `import` inside `app.py` | Dashboard-only map regeneration (no subprocess — avoids the overhead of spawning a process just for re-drawing pins) |

---

### Workflow A — Comps from Input Excel

```
app.py
  │  (user uploads Excel, clicks Run)
  │
  └─► subprocess ──► scan_input_sales_comps.py
                          │                        (same pattern for rent / land variants)
                          │  1. Reads input Excel from deal config path
                          │  2. Calls Ollama: map input columns → OUTPUT_SCHEMA keys
                          │  3. Calls Ollama: classify location, quality, asset type per comp
                          │  4. Calls Mapbox: geocode subject property
                          │  5. Writes *_records.json  ← raw parsed comps (all fields)
                          │  6. Calls Mapbox: geocode each comp, sort by Haversine distance
                          │  7. Writes *_geo.json       ← lon/lat per comp + map settings
                          │
                          ├─► import ──► generate_sales_comps_table.py
                          │                  OUTPUT_SCHEMA, build_workbook(),
                          │                  subject_to_row(), comp_to_row(), bala_factor()
                          │                  reads:  Input_files/bala_table.xlsx
                          │                  writes: Transaction_Comparables_<DealName>.xlsx
                          │
                          └─► import ──► generate_sales_comps_map.py
                                             └─► import ──► generate_comps_map_base.py
                                                                geocode_with_fallbacks()
                                                                render_map()
                                                                calls Mapbox Static Images API
                                                                writes: *_map.png
```

**Sidecar files explained:**

| File | Written by | Read by | Purpose |
|------|-----------|---------|---------|
| `*_records.json` | `scan_input_*.py` (step 5) | `app.py` (dashboard table) | Full comp data before geocoding; used to populate the editable preview table and keep data in sync when the user edits |
| `*_geo.json` | `scan_input_*.py` (step 7) | `app.py`, `generate_*_map.py` | Geocoded lon/lat per comp, hidden flags, Mapbox render settings; single source of truth for map pin positions and visibility |

---

### Workflow B — Comps from Online Search

```
app.py
  │  (user clicks AI Search)
  │
  └─► subprocess ──► search_online_sales_comps.py
                          │                        (same pattern for rent / land variants)
                          │  1. Calls OpenAI web search (gpt-4o-mini-search-preview)
                          │     Level 1: within proximity_km of subject
                          │     Level 2: submarket fallback if < min_results found
                          │     Level 3: market-wide fallback if still short
                          │     Temporal expansion: extend years_back if still short
                          │  2. Calls Ollama: classify location, quality, asset type
                          │  3. Calls Mapbox: geocode subject + comps
                          │
                          ├─► import ──► generate_sales_comps_table.py  (same as Workflow A)
                          │
                          └─► import ──► generate_sales_comps_map.py
                                             └─► import ──► generate_comps_map_base.py
                                                                writes: *_map.png
```

Note: Online search does not write `_records.json` or `_geo.json` sidecars — it writes directly to Excel and PNG.

---

### Workflow C — Investment Rationale

```
app.py
  │  (user ticks reports, clicks Generate)
  │
  └─► subprocess ──► generate_investment_rationale.py
                          │
                          │  Stage 1 — Extraction (cached per PDF)
                          │  ─────────────────────────────────────
                          │  For each selected market report PDF:
                          │    1. pypdf reads PDF page by page, prefixes [PAGE N]
                          │    2. Smart truncation: keep first 75% + last 25% of text
                          │    3. LLM extracts structured JSON:
                          │       vacancy, rents, cap values, demand drivers, pipeline, outlook
                          │       key stats tagged with page refs: [p.5] Vacancy: 3.2%
                          │    4. Writes cache file: Input_files/market_reports/cache/<hash>.json
                          │       (hash = filename + size + mtime — unchanged PDFs served instantly)
                          │
                          │  Stage 2 — Generation (two LLM calls)
                          │  ──────────────────────────────────────
                          │  Call 1 — Prose
                          │    - Sources anonymised: "Research Report 1 / 2 / …"
                          │    - LLM: plan sections → write prose → self-verify checklist
                          │    - Writes: Investment_Rationale.md
                          │
                          └─  Call 2 — Audit
                               - Receives prose + real source filenames
                               - LLM audits every number, policy, and named claim
                               - openpyxl writes: Source_Audit.xlsx
                                 (rows needing manual PDF verification highlighted red)

app.py  (immediately after subprocess returns, on success)
  └─► writes: Investment_Rationale_meta.json
              { "generated_at": "2026-06-10 14:32", "elapsed_s": 47.1 }
              displayed as caption under the rationale heading
```

---

### Workflow D — New Deal Setup

```
app.py
  │  (user fills in deal form, clicks Generate Config)
  │
  └─► subprocess ──► new_deal.py
                          │  1. Reads optional deal brief (PDF / Excel / txt)
                          │  2. Calls LLM: derive country, currency, GFA unit,
                          │               location descriptor, submarket keywords,
                          │               market search query, land zoning, asset keyword
                          └─► writes: configs/deal_config_<DealName>.json
```

---

### Dashboard-only operations (no subprocess)

These operations happen entirely inside `app.py` — no backend script is spawned:

| User action | Files read | Files written | How |
|-------------|-----------|---------------|-----|
| Save edits to Excel | `*_records.json`, `*_geo.json` | Output `.xlsx`, `*_geo.json` | openpyxl (inline) |
| Save + geocode new row | `*_geo.json` | `*_geo.json`, `*_map.png` | inline + direct import of `generate_*_map.py` |
| Hide / restore a map pin | `*_geo.json` | `*_geo.json`, `*_map.png` | inline + direct import of `generate_*_map.py` |
| Delete pin + row | Output `.xlsx`, `*_geo.json` | Output `.xlsx`, `*_geo.json`, `*_map.png` | openpyxl (inline) + direct import |

For all map PNG regeneration triggered from the dashboard, `app.py` directly imports `generate_sales_comps_map.render_map` (or rent/land equivalent), which in turn imports `generate_comps_map_base.py`. The call chain is:

```
app.py (inline)
  └─► import generate_sales_comps_map.render_map
           └─► import generate_comps_map_base.render_map
                    calls Mapbox Static Images API
                    draws pins with Pillow
                    writes *_map.png
```

---

### Shared modules — who imports what

**`generate_comps_map_base.py`** — imported by all three map builders, never called directly:

```
generate_sales_comps_map.py  ─┐
generate_rent_comps_map.py   ─┼─► generate_comps_map_base.py
generate_land_comps_map.py   ─┘        geocode()
                                        geocode_with_fallbacks()
                                        render_map()
```

**`generate_*_table.py`** — imported by both scan and search scripts for that comp type:

```
scan_input_sales_comps.py    ─┐
search_online_sales_comps.py ─┴─► generate_sales_comps_table.py
                                       OUTPUT_SCHEMA          ← column definitions (name, format, width)
                                       build_workbook()        ← creates the formatted Excel workbook
                                       subject_to_row()        ← writes the subject property row
                                       comp_to_row()           ← writes one comp row
                                       bala_factor()           ← Singapore leasehold adjustment lookup
                                       get_output_schema()     ← adapts schema for currency / area unit
```

The same pattern applies for rent (`RENT_SCHEMA_BASE`) and land (`LAND_SCHEMA_BASE`).

**To change column names, number formats, or column widths** in the output Excel, edit `OUTPUT_SCHEMA` (or `RENT_SCHEMA_BASE` / `LAND_SCHEMA_BASE`) in the relevant table file. Each entry is:

```python
("Display Header",  "internal_key",  "dtype",  "excel_format",  col_width)
#  shown in Excel    used in code     str/int    e.g. "#,##0"     char units
#                                     /float/pct
```

---

## Full Pipeline Flow

### Overview

```text
Deal Config JSON
       │
       ├──► [New Deal] new_deal.py
       │         LLM derives fields from address / deal brief
       │
       ├──► [Comps] scan_input_*.py  or  search_online_*.py
       │         │
       │         ├── Ollama: column mapping (schema mapper prompt)
       │         ├── Ollama: location + quality classification
       │         ├── Mapbox: geocode + sort by distance
       │         └── generate_*_table.py → formatted Excel + Bala Table lookup
       │
       └──► [Rationale] generate_investment_rationale.py
                 │
                 ├── Stage 1 (per report, cached):
                 │       pypdf reads PDF page-by-page → LLM extracts structured JSON
                 │
                 └── Stage 2 (two separate LLM calls):
                         Call 1: anonymised sources → prose rationale
                         Call 2: real sources → citation audit JSON → Source_Audit.xlsx
```

---

### Stage-by-Stage Detail

#### 1. New Deal Setup

- User enters deal name, address, asset class, GFA, quality, price, cap rate
- Optionally uploads a deal brief (PDF / Excel / txt) — LLM extracts all available fields
- LLM auto-derives: country, currency, GFA unit, zoning, location descriptor, submarket keywords, market search query
- Saves to `configs/deal_config_<DealName>.json`

#### 2. Comparable Analysis (Input Excel path)

1. **Sheet detection** — scores all sheets by how many output-field keywords appear in any header row; picks the best match
2. **Header detection** — finds the first row with ≥ 3 text cells
3. **Column mapping** — Ollama maps output column names → input column indices using schema mapper prompt; a name-match post-correction pass overrides the LLM when a header unambiguously matches a single field's keywords (e.g. a column literally named "Address" is always mapped to the address field); keyword fallback fills any remaining null mappings
4. **Record qualification** — rows must have a non-empty name and at least one price value
5. **Classification** — Ollama assigns location tier and quality grade per comp; keyword-rules fallback if Ollama fails
6. **Geocoding** — Mapbox geocodes subject + each comp; sorts by Haversine distance. Strategy: if the address field looks like a real street address (contains a digit and a street-type keyword such as "Road", "Street", "Jalan"), it is used as the geocoding query; otherwise the building name is used as a fallback (labelled `(by name)` in the run log)
7. **Table render** — `generate_*_table.py` writes formatted company-template Excel with Bala Table adjustments

**PDF inputs:** `pdf_extractor.py` routes to one of two extraction paths depending on the selected model:

- **Ollama path** *(Ollama model selected)* — 4 stages: keyword page discovery → pdfplumber table detection → 3-tier rule-based column mapping (exact → keyword → fuzzy → Ollama last resort) → row assembly + deduplication.
- **GPT-4o vision path** *(GPT model selected)* — Skips Stages 1–3 entirely. `pymupdf` renders every page to an image; all images are sent in one API call. GPT-4o finds relevant sections by keyword, detects the table visually, reads all cells (including floating text that pdfplumber misses), and returns a JSON array of records directly.

The GPT-4o vision path exists because some PDFs render property names as floating text elements that sit visually inside table cells but are not registered within the cell boundary boxes — `pdfplumber.extract_tables()` returns empty strings for those cells, while GPT-4o reads what is visually on the page. All three comp types support PDF and image uploads in addition to Excel.

#### 3. Comparable Analysis (Online Search path)

1. **Level 1 (proximity)** — search within `proximity_km` radius using address + asset keywords
2. **Level 2 (submarket)** — expand to `submarket_km` if fewer than `min_results` comps found; uses `submarket_keywords`
3. **Level 3 (market-wide)** — use `broader_market_query` if still insufficient
4. **Temporal expansion** — if still short, extend lookback window by `years_back_step` up to `years_back_max`
5. Remainder same as Input Excel path from step 5 onwards

#### 4. Investment Rationale — Stage 1: Extraction (cached)

1. `pypdf` reads each PDF page-by-page, prefixing each page with `[PAGE N]`
2. First 75% + last 25% of text is kept (smart truncation at 14,000 chars) — captures executive summary and conclusions
3. LLM extracts structured JSON: market overview, supply/demand, rental trends, capital values, demand drivers, pipeline, outlook, key statistics with page tags (`[p.N] ...`)
4. Result cached by `filename + size + mtime` hash — unchanged reports served instantly on re-runs

#### 5. Investment Rationale — Stage 2: Generation (two LLM calls)

**Call 1 — Prose**

- Sources are anonymised: each report is labelled "Research Report 1 / 2 / …" so the LLM cannot echo PDF filenames into body text
- LLM follows a three-step process: (a) map data points to sections before writing, (b) write sections with evidence-dense prose, (c) self-verify against checklist before returning
- Outputs 3–5 sections; default 3

**Call 2 — Audit**

- Receives the prose + real source data (with actual filenames)
- Audits every claim: captures every number, every named policy, every named trend, every sector/location reference
- Outputs a citation JSON array: `source_file`, `page_ref`, `supporting_text`, `citation_type`
- Written to `Source_Audit.xlsx` with backend cross-check against cached extracts; rows needing manual PDF verification are highlighted red

---

## Prompting Techniques

### 1. Role Prompting

All system prompts open with a specific professional persona:

- Rationale: *"You are a senior investment professional at a global institutional real estate fund"*
- Extraction: *"You are a senior real estate research analyst at an institutional investment firm"*
- Column mapping: *"You are a data schema mapper"*
- Classification: *"You are a senior [country] commercial real estate analyst"*

This anchors the LLM's vocabulary, tone, and priorities before any task instructions.

### 2. System / User Message Separation

Hard constraints (data integrity rules, banned phrases, output format) go in the **system message** — highest precedence. Task-specific instructions and variable data go in the **user message**. This prevents the LLM from treating formatting rules as optional suggestions.

### 3. Chain-of-Thought (Explicit Multi-Step)

The rationale generation prompt forces explicit reasoning before writing:

- **STEP 1** — Map every data point to exactly one section before writing a word of prose
- **STEP 2** — Write sections using only the data assigned to each section
- **STEP 3** — Self-verify against a checklist (language, data anchor, source check, policy check, word count, attribution, transitions) and fix failures before returning

### 4. Self-Verification Checklist

The LLM is given a structured checklist it must evaluate against its own output before returning. Each item is a binary pass/fail:

- `SOURCE CHECK` — every number must appear digit-for-digit in the research or deal config
- `POLICY CHECK` — every policy name must be explicitly named in the research
- `DATA ANCHOR` — every opinion or forecast must follow a specific figure or named fact
- `REUSE LIMIT` — no single statistic appears more than 3 times across all sections

### 5. Source Anonymisation (Decoupled Writing and Auditing)

During prose generation (Call 1), reports are labelled "Research Report 1 / 2 / …" — the LLM never sees real PDF filenames and cannot echo them into body text. During the audit call (Call 2), real filenames are revealed for exact source matching. This decoupling keeps the prose clean while enabling precise citation tracking.

### 6. Two-Call Separation (Prose + Audit)

Writing and auditing are two entirely separate LLM calls with different system prompts and temperatures:

- Call 1 temperature: 0.2 (consistent, controlled writing)
- Call 2 temperature: 0.1 (deterministic fact-matching)

Combining them in one call causes the LLM to contaminate prose with citation JSON or compress its writing to fit the audit format.

### 7. Structured JSON Output with Exact Schema

All extraction and classification calls request a specific JSON schema with named keys and null conventions. The schema is included verbatim in the prompt so the LLM knows every expected field. JSON mode (`"format": "json"`) is used in Ollama calls to constrain output format.

### 8. Non-Negotiable Data Integrity Rules

The system prompt labels an entire section `━━ DATA INTEGRITY (non-negotiable) ━━`. Rules include:

- Never estimate, round, or extrapolate figures not in the research
- Never use general market knowledge — omit if not traceable to a source
- Every number must appear digit-for-digit in the Market Research Summary or Deal Config
- Every policy name must be explicitly named in the research

### 9. Smart Truncation (Front + Tail Strategy)

When a PDF exceeds the model context window (14,000 chars default), the script keeps the **first 75%** (executive summary, methodology, key findings) plus the **last 25%** (outlook, conclusions, risks) — rather than a simple head truncation. This preserves both the opening data-dense sections and the forward-looking conclusions that matter most for investment writing.

### 10. Page-Level Markers

`pypdf` extraction prefixes each page with `[PAGE N]`. This propagates through the extraction JSON into key statistics (e.g. `[p.5] Vacancy rate: 3.2% (Q1 2026)`), allowing the audit to record exact page references without the LLM needing to guess.

### 11. Temperature Control

| Use case | Temperature | Reason |
|---|---|---|
| PDF extraction (Stage 1) | 0.1 | Maximise factual fidelity |
| Rationale prose (Call 1) | 0.2 | Controlled, consistent writing |
| Citation audit (Call 2) | 0.1 | Deterministic source matching |
| Column mapping | 0.0 | Exact, repeatable schema mapping |
| Classification | 0.0 | Consistent tier assignment |

### 12. Keyword Fallback for Column Mapping

Ollama column mapping is followed by a second-pass keyword matching step. Any field Ollama left as `null` is matched against column headers using both the output display name and its description as keyword sets. Only columns not already claimed by another field are eligible — preventing false matches.

### 13. Banned Phrase Lists

The rationale system prompt explicitly bans specific phrases that degrade institutional writing quality:

- Attribution hedges: "according to", "the report states", "data shows", "it is noted that"
- Filler transitions: "additionally", "furthermore", "moreover", "in addition", "lastly", "to summarise"
- Gap-flagging: "vacancy data is not available", "specific figures are not provided", "data is limited"

---

## The Dashboard

Launch with `streamlit run frontend/app.py`. Three sections in the sidebar:

| Section | What it does |
|---------|-------------|
| 🏗️ New Deal | LLM-assisted wizard to create a new deal config |
| 📋 Comparable Analysis | Run comps from uploaded Excel or AI web search |
| ✍️ Investment Rationale | Generate a 3-section investment committee memo |

### Model Selector (sidebar)

Every analysis respects the model selected in the **LLM MODELS** panel. The dropdown only shows models actually installed in Ollama, plus GPT options always listed at the bottom. The active model is shown as a badge and applies to all backend calls.

**Text / reasoning models**

| Model | Size | Best for |
|-------|------|----------|
| `qwen2.5:3b` | 1.9 GB | ⚡ Fastest — lightweight tasks |
| `gemma3:4b` | 3.3 GB | ⚡ Fast — decent quality |
| `deepseek-r1:7b` | 4.7 GB | 🧠 Investment reasoning & analysis |
| `llama3.1:8b` | 4.9 GB | ✍️ General purpose, good writing |
| `qwen3:8b` | 5.2 GB | 🏦 Finance + Asian markets |
| `qwen3.5:9b` | ~6 GB | 🏦 Finance + Asian markets, stronger reasoning |
| `gpt-4o-mini` | Cloud ☁️ | ☁️ Fast cloud — investment rationale writing |
| `gpt-4o` | Cloud ☁️ | ☁️ Best cloud quality |

**Vision models** — required only when uploading a comp table as an image (screenshot/photo). Select in the sidebar Vision Model selector.

| Model | Size | Notes |
|-------|------|-------|
| `minicpm-v` | ~5 GB | Recommended — good table extraction |
| `llama3.2-vision` | ~8 GB | Strong quality, larger download |
| `moondream` | ~1.7 GB | Smallest, lower accuracy |

---

## Workflows

### 🏗️ New Deal Setup

Creates a deal config JSON from minimal inputs.

**Quick entry:** Fill in deal name, address, asset class, GFA, quality, sale date, price, and cap rate. Click **Generate Config Preview** — the LLM auto-derives country, currency, GFA unit, zoning, location descriptor, submarket keywords, and market search query.

**With document:** Upload a PDF / Excel / txt deal brief (or paste text). The LLM extracts all available fields before deriving the rest.

Review and edit every field in an interactive table, then save. The config is written to `configs/deal_config_<DealName>.json`.

---

### 📋 Comparable Analysis

Supports three comp types: **Asset Sales**, **Land Sales**, **Rent Comps**.

**Upload Comps** — provide your own curated Excel:
- Upload (or reference a previously configured) input Excel
- Ollama auto-detects column layout via schema mapper prompt + keyword fallback
- Ollama classifies Location, Quality, and Asset Type per comparable
- Produces a formatted Excel table + Mapbox location map

**AI Search** — find comps automatically:
- *Internal Database:* classify comps from an existing Excel in `Input_files/`
- *Online Search:* GPT web search using a proximity-first, submarket-fallback strategy (requires OpenAI API key)

**Editing results in the dashboard:**
- The preview table is editable — change values, delete rows, or add new rows directly
- **💾 Save Edits & Update** — saves display changes (name, address, date) directly to Excel and regenerates the map. Fast path — no pipeline re-run. Also syncs `_records.json` so deletions and edits persist.
- **🔄 Re-Run & Update** — saves all changes back to `_records.json` and re-runs the full pipeline (re-classification, Bala recalculation, metric updates). Use this when a price, GFA, or tenure value was corrected.
- Deleting a row automatically renumbers remaining map markers (1, 2, 3 …) and rebuilds the geo sidecar
- The Average row at the bottom is computed automatically on each save and always reflects only the comps currently in the table

**Refine This Output:**

After the table is generated, expand **🔄 Refine This Output** to filter or adjust the comp list using natural language. Extraction is skipped — only classification and table rendering re-run, so this typically takes 1–2 min.

The refinement uses an **agent loop**: the LLM dispatches to typed Python tools rather than filtering data itself, so numeric comparisons are exact and record values are never corrupted.

| Tool | Example instruction |
| ---- | ------------------- |
| `filter_numeric` | "remove comps with price > 1000M", "exclude GFA below 50,000 sf" |
| `filter_by_marker` | "remove map marker 11" |
| `filter_by_name` | "remove Woodlands Drive" |
| `filter_last_n` | "delete the last 2 rows" |
| `compute_stats` *(query)* | enables "remove outliers" — LLM fetches mean/std first, then filters |

After a successful refinement the preview table updates automatically and the run log is shown in a collapsible panel below the table.

**Location competitiveness score (Singapore):**

For Asset Sales and Rent comps, the **Location** column is scored against the subject
property by sector-specific proximity — not copied from the source text. Each comp gets
**Superior / Comparable / Inferior**:

- Score is normalised to −1…+1 with the **subject = 0**; `|s| ≤ 0.3` → Comparable,
  `> 0.3` → Superior, `< −0.3` → Inferior.
- Factors are sector-specific (office → CBD proximity + commercial density; industrial/
  data centre → business cluster + port/airport; retail → residential catchment + regional
  centre; hospitality → tourist draw + commercial density; mixed → blend).
- Uses the **map-resolved lon/lat** (Google/Mapbox) for consistency with the pin, plus the
  **local URA Master Plan** (`backend/data/MasterPlan2025.geojson`) and OneMap themes.
  Fully on-prem; the per-comp score prints in the run log.
- Only comps of the **same sector** as the subject are scored; others are left blank. Land
  comps are excluded by design.
- See [`backend/docs/location_score_methodology.md`](backend/docs/location_score_methodology.md)
  for the full formula and justification.

**Land Zoning from land use:** when a sales comp's zoning is missing from the source PDF,
it is filled by a token-free point-in-polygon lookup on the URA Master Plan at the comp's
coordinates (e.g. `COMMERCIAL → C`), and shown in both the preview and the Excel.

**Geocoding providers:** selectable in Shared Settings — **Google** (default; best for
KR/foreign addresses), **OneMap** (SG, public), **Kakao** (Korean address engine), and
**Mapbox**. The subject-property star can be toggled off on the map (comps then render red).

**Rule-based mode:** the model selector includes **🚫 Rule-based (no LLM)** — column
mapping + classification run on exact-synonym/keyword rules only (Tier 3 LLM skipped),
useful for testing OneMap/geocoding without Ollama.

---

### ✍️ Investment Rationale

Two-stage LLM pipeline that reads market research PDFs and writes a 3–5 section institutional investment rationale.

**Stage 1 — Extract (cached)**

Each market report PDF is read with page-level tracking. The LLM extracts structured market intelligence:
- Supply/demand dynamics, vacancy rates, net absorption
- Rental trends and growth forecasts
- Capital values, yields, transaction volumes
- Demand drivers, development pipeline, market outlook
- Key statistics tagged with source page numbers (e.g. `[p.5] Vacancy rate: 3.2%`)

Results are cached by file hash in `Input_files/market_reports/cache/`. Unchanged reports are served instantly on re-runs.

**Stage 2 — Generate (two separate LLM calls)**

- *Call 1 — Prose:* Writes the rationale sections using anonymised source labels, so no PDF filename ever appears in body text. LLM follows a plan → write → self-verify sequence.
- *Call 2 — Audit:* Audits every claim against the real source data, capturing every number and every named policy. Outputs a citation JSON with source file, page reference, supporting text, and citation type.

**Source Audit Citation Types:**

| Type | Meaning |
|------|---------|
| `Verbatim` | Direct quote from the source report |
| `Paraphrased` | Rewording of content from the source report |
| `Deal Config` | Drawn from the deal config JSON (property-specific facts) |

General Knowledge is not a permitted citation type — every claim must trace to a report or the deal config.

**Output files:**

| File | Description |
|------|-------------|
| `Investment_Rationale.md` | Full rationale — displayed in dashboard, downloadable as `.md` or `.docx` |
| `Investment_Rationale_meta.json` | Generation timestamp and elapsed time — displayed as caption in dashboard |
| `Source_Audit.xlsx` | Every cited claim with source PDF name, page reference, backend cross-check status, and a human validation column. Red rows = need manual PDF verification. |

**Refine without re-extracting:**

After generation, expand **🔄 Refine This Output** — describe what to change and click **Regenerate with Changes**. Extraction stays cached so only the writing re-runs (~2–4 min local model, <30 s GPT).

---

## Bala Table (Singapore Leasehold Adjustment)

The Bala Table is the official Singapore Land Authority / SISV table showing leasehold values as a percentage of freehold value (Appendix 2, SLA/SISV). It is used to adjust comparable sale prices and subject property values for remaining lease differences.

### Setup (run once per machine)

```bash
# Reads Input_files/bala table.pdf → writes Input_files/bala_table.xlsx
python backend/generate_bala_table_excel.py
```

### How it is used

| Layer | Method |
|---|---|
| Python (comps scan) | Dict lookup: `_BALA_TABLE[years]` returns leasehold % |
| Excel formula | `VLOOKUP` against a hidden 'Bala Tbl' sheet embedded in the output workbook |
| Output workbook | 'Bala Tbl' sheet contains all 99 rows (n=1 → 3.8%, n=99 → 96.0%) |

Key checkpoints: n=30 → 60.0%, n=60 → 80.0%, n=99 → 96.0%.  
Freehold / 999-year leases: Bala factor = 1.0 (no adjustment).  
Leases over 99 years: linear interpolation between 99 years (96%) and freehold (100%).

---

## Deal Config Reference

One JSON file per deal in `configs/`. Create with the New Deal wizard or copy an existing one and update all fields.

### `subject_property`

| Field | Type | Description | Example |
|-------|------|-------------|---------|
| `deal_name` | string | Short name used in filenames and headings | `"88 Cecil Street"` |
| `property_name` | string | Full building name | `"88 Cecil Street"` |
| `address` | string | Full address including city and country | `"88 Cecil Street, Singapore"` |
| `asset_class` | string | `"office"`, `"logistics"`, `"retail"`, `"industrial"`, `"mixed-use"` | `"office"` |
| `asset_type` | string | Transaction structure | `"Whole Block (Office)"` |
| `quality` | string | Building grade | `"Grade A"` |
| `gfa_sf` | int | Gross floor area in the unit set by `gfa_unit` | `88500` |
| `gfa_unit` | string | `"sf"` (sq ft) or `"sqm"` (sq metres) | `"sf"` |
| `sale_date` | string | Label for the sale date column | `"2025E (Mktg)"` |
| `remaining_leasehold_yrs` | int | Remaining lease years; `0` = freehold | `83` |
| `price_sgd_m` | float / null | Price in deal currency (millions) | `320.0` |
| `ftm_noi_cap_rate` | float / null | Forward NOI cap rate as a decimal | `0.040` |
| `location` | string | Qualitative location descriptor | `"CBD Fringe"` |
| `land_zoning` | string | Planning zoning class | `"Commercial"` |
| `country_name` | string | Full country name | `"Singapore"` |
| `country_code` | string | ISO 2-letter code for Mapbox geocoding | `"sg"` |
| `currency` | string | 3-letter currency code | `"SGD"` |
| `currency_symbol` | string | Currency symbol for display | `"S$"` |
| `asset_search_keyword` | string | Phrase used in AI search queries | `"office building"` |
| `submarket_keywords` | string[] | Submarket names for Level 2 search | `["Tanjong Pagar", "Cecil Street"]` |
| `broader_market_query` | string | Market-wide fallback search query | `"Singapore office CBD fringe investment sale"` |

### `llm`

```json
"llm": {
  "provider": "ollama",
  "ollama": {
    "base_url": "http://localhost:11434",
    "model": "deepseek-r1:7b"
  }
}
```

Set `"provider": "openai"` with `"openai_model": "gpt-4o"` to use GPT. The dashboard model selector overrides this at runtime without editing the config file.

### `openai`

```json
"openai": {
  "api_key": null,
  "search_model": "gpt-4o-mini-search-preview",
  "extract_model": "gpt-4o-mini"
}
```

Leave `api_key` as `null` to use the `OPENAI_API_KEY` environment variable.

### `online_search`

Controls the proximity-first, submarket-fallback search strategy:

| Field | Default | Description |
|-------|---------|-------------|
| `proximity_km` | 1.0 | Level 1 radius — nearest comps |
| `submarket_km` | 5.0 | Level 2 radius — submarket fallback |
| `min_results` | 3 | Min comps before escalating to next level |
| `max_results` | 10 | Hard cap on returned comps |
| `years_back` | 2 | Initial lookback window (years) |
| `years_back_max` | 8 | Maximum lookback before stopping |
| `years_back_step` | 2 | Years added per temporal expansion |

### `mapbox`

```json
"mapbox": {
  "token": "pk.eyJ1...",
  "style": "streets-v12",
  "width": 1200,
  "height": 900,
  "padding": 100,
  "pin_size": "l"
}
```

`pin_size` options: `"l"` (default, no extra dependencies), `"xl"` / `"xxl"` (requires Pillow).  
Leave `token` as `""` to skip geocoding and map generation — comps will be kept in input order.

### `parameters`

```json
"parameters": {
  "max_comps": 10
}
```

---

## Investment Rationale — Market Reports Setup

1. Place PDF market reports in `Input_files/market_reports/`
2. In the dashboard, tick the reports to include and click **Generate Investment Rationale**
3. Extraction is cached — subsequent runs with unchanged PDFs skip Stage 1

**To get page numbers in the Source Audit:** tick **♻️ Re-extract reports** once after the initial setup. Thereafter, re-extract only when a report file changes.

---

## CLI Launcher

`run.py` provides an interactive menu for all backend scripts without opening the dashboard:

```bash
python3 run.py
```

Options:
1. Asset Sales Comps — Online Search
2. Asset Sales Comps — From Input Excel
3. Rent Comps — Online Search
4. Rent Comps — From Input Excel
5. Land Sales Comps — From Input Excel
6. Land Sales Comps — Online Search
7. Investment Rationale — Generate from Market Reports
8. New Deal Setup

---

## Two-Laptop Workflow

The platform is designed to work across two machines with different access levels:

| Machine | Role | Has access to |
|---|---|---|
| **Personal laptop (Mac)** | Development + AI | Claude Code, internet, market report PDFs, code |
| **Corporate laptop (Windows)** | Data + execution | Internal deal data (comps Excel, deal documents), Ollama |

**Recommended workflow:**
1. Develop and test all code changes on the Mac
2. Copy only the `backend/` folder to Windows — no other folders need to change
3. Also copy `Input_files/bala_table.xlsx` if it was regenerated
4. Run the pipeline on Windows using internal data files

**Never copy back from Windows to Mac** unless you are intentionally syncing internal data to the Mac as well.

---

## Dependencies

| Package | Required | Purpose |
|---------|----------|---------|
| `streamlit` | ✅ | Dashboard UI |
| `pandas` | ✅ | Data handling |
| `openpyxl` | ✅ | Excel read / write |
| `pypdf` | ✅ | PDF text extraction (market reports, deal briefs, Bala Table) |
| `python-docx` | ✅ | `.docx` export of Investment Rationale |
| `camelot-py[cv]` | ✅ | PDF table extraction (lattice/stream) |
| `pdfplumber` | ✅ | Text/table fallback extraction |
| `fastembed` | ⚠️ Recommended | Tier-2 embedding column mapping. **Must be in the same Python that launches Streamlit** or it silently disables (falls back to exact-match + LLM). |
| `truststore` | ⚠️ Recommended | Trust the OS cert store behind a TLS-intercepting proxy (Zscaler). No-op otherwise. |
| `openai` | ⚠️ Optional | Online comp search + GPT analysis models |
| `pymupdf` | ⚠️ Optional | PDF → image rendering for GPT-4o vision path (`pip install pymupdf`) |
| `Pillow` | ⚠️ Optional | Map rendering for `xl`/`xxl` custom pins only |
| **Ollama** | ⚠️ Runtime | Local LLM — install from [ollama.com](https://ollama.com). Fallback only for column mapping/classification. |
| **Mapbox / Google token** | ⚠️ Runtime | Static map generation + geocoding — set in each deal config / shared settings |
| **`MasterPlan2025.geojson`** | ⚠️ Runtime data | `backend/data/` — URA land-use polygons for the SG Location score + Land-Zoning-from-land-use. Not a pip package; local/on-prem. |
| **OneMap account** | ⚠️ Runtime | Email + password in `shared_settings.json` for token auto-refresh (Themes only; public geocoding needs no token). |

```bash
pip install -r requirements.txt
```

```bash
# Install Ollama: https://ollama.com
ollama pull deepseek-r1:7b
ollama pull qwen2.5:3b
ollama pull llama3.1:8b
ollama pull qwen3:8b
ollama pull qwen3.5:9b
ollama pull gemma3:4b
```

---

## Environment Variables

| Variable | Description |
|----------|-------------|
| `OPENAI_API_KEY` | OpenAI API key — used for online comp search and GPT models. Falls back to `openai.api_key` in deal config. |

---

## Output Files Reference

| File | Location | Description |
|------|----------|-------------|
| `Transaction_Comparables_<DealName>.xlsx` | `output/<deal>/` | Asset sales comps table |
| `Transaction_Comparables_<DealName>_records.json` | `output/<deal>/` | Raw parsed comp data; read by dashboard for editable preview |
| `Transaction_Comparables_<DealName>_geo.json` | `output/<deal>/` | Geocoded lon/lat per comp, hidden flags, map settings; source of truth for map |
| `Transaction_Comparables_<DealName>_map.png` | `output/<deal>/` | Mapbox location map |
| `Land_Sale_Comps_<DealName>.xlsx` | `output/<deal>/` | Land sales comps table |
| `Rent_Comps_<DealName>.xlsx` | `output/<deal>/` | Rent comps table |
| `Online_Comparables_<DealName>.xlsx` | `output/<deal>/` | AI-searched asset sales comps |
| `Online_Land_Comps_<DealName>.xlsx` | `output/<deal>/` | AI-searched land comps |
| `Online_Rent_Comps_<DealName>.xlsx` | `output/<deal>/` | AI-searched rent comps |
| `Investment_Rationale.md` | `output/<deal>/` | 3–5 section investment rationale (markdown) |
| `Investment_Rationale_meta.json` | `output/<deal>/` | Generation timestamp and elapsed seconds |
| `Investment_Rationale.docx` | Download only | Word version of the rationale |
| `Source_Audit.xlsx` | `output/<deal>/` | Citation audit — every number and named claim with source PDF, page reference, backend cross-check status, and human validation column |
