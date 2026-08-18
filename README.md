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
   - [1.1 The app interface — sidebar pages](#11-the-app-interface--sidebar-pages)
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
18. [Extraction accuracy — gold-file evaluation](#18-extraction-accuracy--gold-file-evaluation)
19. [Known limits and review notes](#19-known-limits-and-review-notes)
20. [Technical limitations — where we need help](#20-technical-limitations--where-we-need-help)

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

### 1.1 The app interface — sidebar pages

The rest of this document describes the pipeline internals for reviewers; this
subsection is the map of the Streamlit UI itself. The sidebar has one standalone entry
plus two collapsible groups (`frontend/app.py`, sidebar block):

- **📖 How to Use** — a static onboarding page (`render_user_guide()`). Explains the
  two deliverables and the intended workflow order below. No deal/config dependency,
  so it renders even before any deal exists — the default landing page for a new user.

- **📁 Deal List**
  - **🏗️ New Deal** — 2-step wizard (`render_new_deal_form()`) that creates a deal
    config: fill the essentials by hand, or upload a deal brief (PDF/Excel/text) and
    let `extract_from_document()` pre-fill fields via LLM.
  - **📁 Existing Deals** — select a previously created deal to open its Analysis
    Output pages.

- **📊 Analysis Output** — one deal-centric workspace (`render_analysis_output()`)
  with a single deal selector at the top and four views:
  - **📊 Overview** — read-only aggregation of everything generated for the deal
    (all comp tables, the map, the rationale); re-reads `output/<Deal>/` on every
    rerun, so edits made in the other views show up here automatically. One-click
    combined Word export (§14).
  - **📋 Comparable Analysis** — the editable extraction workspace per comp type
    (Asset Sales / Rent / Land): upload broker files, or use the **✏️ Enter or paste
    records manually** table for a blank/hand-typed comp list (tick **"Property name
    only"** to skip price/area requirements — useful for a curated name list to
    geocode/map with no file at all). **▶ Run Analysis** invokes the matching
    `scan_input_*.py` backend script as a subprocess (§5).
    - **🔄 Refine This Output** — a collapsed panel under the generated table: type a
      free-text instruction ("remove comps sold before 2022", "only keep freehold
      and 999-year leasehold"), pick an **Engine**, and re-run. This re-invokes the
      *same* `scan_input_*.py` script with `--from-records <existing records.json>`
      (skips re-extraction entirely, so it's fast, ~1–2 min) plus
      `--refinement-file <instructions>`; the script reads the instructions and
      calls `apply_refinement()` (`tools/llm_client.py:1241`), a router over three
      distinct mechanisms depending on provider/engine:
      - **"Standard"** (default, GPT) → `run_agent_loop_gpt()` — a ReAct-style
        tool-calling loop (not one of §4.2's three named agents, but the same
        reason→act→observe shape): the model calls *query* tools
        (`get_field_values`, `compute_stats`) to inspect the data,
        then *action* tools (`filter_range`, `keep_top_n`, `select_by_name`, …)
        that Python actually executes — no index-guessing or model-side numeric
        comparison. Loops up to 6 turns, each tool result fed back before the next
        decision, until the model calls `no_change` or stops.
      - **"Advanced (sandbox)"** (GPT, `--refine-engine code_interpreter`) →
        `run_code_interpreter_refine()` — a different mechanism entirely: sends the
        record list (each row stamped with a stable `_rid`) to OpenAI's **hosted**
        Code Interpreter, letting the model write and run real Python for logic
        fixed tool primitives can't express (e.g. "within 20% of the subject's
        size"). Executes in OpenAI's own container, never on this host. Safety
        design: the model must emit a final `RESULT: [id list]` line — only that id
        list is trusted and applied locally; the model's own echo of the record
        data is never trusted. Any failure (timeout, bad format) is a safe
        no-op — all records kept unchanged, not dropped. Opt-in only (sends comp
        data to OpenAI) — avoid for confidential deals.
      - **Ollama** (non-OpenAI provider) → `run_agent_loop()` — the same
        query-tool/action-tool ReAct pattern as "Standard", against a local model
        instead (5-turn cap vs. GPT's 6).
      Refined records overwrite `_records.json` and the Excel is regenerated from
      them — same render path as a normal run, not a separate code path.
  - **✍️ Investment Rationale** — generate/refine the rationale memo (§12) and its
    RAG source audit (§13).
  - **💬 Ask this Deal** (`_render_deal_chat()`) — grounded Q&A restricted to *this
    deal's* own uploaded documents and generated comp tables (listed in a "N
    document(s) in scope" expander). **Mechanism: full-text context stuffing by
    default, retrieval only as a fallback.** `_deal_corpus_files()` collects every
    in-scope file (raw uploads + shared market reports + this deal's generated comp
    Excels); `_extract_text_for_chat()` extracts plain text per file (PyMuPDF for
    PDF, capped 60 pages; CSV dump per Excel sheet), capped at 24k chars/doc and
    cached in `st.session_state` keyed by each file's (path, mtime, size) — so
    re-extraction only happens when a file actually changes, not on every message.
    - **Below `_CHAT_STUFF_BUDGET` (120k chars total, the common case):**
      `_build_chat_system_prompt()` concatenates ALL cached document text into one
      system prompt whole — no embedding, no chunking, no similarity search, unlike
      §13's page-level retrieval. Chosen deliberately: for a deal-sized document set
      this avoids RAG's main failure mode (a relevant chunk existing but never being
      retrieved), at the cost of not scaling past the budget.
    - **At/above the budget:** rather than silently flagging excess files "omitted"
      (as it did before), it switches to the same retrieval mechanism as §13,
      reused directly (`generate_investment_rationale._embed_batch` /
      `_cosine_sim`) — requires an OpenAI key. `_chunk_text_for_rag()` splits every
      document into overlapping ~2k-char windows (format-agnostic fixed-size
      chunking, since this has to handle both PDF prose and CSV-dumped Excel
      uniformly, unlike §13's page-aware PDF-only chunker); `_get_chat_rag_index()`
      embeds and caches all chunks (same file-signature cache pattern); each
      question is embedded and ranked against every chunk in the corpus by cosine
      similarity (`_retrieve_relevant_chunks()`), taking the best matches until the
      budget is filled; `_build_chat_prompt_from_chunks()` builds the prompt from
      only those excerpts. Without an OpenAI key, falls back to the old
      stuff-and-omit behavior instead (embeddings need a key; nothing else does).
    - A v1 limitation of the retrieval path: only the current question is embedded,
      not prior chat turns — a bare follow-up ("what about Q3?") may retrieve worse
      than a self-contained question.
    - Either way, the built prompt + chat history go to whichever model is selected
      in ⚙️ Shared Settings (GPT or a local Ollama model) in one chat completion call.

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

**PDF inputs** route through `pdf_extractor.py`: keyword page discovery →
`camelot`/`pdfplumber` table detection → column mapping → row assembly + dedup. This
one pipeline runs for both Ollama and GPT — they differ only in which model backs the
LLM-dependent steps (column mapping's LLM tier, the text-fallback prose parser), not in
a separate "vision path" that reads whole pages up front. (An earlier version of this
doc described a standalone whole-PDF vision extraction mode; that was never actually
implemented as a distinct path — what exists instead are the three narrower,
targeted vision checks below, all GPT-only.)

Three separate, narrow places where a rendered page image gets involved — not one
big "vision path":

- **Column-mapping vision cross-check** (§7.6) — fires when a column mapping is
  already flagged low-confidence; renders that page and asks the model to confirm
  which column is which. Corrects a *mapping*, never invents a row.
- **Stage 3+4 vision retry** (below) — fires when a page Stage 1 was confident about
  (title matched) ended up with zero usable rows after column mapping + assembly;
  renders that page and asks the model to locate and extract the missing table
  directly, bypassing table-grid detection. Evidence-bound against the page's own
  text layer — see below.
- **Stage 6 name cross-check** (§7.10, off by default) — renders a page to catch a
  font-encoding glyph substitution no text-based stage can see. Only ever corrects a
  name's characters, never adds a row or a value.

#### Stage 3+4 vision retry (new 2026-07)

Not a new numbered stage — a retry loop inside record assembly (Stage 3+4/§7.6-7.7),
not a standalone step, because its trigger condition (did this page's assembly
actually produce a record with the field we need) can only be evaluated *after*
column mapping has run — Stage 2's table-grid detection has no way to know yet
whether a table it found is even the right type.

Table-grid detection (camelot) sometimes merges the target table with a neighbouring
one — e.g. a report page with a 2-row "Key Sales Transactions" table sitting flush
against a much larger market-statistics table; camelot grabs both as one region, and
after that region gets auto-split it can lose its header row and end up unmapped
(`property_name`/`price_sgd_m` both read "not found") even though the 2 real rows are
still sitting there in the page's own text, just outside any table camelot could
isolate cleanly. Before this fix such a page silently produced 0 records — Stage E's
fallback gate (§7.5) never re-tried it, because camelot *did* find something on that
page (the stats table, or a wrong-type table like a lease table), and the gate only
tracks "did this page yield any table," not "did it yield the table we actually need."

Fires only when BOTH hold: (a) Stage 1's page discovery matched this page by its own
title/heading keyword (independent evidence the target table should be here), and
(b) after full table detection + column mapping, that page contributed zero records
with a real value for the type's key field(s) (`price_sgd_m`/`price_psf_gfa` for sales,
`asking_rent`/`eff_rent`/`nla_sf` for rent, `price_sgd_m`/`price_psf_ppr` for land —
passed in as `value_fields`). A page whose only "records" are name-only rows from the
WRONG table (e.g. a lease table's rows, which have a name and an SF value but no price)
still counts as zero and triggers the rescue — see `_is_useful` in `pdf_extractor.py`.

When it fires: renders the page, sends it to GPT-4o with the target field
definitions, and asks it to locate the one matching table on the page and copy its
rows — explicitly told to preserve numbers exactly as printed (no reformatting) and to
ignore any other table on the page (statistics, or the wrong transaction type).
**Evidence-bound like every other stage**: every returned field value must appear
verbatim in that page's own `pdfplumber` text layer or it is dropped from the row (not
the whole row); vision can only relocate/re-associate a value the text layer already
contains, never introduce a number the page's own text doesn't have. GPT-only; on by
default (only fires on an already-failing page, so it adds no cost to a page that
already works); `PDF_VISION_RESCUE=0` disables it.

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
`shared_settings.geocoding_provider`, **falling back to Google** on any failure —
not Mapbox, which is only ever used to render the static map image (§10.2); the two
are fully independent and share no credential.

For file/upload comps (sales, rent, land — `scan_input_*.py`), every record is
geocoded using one shared address-vs-name policy across all three comp types
(`resolve_geocode_queries()`, detailed in §10.1): a genuine street address is geocoded
alone; a building name with no real address falls back to NAME plus a submarket/
district hint, since most market-report tables give only a name and no address. A comp
that resolves to the country centroid is flagged `ON COUNTRY CENTROID` (`_geo_suspect`)
for the analyst rather than silently plotted — see §10.1 for exactly what that flag
does and does not catch. Records from different uploaded reports describing the same
comp are then merged (if they agree) or flagged for review (if they don't) — §10.3.

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

**Verify & reflect** (inside the same call): a deterministic check flags a mapping
whose column *values* don't fit the field (a "money" field that's mostly text, a
"label" field that's pure numbers, …) or whose embedding score is below `0.45`. An LLM
then re-maps **only the flagged fields** — the confident ones stay locked, so a good
mapping can never be reshuffled by a bad one. When the flagged table came from a PDF
(`pdf_path`/`page_num` passed through from Stage 7.7's caller) and the provider is GPT,
this re-map is **vision-augmented**: the page is rendered and sent alongside the header
text, so a column whose true identity survives only in the page's visual layout (e.g. a
table title fused onto its header by extraction) can still be resolved correctly, not
just one whose header text still contains a recognisable trace. Falls back to the
text-only call on any render/API failure — never a regression versus not having it.
Excel/CSV input scanning has no page to render and always uses the text-only call.

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

### 7.9 Stage I — LLM verification (`_llm_verify_records`)

Runs once per document, after every record is assembled — the only stage that checks a
**whole record** against the source rather than one column or one field. Given the raw
table text (headers + rows, no narrative prose — `_page_tables_context(tables_only=True)`,
so the model can't "correct" a table figure toward a rounded number from surrounding
commentary) and the assembled records, it checks two things per record:

1. **Cell correctness** — does the value match the actual source-table cell for that
   transaction's row (catches wrapped text fused from two different rows)?
2. **Column mapping accuracy** — does the value belong under the field it's currently
   filed in, per the table's own column header?

Corrections must be evidence-bound: any change to a non-null value requires a verbatim
quote from the source (checked in code, not taken on the model's word); a value it
can't verify is blanked, never guessed. A run of deterministic pre-checks
(`_deterministic_flags`) flags obvious problems first (a value outside a plausible
range, a size field holding prose, a price ÷ area disagreement) so the LLM gets
targeted questions instead of blind-auditing every cell.

**Guardrails on what the model is allowed to touch** — each added after a specific
observed failure, not speculatively:

| Guardrail | Blocks | Added after |
|---|---|---|
| Name lock | Any rewrite of `property_name` | Every observed name edit lost fidelity (a stake `(50.0%)` dropped) and never improved one |
| Numeric substitution | One non-empty number replacing a *different* non-empty number | Prose-rounding: a table's `1,133.00` overwritten by a paragraph's rounded `1,100` |
| Text blanking | Blanking a short text/category value still verbatim in source | `land_zoning='Commercial'` blanked despite the row being right there in SOURCE |
| Qualified-number blanking | Blanking a number carrying a scale word (`million`, `per key`, `psf`, …) still verbatim in source | A hotel's "0.78 mil per key" price blanked outright on two separate reports |
| **Bare-number blanking** | Blanking a plain, unqualified number (e.g. `price_sgd_m`) still verbatim in source | Savills Q3 2025 land: `price_sgd_m` nulled on 2 of 3 identical reruns — the one value *shape* the two guards above didn't cover (not text, no scale word) |
| Evidence gate | Any non-null correction lacking a verbatim source quote | Generalizes all of the above — stops outright fabrication |

The bare-number guardrail (last row) closed the only remaining gap: **before it, this
stage was the one place in the whole pipeline that could silently destroy an already-
correct value with zero protection**, and it's what made Savills Q3 2025 land's land
table non-deterministic (5/5 correct on one run, 1/5 on another, same code, same input).

### 7.10 Stage J — vision name cross-check (`_vision_verify_names`, off by default)

Every stage above — including 7.9 — reasons over **text already extracted**, however
mangled. This is the one stage that looks at the **rendered page image** instead, and
it exists for exactly one failure mode nothing text-based can catch: a PDF's embedded
font maps the glyph that *displays* as one character to a *different* character's code
in the text layer (capital `I` shown on the page, `l` in the text). The extracted text
still reads as perfectly plausible English, so re-reading the same text (all Stage 7.9
ever does) can never recover the true character — only the pixels can.

Deliberately narrow: only `_NAME_KEYS` fields are checked, and a correction is applied
**only** when it differs from the original by known font-confusable character pairs at
matching positions (`I`/`l`/`1`/`\|`, `O`/`0`, `S`/`5`, `B`/`8`) — anything else is a
wholesale rewrite and blocked, same discipline as the name lock in 7.9. Off by default
(`vision_verify_names=True` or `PDF_VISION_VERIFY=1`); no-ops for any provider but GPT.

**Measured, not assumed, against the 52-file gold set (§18):** turning it on for a full
run — 249 names checked across 50/52 files — applied **zero** corrections (no genuine
glyph-confusion errors exist in this corpus) and **blocked 135** attempted rewrites the
guardrail correctly recognised as *not* character-confusion (dropping a stake `%`,
"improving" capitalisation that was actually correct as printed) — real evidence the
narrow-correction discipline holds under load, not just in theory. Added roughly 15%
wall-clock time per file, and one page failed outright on an OpenAI rate limit during a
back-to-back 52-file run — a real reliability cost with zero measured benefit on known
sources, which is why it stays off by default rather than being enabled broadly. The
code is not dead weight to remove, though: it costs nothing while idle, and it remains
the only way to diagnose a suspected glyph-confusion case on a single file by hand (see
the Collyer Quay case in §18, where it correctly confirmed the pipeline's reading was
right and the gold file's was wrong).

### 7.11 Debugging a table that will not extract

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

Deal configs carry **no** geocoding token; the key lives in Shared Settings only —
`shared_settings.json`'s `google_maps_key` locally, or the `GOOGLE_MAPS_KEY` secret on
Streamlit Cloud (bootstrapped into `shared_settings.json` at startup, §16). This is a
single **global** setting, not per-country — a Korea deal and a Singapore deal use
whichever provider is configured, both locally and on the cloud; there is no per-deal
or per-country override.

`geocode_any(queries, …)` tries each query in `queries` **in order** and returns the
first one any provider resolves — it does not compare candidates for precision, so a
broad-but-successful early match wins over a more precise query later in the list.

**Address-vs-name query policy — shared by sales, rent and land**
(`resolve_geocode_queries()` in `generate_comps_map_base.py`, used identically by all
three `scan_input_*_comps.py` scripts):

1. **`looks_like_real_address(addr, country_name)`** decides whether the Address field
   is worth geocoding on its own: it must be non-empty, not a bare asset-class/
   placeholder word (`_NON_ADDRESS_WORDS` — "Office", "N/A", "—", …), contain a digit,
   and — for Singapore — also contain a recognised street-type keyword (Road, Street,
   Ave, …); for any other country a digit alone is enough (a full local address is
   still worth geocoding even without an English street-type word, e.g. a Korean or
   Japanese address).
2. **If it looks like a real address → geocode by ADDRESS ONLY**, with the country name
   appended as a text suffix (`", Singapore"`, `", South Korea"`, …) — no name fallback
   is added to that query list. Falling back to the name would collapse distinct
   properties that share a brand ("Weave Place – Hoegi" and "Weave Place – Gangnam
   Station" both strip to "Weave Place" → the same pin).
3. **Otherwise → geocode by NAME**, with the submarket/district value (if the source
   table has one) passed as a locality hint, and the country name appended as a text
   suffix here too.

**The country-name text suffix matters even for Singapore, not just foreign
countries.** `country_code` is *also* applied separately as an API-level component
filter (`components=country:SG`), but that restricts *where* a match may land — it does
not help Google's free-text matcher recognise an ambiguous name. Measured case: Google
resolves `"Bank of Singapore Centre"` (name only, `components=country:SG`) to the
country centroid, but resolves `"Bank of Singapore Centre, Singapore"` (same
restriction, country name added to the query **text**) to the correct rooftop address.
Every comp type now includes this suffix unconditionally (a 2026-07 fix — sales
previously omitted it specifically for Singapore, rent/land omitted it for every
country; §18 has no bearing on this, it's a geocoding fix, not an extraction one).

**The `ON COUNTRY CENTROID` flag only catches one specific failure mode — a geocode
landing within 1.5 km of the country's centroid — and nothing else.** It does **not**
catch: a result that resolves to the correct country/city but the wrong specific
building (e.g. a same-named branch elsewhere, or a submarket/neighbourhood centroid
that is not the actual building — still off by up to ~1 km but far enough from the
country centroid to go unflagged); or a precisely-geocoded but wrong address (if the
extracted address itself was wrong, Google will happily geocode it precisely to the
wrong place, silently). A comp with no flag is not proof its pin is correct — only proof
it isn't the crudest kind of failure. There is currently no automated check beyond this;
verifying pin placement for any given comp is a manual, spot-check exercise.

`country_code` must be set explicitly in the deal config. There is deliberately **no
address-sniffing heuristic** — a wrong country guess silently geocodes a comp onto the
wrong continent.

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

### 10.3 Cross-report dedup and conflict-flagging (uploaded files)

**This is a different mechanism from §11.6** — that one dedups results from ONLINE
search; this one dedups records extracted from multiple UPLOADED reports (e.g. the
analyst drops in both a Colliers PDF and a Savills PDF, and one deal appears in both).
Both solve the same underlying problem (the same real-world thing reported by more than
one source) but use different code, different thresholds, and run at a different point
in the pipeline — they are not the same check twice.

The shared logic lives in `tools/calculations.py` (`dedup_cross_source()` and
`flag_cross_source_conflicts()`), called identically by all three
`scan_input_*_comps.py` scripts right after geocoding, before map markers are numbered.
As of 2026-07, **all three comp types get this** — previously only sales did; rent and
land are now on the same code path.

Two records are only ever compared when they came from *different* uploaded files
(`_source` differs) — same-file duplicates are a different problem, handled earlier by
the extraction pipeline's own name-based dedup.

**Merge when they agree** (`dedup_cross_source`) — geocoded within **50 m** of each
other, AND (for sales/land only) the property names token-overlap ≥ **90%**, AND the
first mutually-present value field agrees within a per-type tolerance:

| Comp type | Also requires name overlap? | Value field(s) checked, in order | Tolerance |
|---|---|---|---|
| Sales | Yes | `price_sgd_m` | 5% |
| Land | Yes | `price_sgd_m` | 5% |
| Rent | **No** | `eff_rent`, then `asking_rent` | 5% |

Rent is deliberately the odd one out: the SAME building legitimately produces many
different leases at many different rents (different tenants, different units,
different dates). Its name always "matches" trivially against itself, so requiring
name overlap on top of that would add nothing — the rent figure agreeing is what
actually establishes "this is the same lease reported twice," not the name. All three
comp types now use the same 5% value tolerance — even a lump-sum sale/land price is
not expected to legitimately vary more than that across two sources describing the
same deal.

Sales/land get the name-overlap check because the reverse risk dominates there: a
building is normally sold/tendered once in a reporting window, so two DIFFERENT nearby
buildings that happen to land within 50 m of each other and coincidentally report a
similar price are a real (if rarer) risk without it.

**Exact-coordinate override (sales/land only, `exact_km=0.02` — 20 m).** If two
records (the name check above still applies) geocode within 20 m of each other —
near-identical, not just "same building" — that alone is treated as sufficient
evidence to merge, even if the value fields disagree by more than 5%: two sources
rounding or transcribing one real price differently is judged more likely than two
unrelated deals landing on the exact same rooftop pin. **This never applies to
rent, deliberately** — it would merge away a genuine price *disagreement* between two
different leases at the same coordinates, which for rent is the entire point (many
legitimately different leases, same building, same pin); silently discarding one
there instead of flagging it would be a real data-loss bug, not a convenience. For
sales/land, that same risk doesn't arise the same way, since a building is normally
sold once.

The most complete record (most non-blank fields) is kept as the base; any field still
blank on it is filled from the other record, so complementary data from both sources
survives (e.g. one source has psf, the other has the cap rate). `_source` becomes the
combined list (`"colliers+savills"`).

**Flag, don't merge, when they disagree** (`flag_cross_source_conflicts`) — geocoded
within **200 m** (looser than the merge check above — flagging is low-risk, it never
deletes a row, so it can afford to be more inclusive), plus a property-name token-
overlap check (≥ 90%, all three comp types — the flag check, unlike the merge check,
has no reason to skip this for rent: flagging two same-named, nearby, disagreeing rows
for review is harmless even there) so two genuinely different buildings that merely
sit close together in a dense CBD are never flagged, but the value field(s) disagree
by more than 5%:

| Comp type | Value fields checked (both, independently) |
|---|---|
| Sales | `price_sgd_m`, `price_psf_gfa` |
| Rent | `eff_rent`, `asking_rent` |
| Land | `price_sgd_m`, `price_psf_ppr` |

Both rows are kept, both get a `_review_flag` note naming the disagreeing sources and
values, and the run log prints a `[cross-source CONFLICT]` line — silently picking one
source or silently keeping both without comment would hide a real data question, so
neither happens.

Threshold note: the merge check's 50 m now matches §11.6's online-search dedup radius
(both were tightened for the same reason — two distinct nearby buildings should not be
silently fused into one row); the conflict-flag check stays at a looser 200 m since it
only raises a flag, never deletes data.

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

**For ONLINE search results only** — the different mechanism that dedups records across
multiple UPLOADED files is §10.3, not this section. The two are separate code paths with
different thresholds; §10.3 explains why.

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
known-good run. For the standing accuracy record (which reports have been checked
against a hand-built gold file, and what's still wrong), and for the methodology behind
it, see §18.

### 17.1 Detection without an LLM (free, deterministic)

`extract_page_tables` makes no model call, so table *detection* can be checked
instantly and repeatably (see §7.11):

```python
from scan_input_sales_comps import _PDF_SECTION_KEYWORDS
from pdf_extractor import find_relevant_pages, extract_page_tables
pages = find_relevant_pages("report.pdf", _PDF_SECTION_KEYWORDS)
for t in extract_page_tables("report.pdf", pages):
    print(t["page_num"], t["headers"][:4], len(t["rows"]))
```

This is the fastest way to tell a **detection** bug from a **model** bug.

### 17.2 The baseline diff — the regression check that matters

Before changing detection, snapshot every PDF;
afterwards, re-run and diff. **Any file you did not intend to change must be
byte-identical.**

```python
# capture: headers + row counts per table, per PDF, for every sales-keyword PDF
# (no LLM calls — extract_page_tables only), then re-run after the change and diff.
```

A detection change should alter only the files it targets. One that also moves
unrelated files is matching too broadly and needs a tighter condition.

### 17.3 House rules a change must not break

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

## 18. Extraction accuracy — gold-file evaluation

This is a standing audit of the extraction pipeline's accuracy, run against hand-built
ground truth — separate from §17's "verify a code change didn't regress anything."
For each report below, someone transcribed the source table by hand into a gold file
(`eval/gold/`), ran that report through the actual extraction pipeline
(`eval/run_extract.py`), and compared the two row by row. A report only appears here
once that comparison has actually been done — nothing below is inferred from the file
existing or from a table merely being found.

**Methodology: this project follows eval-driven development** — the standard approach
for a pipeline with a non-deterministic (LLM) component, where a traditional unit test
can't assert one exact output:

- **Golden/gold-labeled test set** (`eval/gold/`) — hand-transcribed ground truth,
  audited and corrected when it's found wrong, not just assumed correct forever (the
  Woods Square fix below is a gold-file fix made this way, not a pipeline fix).
- **Full regression run on every change** — all 52 gold files re-run and diffed after
  any change to `pdf_extractor.py`, the scan modules, or the column mapper — not just
  the one file that motivated the change (§17.2 has the equivalent check for a
  non-extraction-accuracy code change).
- **Guardrails around LLM output, not trust in it** — every LLM correction must be
  evidence-bound (a verbatim quote checked against source in code) or it's rejected,
  never accepted on the model's word alone (§7.9).
- **A real regression vs. a pre-existing gap merely newly surfaced are not the same
  thing** — compare against the last known-good baseline before concluding a change
  broke something.
- **Repeat runs to catch sampling variance** — an LLM-backed stage passing once is not
  proof it's stable. The Savills Q3 2025 land bug (§7.9) below was only found by
  rerunning the identical input several times and diffing the outputs against each
  other.

Two ways this is lighter-weight than a mature production eval setup, worth being honest
about: the gold set (52 files, 223 rows) is smoke-test-sized, not statistically large,
and there's no held-out set kept apart from tuning — so a good score below is evidence
the pipeline handles *these* sources well, not proof it generalizes to a source it
hasn't seen.

**Every result below was produced with GPT (`gpt-4o-mini`), not Ollama** —
`eval/run_extract.py` defaults to it and none of these runs changed that. If a deal in
the app is configured for Ollama, or for `gpt-4o` instead of `gpt-4o-mini`, its output
is not guaranteed to match what's recorded here.

**Summary — by report, not by row, counting only bugs that reach the actual
deliverable** (the Transaction Comparables Excel table an analyst opens —
`generate_sales_comps_table.py`'s 14-column `OUTPUT_SCHEMA` has no Buyer or Sale Type
column, so a wrong value in either field is a real extraction error but never a
rendered one, and doesn't count against a report below):

| Source | Quarters covered | Passing |
| --- | --- | --- |
| Colliers | Q3 2023 – Q1 2026 (sales) | 10/11 |
| CMMB Capital Markets | Q4 2023 – Q1 2026 (sales) | 9/9 |
| CMMB Office MarketBeat | Q1 2025 – Q1 2026 (sales + rent) | 10/10 |
| Savills | Q3 2023 – Q1 2026 (sales + land) | 21/22 |
| **Total** | | **50/52** |

Each unit counted above is one report-and-type combination (e.g. "Savills Q1 2025
sales" is one unit), not one row. 52 is the count of unique gold files; one duplicate
upload (`singapore-capital-markets-mb-4q2023__sales.gold (1).json`, byte-identical to
the non-duplicate copy) is excluded.

**How to read the matrix below:** ✅ means every row matched gold exactly, **or** every
difference from gold was confirmed to be gold/source fidelity rather than a pipeline
capability gap (see the counting rule just below) — a fraction (`4/5`) means at least
one row still has a real, uncounted-as-fine difference, tagged ❌ or a still-counted ⚠️:

- **Counts as passing (not a capability gap)** — the pipeline's output is *provably*
  the correct one: either the source PDF's own text layer literally contains what the
  pipeline extracted (confirmed via ≥2 independent extractors and/or a vision
  cross-check, so no extraction approach could ever produce gold's answer instead —
  e.g. Collyer Quay's `l22-24FL`), or the difference is a lossless, purely cosmetic
  character-encoding variant with zero effect on meaning or readability (e.g. a curly
  vs. straight apostrophe). Getting *this* right isn't a matter of the pipeline trying
  harder — there is no alternative correct output to reach.
- **Still counts as a real gap (⚠️, not ❌ — confirmed non-destructive, but still
  open)** — the difference is something a better extraction approach *could* in
  principle fix: a name genuinely truncated by the text-fallback path (info is lost,
  not just re-encoded), or a word visibly split by a ligature-rendering artifact (the
  output reads wrong to a person, even though the root cause is the PDF's font, not a
  wrong-field mapping). Two such gaps remain — see footnotes 1 and 4.
- **❌ = a real extraction error** — a value landed in the wrong field, was fabricated,
  or was lost. None of the differences below are this category — the last one (B6, a
  unit price leaking into the GFA field) was fixed 2026-07-20.

`—` means that source/type wasn't published or checked for that quarter. Buyer and
Sale Type are excluded from this check since neither is a column in the Excel output.

**The 2023–2024-quarter sources (Colliers, CMMB Capital Markets, and Savills) also
show a `land_zoning`/`sale_type` value swap on nearly every row — not counted as a
bug, already excluded from the matrix below.** The source tables have a "TYPE"/"SECTOR"
column (values like "Retail", "Office", "Hospitality") which the pipeline correctly
maps to `land_zoning`, matching the pipeline's own established sector-word rules
(confirmed directly against the source table headers). The gold files for these
reports instead recorded that same value under `sale_type` — a gold-file convention
gap, not a pipeline defect.

**All 52 report/type combinations, one source+type per column:**

| Quarter | Colliers (sales) | CMMB CapMkt (sales) | CMMB Office (rent) | CMMB Office (sales) | Savills (land) | Savills (sales) |
| --- | --- | --- | --- | --- | --- | --- |
| Q3 2023 | ✅ | — | — | — | ✅ | ✅ |
| Q4 2023 | 2/3⚠️¹ | ✅ | — | — | ✅ | ✅ |
| Q1 2024 | ✅ | ✅ | — | — | ✅ | ✅ |
| Q2 2024 | ✅ | ✅ | — | — | ✅ | ✅ |
| Q3 2024 | ✅ | — | — | — | ✅ | ✅ |
| Q4 2024 | ✅ | ✅ | — | — | ✅ | ✅ |
| Q1 2025 | ✅² | ✅ | ✅ | ✅³ | ✅ | 4/5⚠️⁴ |
| Q2 2025 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Q3 2025 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Q4 2025 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Q1 2026 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

1. Property name truncated by the text-extraction fallback (this report has no gridded
   table on that page) — "Visioncrest Commercial" → "Visioncrest". A real accuracy gap
   in that fallback path, but not a wrong-field defect, so ⚠️ not ❌ — **still counts
   against the 50/52 total**, unlike footnotes 2-3 below.

2. "Northpoint city South Wing" — the raw table cell in the source PDF itself already
   has the lowercase `c`; the pipeline reproduces it faithfully. Counted as ✅: there is
   no correct output other than what the pipeline already produces.

3. The gold file and the PDF's own text layer disagree on one character — see below.
   Counted as ✅ for the same reason as footnote 2 — confirmed via two independent text
   extractors *and* a vision cross-check that the source itself reads this way.

4. A ligature-splitting cosmetic near-miss ("floors" → "fl oors") — the pipeline's own
   text-layer extraction, not a gold-file error, but the output does visibly read wrong
   to a person — **still counts against the 50/52 total**.

**CMMB Office Q1 2025 sales' Collyer Quay row isn't a pipeline bug — the gold file and
the PDF's text layer disagree on one character.** The PDF's own text layer reads "20
Collyer Quay (l22-24FL)" — confirmed directly: two independent text extractors
(pdfplumber, PyMuPDF) both read `l` there, and a vision-LLM cross-check against the
rendered image (`_vision_verify_names` in `pdf_extractor.py`, an optional Stage 6 built
for exactly this — off by default, `PDF_VISION_VERIFY=1`) also read it back as `l`. The
gold file transcribes this as "20 Collyer Quay (22-24FL)" with no letter at all — a
human reading the rendered page apparently judged the stroke as decorative, which the
text layer disagrees with. Not fixable by any extraction approach either way.

**Thirteen fixes went into this session (2026-07-20)**, the last two just now: the
name-length filter (`len(name) > 80`) was a redundant, harmful leftover from before
`_is_sentence_fragment()` existed — the latter already distinguishes a genuinely
mangled prose fragment from a real (if verbose) property name structurally, so the
blunt length cutoff was deleted, restoring the long portfolio-style names Savills
Q3/Q4 2024 and Q1 2025 previously lost. And a scaled/qualified unit price (a hotel's
"0.78 mil per key") is no longer blanked outright — it's preserved as reported text in
a new `price_psf_gfa_display` field (mirroring the existing `price_sgd_m_display`
pattern for price ranges), so the analyst still sees the figure even though no psf
number is calculated from it. A full regression sweep across all 52 gold files
confirms nothing broke. Every difference remaining above has been individually
checked and is ⚠️, not ❌ — this table has no open real bugs.

**Three more fixes went in 2026-07-21**, the first two described in full in §7.6 and §7.9:
column-mapping's verify-reflect step (§7.6) can now cross-check a flagged field
against the source PDF's rendered page image, not just header text, when the provider
is GPT — a full regression sweep confirmed zero change to any of the 52 results (32
vision cross-checks fired, all on tables that were already noise-filtered before
reaching the deliverable, so nothing here moved). And Stage 7.9's evidence-bound
guardrails against destructive corrections gained the one case they didn't already
cover — blanking a plain, unqualified number still verbatim in source — closing the
gap that made **Savills Q3 2025 land** non-deterministic (previously 5/5 on some runs,
1/5 on others, same code and input; confirmed fixed by 4 consecutive reruns all
returning 5/5 after the guardrail was added). It reads as ✅ in the matrix above
because a lucky run happened to be the one recorded — it is now ✅ *reliably*, not by
chance. Third, the Woods Square row of **Savills Q1 2025 sales** (footnote 4) was a
genuine gold-file error, not a pipeline one: gold's own transcription was shorter than
what the source table's PROPERTY cell actually says, confirmed against the pipeline's
`_prov`-backed extraction (exact price/date/buyer match). The gold file has been
corrected to the full text, so this row now passes — only the unrelated ligature
near-miss remains in that combination.

**No-LLM baseline — same 52 combinations, every LLM stage disabled
(`llm_cfg = {"provider": "rules"}`, no Ollama running, no OpenAI calls of any kind):**

| Source | Quarters covered | Passing |
| --- | --- | --- |
| Colliers | Q3 2023 – Q1 2026 (sales) | 9/11 |
| CMMB Capital Markets | Q4 2023 – Q1 2026 (sales) | 9/9 |
| CMMB Office MarketBeat | Q1 2025 – Q1 2026 (sales + rent) | 10/10 |
| Savills | Q3 2023 – Q1 2026 (sales + land) | 21/22 |
| **Total** | | **49/52** |

| Quarter | Colliers (sales) | CMMB CapMkt (sales) | CMMB Office (rent) | CMMB Office (sales) | Savills (land) | Savills (sales) |
| --- | --- | --- | --- | --- | --- | --- |
| Q3 2023 | ✅ | — | — | — | ✅ | ✅ |
| Q4 2023 | 0/3❌⁵ | ✅ | — | — | ✅ | ✅ |
| Q1 2024 | ✅ | ✅ | — | — | ✅ | ✅ |
| Q2 2024 | ✅ | ✅ | — | — | ✅ | ✅ |
| Q3 2024 | ✅ | — | — | — | ✅ | ✅ |
| Q4 2024 | 0/3❌⁶ | ✅ | — | — | ✅ | ✅ |
| Q1 2025 | ✅² | ✅ | ✅ | ✅³ | ✅ | 4/5⚠️⁴ |
| Q2 2025 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Q3 2025 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Q4 2025 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Q1 2026 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

⁵ genuinely LLM-required, not a code gap: this report's transaction table sits in a
page with no gridded structure camelot/pdfplumber can recover at all — the pipeline
falls through to its text-only prose-extraction path, which is itself an LLM call
(read a paragraph, pull out name/price/psf as JSON). With no LLM reachable that path
fails outright (`Connection refused`) and the page contributes zero records. With GPT
this same path correctly parses it back to the ✅² result in the table above — no
deterministic rule-based approach can substitute here; parsing structured facts out
of free-form prose is exactly the kind of task only a language model does
⁶ different failure mode from Q4 2023 — a real gridded table IS found here, but its
headers only resolve enough columns to pass the deterministic tiers (exact/price-rule/
embedding); without the LLM column-mapping fallback the remaining fields stay
unmapped, the assembled rows come out with almost no fields filled, and the
pre-Stage-5 garbage filter (correctly) discards them as noise before they'd ever reach
a human. With GPT the LLM tier resolves the missing columns and the report passes.

The other three combinations (Colliers Q1 2025, CMMB Office Q1 2025 sales, Savills
Q1 2025 sales) show **identical results with or without any LLM** — proof those
particular differences are text-layer/gold-transcription artifacts, not something an
LLM stage is fixing or could fix (two of the three, footnotes 2-3, are ✅ under the
counting rule above regardless; the third, Savills' "fl oors" ligature, is a real gap
neither the LLM nor its absence changes). Net effect of every LLM stage in this
pipeline, measured across all 52 files: it rescues exactly 2 files (6 rows) that are
otherwise completely unextractable by deterministic code alone; it changes nothing else.

**Row-level accuracy — same 52 files, 223 rows total, derived from the matrices and
footnotes above (not a separately re-run metric):**

| Baseline | Rows correct | Wrong rows |
| --- | --- | --- |
| GPT | **221 / 223** (99.1%) | 2 |
| No-LLM | **216 / 223** (96.9%) | 7 |

The wrong rows, by combination:

| Combination | Baseline | Wrong rows | Which rows, and why |
| --- | --- | --- | --- |
| Colliers Q4 2023 sales | GPT | 1 of 3 | Footnote 1 — property name truncated by the text-fallback path ("Visioncrest Commercial" → "Visioncrest") |
| Colliers Q4 2023 sales | No-LLM | 3 of 3 (all) | Footnote 5 — the text-fallback path is itself an LLM call; with no LLM reachable it fails outright and the page contributes zero rows |
| Colliers Q4 2024 sales | No-LLM | 3 of 3 (all) | Footnote 6 — without the LLM column-mapping fallback, headers stay unresolved, rows come out nearly empty, and the pre-Stage-5 garbage filter correctly discards them |
| Savills Q1 2025 sales | GPT **and** No-LLM | 1 of 5 (same row, both baselines) | Footnote 4 — a font-ligature artifact ("floors" → "fl oors"), unrelated to the LLM either way |

Reading the table: **Colliers Q4 2023's severity is worse under No-LLM (3 wrong) than
GPT (1 wrong), but the report-level matrix above counts it as "not clean" in both** —
it was already 2/3⚠️ under GPT, so going to 0/3❌ under No-LLM doesn't move the
50/52-vs-49/52 report-count. Only **Colliers Q4 2024** flips from a clean ✅ (GPT) to
a full miss (No-LLM), which is the one report that does move the report-level count.
This is why the row-level gap (221 vs 216 = 5 rows) is wider than the report-level
gap (50 vs 49 = 1 report) — the two counts are answering different questions.

**2026-07-26: passing rate refined from 48/52 to 50/52 — a counting-convention fix, not
a code change.** Re-running the full 52-file suite with a stricter field-by-field
diff (not just name+price) turned up two more confirmed-cosmetic variants (a curly vs.
straight apostrophe in "Moody's" and in "MCL Land's assets" — already ✅ cells, unaffected)
and, in reviewing all of it together, established the counting rule in the legend above:
a difference only counts against the total when the pipeline *could* in principle have
produced a different, more correct output. Applying that consistently moved Colliers
Q1 2025 sales and CMMB Office Q1 2025 sales from ⚠️-fractions to ✅ (their only
differences — the "Northpoint city" lowercase `c` and the Collyer Quay `l` — are both
confirmed source-fidelity, not pipeline misses), while Colliers Q4 2023 sales
(Visioncrest truncation) and Savills Q1 2025 sales (the "fl oors" ligature) correctly
stay counted, since both are the kind of gap a better extraction approach genuinely
could fix. **48/52 and 50/52 describe the exact same underlying results** — nothing in
the pipeline changed between them, only which differences are judged to count.

**Same date: a new vision retry was added for table detection itself** (§7, "Stage 3+4
vision retry") — the extraction stages already documented here (§7.6's column-mapping
cross-check, §7.9's evidence-bound guardrails) assume the right table was already
located; this new retry fires when it wasn't. A page Stage 1 confidently matched by its
own title/heading that nonetheless assembled zero records with a real value for the
type's key field(s) gets re-read via vision, evidence-bound against that page's own
text layer exactly like every other LLM-touching stage in this pipeline. A full
52-file regression (this stricter field-by-field diff) confirms **zero change** — the
retry never fires on any of the 52, because none of them have a page Stage 1 matched
that still comes up empty after assembly. It was written for, and fixes, a report
**outside this 52-file set**: a Cushman & Wakefield Singapore Office Q2 2026 report
where camelot merged a genuine 2-row "Key Sales Transactions" table with a much larger
neighbouring statistics table, losing both real rows entirely (0 records) even though
Stage 1 had correctly matched the page by title. The retry recovers both rows correctly
(confirmed against the PDF's own text). A second, related addition: if Stage 1 finds
*no* candidate page at all for a comp type deliberately run on this file (routed there
by the classifier or by analyst override), a blind vision search of the first few pages
now runs before giving up outright, using the same evidence-binding.

---

## 19. Known limits and review notes

Ranked by what a reviewer should look at first.

1. **Credentials live only on disk in `configs/` — never in git.** Deal configs carry no
   map credential (the Mapbox token moved to Shared Settings when geocoding split to
   Google). The real secrets are `shared_settings.json` (all API keys) and ~810
   `tmp*.json` left by a failed cleanup path (each still holds the old Mapbox token).
   Neither is tracked by git, and — verified — no real token exists in the local git
   history or the cloud repo, so **git/GitHub is clean**. Caveat: there is no `.gitignore`
   fencing these off (the local checkout is rooted at the home directory, not the project
   folder), so they are merely *untracked* — a stray `git add`, or a Finder "Compress" of
   the project folder, WOULD include them. The only exposure path is copying/zipping
   `configs/` directly. **Share via the cloud repo, or a zip that excludes
   `configs/shared_settings.json` and `configs/tmp*.json`.** The temp files are safe to
   delete; the generating path should clean up after itself. Rotate the Mapbox token only
   if the raw folder was ever shared outside git.
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

## 20. Technical limitations

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
