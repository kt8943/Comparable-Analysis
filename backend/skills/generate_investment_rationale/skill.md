---
name: generate_investment_rationale
description: Generate a Markdown investment rationale from market report PDFs using a two-stage Ollama pipeline
type: pipeline
requires:
  config_keys:
    - output_file
    - market_report_files      # list of PDF paths
    - llm.ollama.base_url
    - llm.ollama.model
    - subject_property.property_name
    - subject_property.asset_class
  skills: []
allowed_tools:
  - tools.llm_client.ollama_post
---

## When to use
The analyst has uploaded one or more market report PDFs and wants a structured investment rationale document generated from their content.

## Instructions
1. Load deal config from `--config` path
2. **Stage 1 — Extract** (cached per PDF):
   - Extract raw text from each PDF using `pypdf`
   - Call Ollama to extract structured data points (market rents, vacancy, capital values, outlook)
   - Cache result to `<pdf_name>_stage1.json`; skip re-extraction if cache exists
3. **Stage 2 — Write**:
   - Call 1: Generate prose rationale sections (Executive Summary, Market Context, Subject Property, Investment Thesis, Risks)
   - Call 2: Audit sources — verify every statistic cited in Call 1 is traceable to Stage 1 extracts
4. Write `Investment_Rationale.md` and `Source_Audit.xlsx`
5. Write `Investment_Rationale_meta.json` with generation timestamp and elapsed seconds (controlled by `_SHOW_RATIONALE_TIMING` flag)

## Output format
| File | Description |
|---|---|
| `Investment_Rationale.md` | Full Markdown rationale document |
| `Source_Audit.xlsx` | Table linking each statistic to its source PDF and page |
| `Investment_Rationale_meta.json` | `{generated_at, elapsed_s}` for display in the dashboard |

## Examples
```bash
python3 generate_investment_rationale.py --config configs/deal_config_88_Cecil.json
python3 generate_investment_rationale.py --config configs/deal_config_88_Cecil.json --refine
```

## Notes
- Stage 1 output is cached — delete `*_stage1.json` files to force re-extraction from PDF
- `--refine` flag runs Stage 2 only (reuses cached Stage 1), useful for tweaking without re-reading PDFs
- Larger models (e.g. qwen2.5:14b, llama3.1:8b) produce better prose than small models
- `_SHOW_RATIONALE_TIMING = False` in `app.py` to hide the "Generated on … took …s" caption
