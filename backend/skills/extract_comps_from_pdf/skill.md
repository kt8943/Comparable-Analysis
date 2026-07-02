---
name: extract_comps_from_pdf
description: Extract comp records from a PDF using pdfplumber page discovery, table detection, and Ollama field mapping
type: atomic
requires:
  config_keys:
    - llm.ollama.base_url
    - llm.ollama.model
  skills: []
allowed_tools:
  - tools.calculations.parse_num
  - tools.calculations.parse_remaining_yrs
  - tools.llm_client.ollama_post
  - tools.json_utils.fix_json
---

## When to use
When the input source for a comp pipeline is a PDF file. Called inside `analyse_sales_comps`, `analyse_rent_comps`, and `analyse_land_comps` when the `input_file` extension is `.pdf`. Delegates to `pdf_extractor.extract_pdf_records` (4-stage shared pipeline).

## Instructions
1. Call `pdf_extractor.extract_pdf_records(pdf_path, section_keywords, field_schema, llm_cfg, subject_name=...)`
   - Stage 1: pdfplumber scans pages to find sections matching `section_keywords`
   - Stage 2: table detection extracts rows from identified sections
   - Stage 3: Ollama maps extracted column headers to `field_schema` keys
   - Stage 4: record assembly — applies `parse_num`, `parse_remaining_yrs`, stake-pct extraction
2. Post-process: filter out subject property rows, skip records with no price, normalise `price_sgd_m` (divide by 1,000,000 if raw value > 100,000)
3. Return list of raw comp dicts (same format as `parse_input_excel`)

## Output format
List of raw comp record dicts. Keys include: `property_name`, `address`, `gfa_sf`, `price_sgd_m`, `remaining_yrs`, `npi_yield`, `adj_npi_yield`, `sale_type`, `land_zoning`, `sale_date`, `stake_pct`, `_source: "pdf"`.

## Examples
```python
from pdf_extractor import extract_pdf_records

llm_cfg = {"ollama": {"base_url": base_url, "model": model}}
raw_records = extract_pdf_records(
    pdf_path, _PDF_SECTION_KEYWORDS, _PDF_FIELD_SCHEMA,
    llm_cfg, subject_name="88 Cecil Street",
)
```

## Notes
- Requires `pdfplumber` (`pip install pdfplumber`)
- Larger models (qwen2.5:14b, llama3.1:8b) produce more reliable column mapping than small models
- PDFs with scanned/image pages will not extract text — use `extract_comps_from_image` instead
- Subject property deduplication uses a token-overlap heuristic (≥ 60 % or ≥ 2 shared tokens)
