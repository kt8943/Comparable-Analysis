---
name: parse_input_excel
description: Detect the best sheet in an input Excel workbook, find the header row, and Ollama-map columns to output schema keys
type: atomic
requires:
  config_keys:
    - input_file
    - llm.ollama.base_url
    - llm.ollama.model
  skills: []
allowed_tools:
  - tools.excel_reader.find_best_sheet
  - tools.excel_reader.find_header_row
  - tools.excel_reader.sheet_keywords
  - tools.column_mapper.map_columns_ollama
  - tools.calculations.parse_num
  - tools.calculations.parse_remaining_yrs
  - tools.calculations.parse_sale_date
---

## When to use
When the input source for a comp pipeline is an Excel file. Called in Stage 1 of `analyse_sales_comps`, `analyse_rent_comps`, and `analyse_land_comps` when the `input_file` extension is `.xlsx` or `.xls`.

## Instructions
1. Open the workbook with `openpyxl`
2. Call `tools.excel_reader.sheet_keywords(output_fields)` to build the keyword set for the comp type
3. Call `tools.excel_reader.find_best_sheet(wb, keywords)` → returns the sheet name most likely to contain comp data
4. Read all rows from the selected sheet
5. Call `tools.excel_reader.find_header_row(rows)` → returns 0-based index of the header row
6. Extract `headers` (row at header index) and `sample_rows` (next 3 rows) for the LLM
7. Call `tools.column_mapper.map_columns_ollama(headers, sample_rows, output_fields, col_to_key, base_url, model)` → returns `{internal_key: col_index}` mapping
8. Iterate remaining rows; for each row apply `parse_num`, `parse_remaining_yrs`, `parse_sale_date` to convert cell values to typed Python values
9. Return list of raw comp dicts

## Output format
List of raw comp record dicts. Keys match the `internal_key` values from `_OUTPUT_FIELDS` for the relevant comp type (e.g. `property_name`, `sale_date_raw`, `gfa_sf`, `price_sgd_m`, …). No classification, geocoding, or metric computation yet.

## Examples
```python
import openpyxl
from tools.excel_reader import sheet_keywords, find_best_sheet, find_header_row
from tools.column_mapper import map_columns_ollama

wb = openpyxl.load_workbook("input.xlsx", data_only=True)
kws = sheet_keywords(OUTPUT_FIELDS)
sheet_name = find_best_sheet(wb, kws)
ws = wb[sheet_name]
rows = [[c.value for c in row] for row in ws.iter_rows()]
header_idx = find_header_row(rows)
headers = rows[header_idx]
mapping = map_columns_ollama(headers, rows[header_idx+1:header_idx+4],
                              OUTPUT_FIELDS, COL_TO_KEY, base_url, model)
```

## Notes
- Sheet selection scores sheets by how many output column name keywords appear in each row — avoids picking pivot or summary sheets
- Column mapping calls Ollama with a JSON-structured prompt; falls back to keyword matching if Ollama is unreachable
- `find_header_row` returns row 0 if no row with ≥ 3 text cells is found — safe default for files with a single header row at the top
