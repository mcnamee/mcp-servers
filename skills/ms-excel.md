---
name: ms-excel
description: Read and analyse Excel workbooks via the excel MCP server (read-only .xlsx/.xlsm). Use when the user asks what's in a spreadsheet, wants data read, searched, or summarised from Excel files.
---

# Excel (via the `excel` MCP server)

Requires the `ms-excel.py` MCP server (READ-ONLY; parses .xlsx directly, no
Excel needed). If its tools are not available, tell the user to wire it in
first (see the repo README) and to verify with `python ms-excel.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `excel_list_workbooks` | What workbooks are available |
| `excel_list_sheets` | Sheet names + dimensions of one workbook |
| `excel_get_headers` | The header row of a sheet |
| `excel_read_range` | Read cells (optionally an A1 range like `A1:D50`) |
| `excel_search` | Find cells matching a query across a workbook |
| `excel_column_stats` | Sum/average/min/max of one column |

## Workflow

Always orient before reading — workbook names resolve fuzzily, so:

1. `excel_list_workbooks` → confirm the workbook exists.
2. `excel_list_sheets` → pick the sheet.
3. `excel_get_headers` → learn the columns before reading data.
4. Then `excel_read_range` for data, `excel_search` to locate values, or
   `excel_column_stats` for numeric summaries (prefer it over reading whole
   columns to compute stats yourself).

## Notes

- Strictly read-only: it cannot write cells, create files or run macros —
  never promise to update a spreadsheet.
- Reads are capped (rows/columns/search hits) to protect the context
  window; page through large ranges rather than requesting everything.
- Formula cells return the value Excel last cached, not a re-evaluation.
  Legacy `.xls`/`.xlsb` and password-protected files are unsupported and
  reported as such.
