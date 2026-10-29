# Compatibility Matrix

This document tracks API compatibility between openpyxl and RPyXL.

## Phase 1 Compatibility Status

| API / Feature | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences / Notes |
|---|---|---|---|---|---|
| `Workbook()` | Creates in-memory workbook with default sheet | Matches | Supported | `test_workbook.py`, `test_openpyxl_compat.py` | None |
| `wb.active` | Returns active worksheet; allows assignment | Matches | Supported | `test_workbook.py` | Supports int, title, or Worksheet object |
| `wb.sheetnames` | Returns list of sheet titles | Matches | Supported | `test_workbook.py` | Preserves deterministic order |
| `wb.create_sheet(title, index)` | Creates sheet, auto-titles if None | Matches | Supported | `test_workbook.py` | Same validation and auto-numbering |
| `wb.remove(worksheet)` | Removes sheet; raises ValueError on only sheet | Matches | Supported | `test_workbook.py` | Matches |
| `wb[title]` | Gets worksheet by name; raises KeyError if missing | Matches | Supported | `test_workbook.py` | Matches |
| `title in wb` | Checks if worksheet exists | Matches | Supported | `test_workbook.py` | Matches |
| `len(wb)` | Number of worksheets | Matches | Supported | `test_workbook.py` | Matches |
| `for ws in wb` | Iterates worksheets | Matches | Supported | `test_workbook.py` | Matches |
| `wb.close()` | Closes file resources | Matches | Supported | `test_workbook.py` | In-memory no-op |
| `ws.title` | Get/set worksheet title | Matches | Supported | `test_worksheet.py` | Validates forbidden chars `: \ / ? * [ ]` |
| `ws.cell(row, col, value)` | 1-based coordinate access/set | Matches | Supported | `test_cell.py`, `test_worksheet.py` | Enforces Excel bounds (1-1048576, 1-16384) |
| `ws["A1"]` | Single cell coordinate lookup | Matches | Supported | `test_cell.py` | Case-insensitive column letters |
| `ws["A1"] = val` | Cell assignment | Matches | Supported | `test_cell.py` | Supports str, int, float, bool, None |
| `ws["A1:B2"]` | Range slice returning 2D tuple of cells | Matches | Supported | `test_worksheet.py` | Exact tuple of tuples |
| `ws.append(iterable)` | Appends row to max_row + 1 | Matches | Supported | `test_worksheet.py` | Bounded memory |
| `ws.iter_rows()` | Row iterator with bounds and values_only | Matches | Supported | `test_worksheet.py` | Supports min_row, max_row, min_col, max_col |
| `ws.iter_cols()` | Column iterator with bounds | Matches | Supported | `test_worksheet.py` | Matches |
| `ws.values` | Property yielding values_only row tuples | Matches | Supported | `test_worksheet.py` | Matches |
| `ws.rows` | Property yielding Cell row tuples | Matches | Supported | `test_worksheet.py` | Matches |
| `ws.columns` | Property yielding Cell col tuples | Matches | Supported | `test_worksheet.py` | Matches |
| `ws.dimensions` | Coordinate range string ("A1" or "A1:B10") | Matches | Supported | `test_worksheet.py` | Matches |
| `cell.value` | Get/set cell value | Matches | Supported | `test_cell.py` | Direct mutation reflected in sheet |
| `cell.coordinate` | "A1" format | Matches | Supported | `test_cell.py` | Matches |
| `cell.row`, `cell.column`| 1-based indices | Matches | Supported | `test_cell.py` | Matches |
| `cell.column_letter` | Column letter | Matches | Supported | `test_cell.py` | Matches |
| `cell.data_type` | 's', 'n', 'b', 'f' | Matches | Supported | `test_cell.py` | Matches |
| `repr(cell)` | `<Cell 'Sheet'.A1>` | Matches | Supported | `test_cell.py` | Exact match |
