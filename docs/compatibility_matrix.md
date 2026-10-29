# RPyXL API Compatibility Matrix (vs openpyxl)

This document tracks API compatibility between `openpyxl` (v3.1+) and `RPyXL` (v0.1.0).

Status Definitions:
- **Supported**: Full behavioural parity with openpyxl, verified by unit and cross-library tests.
- **Partial**: Core functionality matches; advanced edge cases or metadata may differ.
- **Planned**: Targeted for subsequent phases (v0.2+).
- **Not Planned**: Explicitly out of scope for RPyXL v0.1 (e.g., VBA macros, chart rendering).

---

## 1. Top-Level Functions

| API Name | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences | Version |
|---|---|---|---|---|---|---|
| `load_workbook(filename, read_only=False, data_only=False, keep_vba=False)` | Opens an XLSX file into a Workbook; supports streaming read-only and evaluating formulas to cached values | Reads XLSX via Rust calamine engine; full support for standard & streaming `read_only=True` mode, `data_only=True` | **Supported** | `test_reader.py`, `test_reader_compat.py` | `keep_vba` accepted as parameter but VBA macros not executed | v0.1.0 |
| `Workbook(write_only=False)` | Instantiates a new in-memory Workbook or streaming write-only Workbook | Instantiates Rust-backed in-memory `Workbook` or streaming `WriteOnlyWorkbook` | **Supported** | `test_workbook.py`, `test_writer.py`, `test_phase4_compat.py` | None |

---

## 2. Workbook APIs (`rpyxl.Workbook` & `rpyxl.workbook.Workbook`)

| API Name | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences | Version |
|---|---|---|---|---|---|---|
| `wb.active` | Returns active worksheet; allows assigning int index, title str, or Worksheet object | Returns active worksheet; supports assignment by int, title str, or `Worksheet` | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `wb.sheetnames` | Returns ordered list of worksheet title strings | Returns ordered list of worksheet title strings | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `wb.worksheets` | Returns ordered list of `Worksheet` objects | Returns ordered list of `Worksheet` objects | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `wb.create_sheet(title=None, index=None)` | Creates a new sheet at position index; defaults to "SheetN" | Creates sheet at optional index; auto-numbers duplicate/unspecified names | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `wb.remove(worksheet)` | Removes sheet; raises ValueError if trying to remove sole sheet | Removes sheet; raises `ValueError` if sole sheet | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `wb.copy_worksheet(from_worksheet)` | Clones worksheet cells, values, and styles, naming it "Title Copy" | Clones worksheet with cells, values, formulas, comments, and hyperlinks | **Supported** | `test_phase4_compat.py` | None |
| `wb.index(worksheet)` | Returns 0-based index of worksheet; raises `ValueError` if sheet from other wb | Returns 0-based index; validates pointer identity against parent workbook | **Supported** | `test_phase4_compat.py` | None |
| `wb[title]` | Returns worksheet by name; raises `KeyError` if absent | Returns worksheet by name; raises `KeyError` if absent | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `title in wb` | Boolean check if sheet title exists | Boolean check if sheet title exists | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `len(wb)` | Returns number of worksheets | Returns number of worksheets | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `for ws in wb` | Iterates over worksheets in order | Iterates over worksheets in order | **Supported** | `test_workbook.py`, `test_phase4_compat.py` | None |
| `wb.save(filename)` | Serializes workbook to standard XLSX ZIP package | High-performance Rust OOXML serializer writing valid ZIP archive | **Supported** | `test_writer.py`, `test_writer_compat.py` | ~4-6x faster than openpyxl | v0.1.0 |
| `wb.close()` | Closes file resources | Closes underlying file resources / context-manager cleanup | **Supported** | `test_workbook.py`, `test_writer.py` | None |
| `with Workbook() as wb:` | Context manager closing workbook upon exit | Context manager closing workbook upon exit | **Supported** | `test_workbook.py`, `test_writer.py` | None |

---

## 3. Worksheet APIs (`rpyxl.worksheet.worksheet.Worksheet`)

| API Name | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences | Version |
|---|---|---|---|---|---|---|
| `ws.title` | Get / set sheet title with validation (rejects invalid chars `: \ / ? * [ ]` and >31 chars) | Get / set sheet title with exact regex and length validation | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.parent` | Reference to parent `Workbook` | Reference to parent `Workbook` | **Supported** | `test_phase4_compat.py` | None |
| `ws.cell(row, column, value=None)` | 1-based coordinate cell lookup or assignment; raises `ValueError` if row/col < 1 | 1-based coordinate lookup; validates `row >= 1 && col >= 1`, optional value set | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.append(row)` | Appends row from `list`, `tuple`, or `dict` (keys: col letters or 1-based ints) | Supports `list`, `tuple`, and `dict` (letter & int keys, auto-padding gaps) | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.iter_rows(...)` | Iterates row tuples with `min_row`, `max_row`, `min_col`, `max_col`, `values_only` | Iterates row tuples; empty sheets yield empty iterator; values_only mode | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.iter_cols(...)` | Iterates col tuples with `(min_col, max_col, min_row, max_row, values_only)` | Positional and keyword parameter matching openpyxl signature | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.max_row` | Maximum row index (1 if empty) | Maximum row index (1 if empty) | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.max_column` | Maximum column index (1 if empty) | Maximum column index (1 if empty) | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.min_row` | Minimum row index (1 if empty) | Minimum row index (1 if empty) | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.min_column` | Minimum column index (1 if empty) | Minimum column index (1 if empty) | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.values` | Property yielding generator/iterator of row value tuples | Property yielding generator of row value tuples | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.rows` | Property yielding tuple of row cell tuples | Property yielding tuple of row cell tuples | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.columns` | Property yielding tuple of column cell tuples | Property yielding tuple of column cell tuples | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws.dimensions` | Coordinate range string (e.g. `"A1:D10"` or `"A1"`) | Coordinate range string | **Supported** | `test_worksheet.py` | None |
| `ws["A1"]` | Access single cell by coordinate string | Access single cell by coordinate string | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `ws["A1"] = val` | Assign cell value by coordinate string | Assign cell value by coordinate string | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `ws["A1:C5"]` | Slicing by coordinate range string -> 2D tuple `((cell, ...), ...)` | Slicing by coordinate range string -> 2D tuple | **Supported** | `test_worksheet.py`, `test_phase4_compat.py` | None |
| `ws["A1":"C5"]` | Slicing by coordinate slice -> 2D tuple | Slicing by coordinate slice -> 2D tuple | **Supported** | `test_phase4_compat.py` | None |
| `ws["A"]` | Column access by letter -> tuple of cells in column `(A1, A2, ...)` | Column access by letter -> 1D tuple of cells | **Supported** | `test_phase4_compat.py` | None |
| `ws["A:C"]` | Column range access -> 2D tuple of columns `((A1, A2), (B1, B2), ...)` | Column range access -> 2D tuple of columns | **Supported** | `test_phase4_compat.py` | None |
| `ws[1]` | Row access by 1-based integer -> tuple of cells `(A1, B1, ...)` | Row access by 1-based integer -> 1D tuple of cells | **Supported** | `test_phase4_compat.py` | None |
| `ws["1:5"]` | Row range access by string -> 2D tuple `((A1, B1), (A2, B2), ...)` | Row range access by string -> 2D tuple | **Supported** | `test_phase4_compat.py` | None |
| `ws[1:5]` | Row range access by slice -> 2D tuple `((A1, B1), ..., (A5, B5))` | Row range access by slice -> 2D tuple | **Supported** | `test_phase4_compat.py` | None |

---

## 4. Cell APIs (`rpyxl.cell.cell.Cell`)

| API Name | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences | Version |
|---|---|---|---|---|---|---|
| `cell.value` | Get / set cell value (`str`, `int`, `float`, `bool`, `date`, `datetime`, `None`) | Get / set cell value backed by Rust `CellValue` enum | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `cell.coordinate` | Coordinate string, e.g. `"A1"` | Coordinate string, e.g. `"A1"` | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `cell.row` | 1-based row index | 1-based row index | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `cell.column` | 1-based column index | 1-based column index | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `cell.column_letter` | Column letter, e.g. `"A"`, `"AA"` | Column letter, e.g. `"A"`, `"AA"` | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `cell.data_type` | Single character code: `'s'`, `'n'`, `'b'`, `'f'`, `'d'`, `'e'`, `'inlineStr'` | Matches openpyxl codes (`'s'`, `'n'`, `'b'`, `'f'`, `'d'`, `'null'`); allows set | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |
| `cell.is_date` | Returns `True` if cell holds `datetime.date` or `datetime.datetime` | Returns `True` if `CellValue::Date` or `CellValue::DateTime` | **Supported** | `test_phase4_compat.py` | None |
| `cell.parent` | Reference to parent `Worksheet` | Reference to parent `Worksheet` | **Supported** | `test_phase4_compat.py` | None |
| `cell.number_format` | Format string (e.g. `'yyyy-mm-dd'`, `'0.00%'`) | Get / set number format string | **Supported** | `test_phase4_compat.py` | OOXML format indexing mapped in serializer | v0.1.0 |
| `cell.comment` | Cell comment string or Comment object | Get / set cell comment string | **Supported** | `test_phase4_compat.py` | Rich text comment formatting simplified to string | v0.1.0 |
| `cell.hyperlink` | Hyperlink target URL string or Hyperlink object | Get / set cell hyperlink URL string | **Supported** | `test_phase4_compat.py` | Stored as direct URL string | v0.1.0 |
| `repr(cell)` | `<Cell 'Sheet'.A1>` | `<Cell 'Sheet'.A1>` | **Supported** | `test_cell.py`, `test_phase4_compat.py` | None |

---

## 5. Streaming & Write-Only APIs

| API Name | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences | Version |
|---|---|---|---|---|---|---|
| `Workbook(write_only=True)` | Creates streaming workbook optimized for minimal memory write | Zero-buffering Rust streaming workbook writing directly to ZIP | **Supported** | `test_writer.py`, `test_writer_compat.py` | ~7-10x memory reduction vs openpyxl | v0.1.0 |
| `load_workbook(..., read_only=True)` | Generator-based streaming reader consuming minimal memory | Rust calamine streaming reader with zero Python object accumulation | **Supported** | `test_reader.py`, `test_reader_compat.py` | ~10x memory reduction vs openpyxl | v0.1.0 |
| `ws.append(...)` in write-only | Appends row and flushes directly to disk XML stream | Appends row directly to temporary XML buffer inside Rust | **Supported** | `test_writer.py`, `test_writer_compat.py` | None |

---

## 6. Utilities & Exceptions

| Module / API | openpyxl Path | RPyXL Path | Status | Test Coverage | Notes | Version |
|---|---|---|---|---|---|---|
| `get_column_letter(idx)` | `openpyxl.utils` | `rpyxl.utils` | **Supported** | `test_phase4_compat.py` | High-speed Rust implementation | v0.1.0 |
| `column_index_from_string(col)` | `openpyxl.utils` | `rpyxl.utils` | **Supported** | `test_phase4_compat.py` | Case-insensitive Rust parser | v0.1.0 |
| `coordinate_to_tuple(coord)` | `openpyxl.utils` | `rpyxl.utils` | **Supported** | `test_phase4_compat.py` | Returns `(row, col)` | v0.1.0 |
| `range_boundaries(range_str)` | `openpyxl.utils` | `rpyxl.utils` | **Supported** | `test_phase4_compat.py` | Returns `(min_col, min_row, max_col, max_row)` | v0.1.0 |
| `quote_sheetname(sheetname)` | `openpyxl.utils` | `rpyxl.utils` | **Supported** | `test_phase4_compat.py` | Handles escaping single quotes | v0.1.0 |
| `InvalidFileException` | `openpyxl.utils.exceptions` | `rpyxl.utils.exceptions` | **Supported** | `test_phase4_compat.py` | Inherits `ValueError` & `OpenpyxlException` | v0.1.0 |
| `SheetTitleException` | `openpyxl.utils.exceptions` | `rpyxl.utils.exceptions` | **Supported** | `test_phase4_compat.py` | Inherits `ValueError` & `OpenpyxlException` | v0.1.0 |
| `CellCoordinatesException` | `openpyxl.utils.exceptions` | `rpyxl.utils.exceptions` | **Supported** | `test_phase4_compat.py` | Inherits `ValueError` & `OpenpyxlException` | v0.1.0 |
| `NamedRangeException` | `openpyxl.utils.exceptions` | `rpyxl.utils.exceptions` | **Supported** | `test_phase4_compat.py` | Inherits `OpenpyxlException` | v0.1.0 |

---

## 7. Styling Foundation (`rpyxl.styles`)

| API Name | openpyxl Behavior | RPyXL Behavior | Status | Test Coverage | Known Differences | Version |
|---|---|---|---|---|---|---|
| `Font(name, size, bold, italic, color)` | Configures font style | Openpyxl-compatible dataclass | **Supported** | `test_phase4_compat.py` | Preserved for API drop-in compatibility | v0.1.0 |
| `PatternFill(fill_type, start_color, end_color)` | Configures background fill pattern | Openpyxl-compatible dataclass | **Supported** | `test_phase4_compat.py` | Preserved for API drop-in compatibility | v0.1.0 |
| `Alignment(horizontal, vertical, wrap_text)` | Configures text alignment | Openpyxl-compatible dataclass | **Supported** | `test_phase4_compat.py` | Preserved for API drop-in compatibility | v0.1.0 |
| `Border(left, right, top, bottom)` | Configures cell borders | Openpyxl-compatible dataclass | **Supported** | `test_phase4_compat.py` | Preserved for API drop-in compatibility | v0.1.0 |
| `Side(style, color)` | Configures individual border side | Openpyxl-compatible dataclass | **Supported** | `test_phase4_compat.py` | Preserved for API drop-in compatibility | v0.1.0 |
| `Color(rgb)` | Configures aRGB hex color | Openpyxl-compatible dataclass | **Supported** | `test_phase4_compat.py` | Preserved for API drop-in compatibility | v0.1.0 |

---

## 8. Unsupported / Out of Scope (v0.1)

| Feature / API | Status | Reason / Roadmap |
|---|---|---|
| VBA Macros (`keep_vba=True` execution) | **Not Planned** | Security and complexity; macro execution out of scope for data-focused XLSX engines |
| Chart Creation & Rendering (`openpyxl.chart.*`) | **Planned (v0.3+)** | Requires complex DrawingML serialization; focus on tabular data first |
| Pivot Tables (`openpyxl.pivot.*`) | **Planned (v0.3+)** | Requires PivotCache and PivotTable definition XML streams |
| Conditional Formatting rules | **Planned (v0.2+)** | Style and rule evaluation planned for post-v0.1 |
