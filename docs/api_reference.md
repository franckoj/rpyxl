# RPyXL API Reference Manual

This document details the public API surface of RPyXL (`rpyxl`).

---

## 1. Top-Level Functions

### `rpyxl.load_workbook(filename, read_only=False, data_only=False, keep_vba=False)`
Loads an existing XLSX workbook from disk.

- **Parameters**:
  - `filename` (*str | Path*): Path to the `.xlsx` file.
  - `read_only` (*bool*, default `False`): If `True`, enables lazy, low-memory streaming mode.
  - `data_only` (*bool*, default `False`): Controls whether formulas return their cached results.
  - `keep_vba` (*bool*, default `False`): Parameter accepted for openpyxl signature compatibility.
- **Returns**: `Workbook` or `ReadOnlyWorkbook`.
- **Raises**: `FileNotFoundError`, `InvalidFileException`, `ValueError`.

### `rpyxl.Workbook(write_only=False)`
Instantiates a new workbook.

- **Parameters**:
  - `write_only` (*bool*, default `False`): If `True`, creates a write-only streaming workbook with constant memory consumption.
- **Returns**: `Workbook` or `WriteOnlyWorkbook`.

---

## 2. `rpyxl.Workbook` (Standard In-Memory)

### Properties
- `active`: (*Worksheet*) Returns or sets the active worksheet. Can be assigned an integer index, title string, or `Worksheet` instance.
- `sheetnames`: (*list[str]*) Returns list of sheet names in order.
- `worksheets`: (*list[Worksheet]*) Returns list of `Worksheet` objects in order.

### Methods
- `create_sheet(title=None, index=None) -> Worksheet`:
  Creates and inserts a new worksheet at optional 0-based position `index`. If `title` is `None`, defaults to `"SheetN"`.
- `remove(worksheet: Worksheet) -> None`:
  Removes a worksheet. Raises `ValueError` if attempting to remove the sole remaining sheet.
- `copy_worksheet(from_worksheet: Worksheet) -> Worksheet`:
  Clones a worksheet and all contained cells, naming the copy `"{title} Copy"`.
- `index(worksheet: Worksheet) -> int`:
  Returns the 0-based index of the worksheet within this workbook. Raises `ValueError` if the sheet belongs to another workbook.
- `save(filename: str | Path) -> None`:
  Serializes the workbook to an ECMA-376 OOXML `.xlsx` ZIP archive.
- `close() -> None`:
  Closes underlying resources.
- `__getitem__(title: str) -> Worksheet`:
  Retrieves a worksheet by name (`wb["Sheet1"]`). Raises `KeyError` if absent.
- `__contains__(title: str) -> bool`:
  Checks if a sheet exists (`"Sheet1" in wb`).
- `__iter__()`:
  Iterates over worksheets in order.
- `__len__() -> int`:
  Returns total number of worksheets.

---

## 3. `rpyxl.Worksheet`

### Properties
- `title`: (*str*) Title of the sheet. Enforces length <= 31 chars and rejects forbidden characters (`\ / ? * : [ ]`).
- `parent`: (*Workbook*) Reference back to the parent `Workbook`.
- `max_row`: (*int*) 1-based index of maximum row containing data (1 if empty).
- `max_column`: (*int*) 1-based index of maximum column containing data (1 if empty).
- `min_row`: (*int*) 1-based index of minimum row (1 if empty).
- `min_column`: (*int*) 1-based index of minimum column (1 if empty).
- `dimensions`: (*str*) Range coordinate string representing bounds (e.g. `"A1:D10"`).
- `values`: (*generator*) Generator yielding row tuples of raw cell values.
- `rows`: (*tuple*) 2D tuple of `Cell` objects by row.
- `columns`: (*tuple*) 2D tuple of `Cell` objects by column.

### Methods
- `cell(row: int, column: int, value=None) -> Cell`:
  Accesses or assigns a cell by 1-based coordinate. Raises `ValueError` if `row < 1` or `column < 1`.
- `append(iterable: list | tuple | dict) -> None`:
  Appends a row below `max_row`. Supports `dict` with column letters (`"A"`) or 1-based integers (`1`).
- `iter_rows(min_row=None, max_row=None, min_col=None, max_col=None, values_only=False)`:
  Row-by-row iterator.
- `iter_cols(min_col=None, max_col=None, min_row=None, max_row=None, values_only=False)`:
  Column-by-column iterator.
- `__getitem__(key)`:
  Subscript slicing supporting:
  - `ws["A1"]` -> single `Cell`
  - `ws["A1:C5"]` -> 2D tuple of `Cell`s
  - `ws["A"]` -> 1D tuple of `Cell`s
  - `ws["A:C"]` -> 2D tuple of column `Cell` tuples
  - `ws[1]` -> 1D tuple of `Cell`s
  - `ws["1:5"]` / `ws[1:5]` -> 2D tuple of row `Cell` tuples

---

## 4. `rpyxl.Cell`

### Properties
- `value`: (*Any*) Get or set cell value (`str`, `int`, `float`, `bool`, `date`, `datetime`, `None`).
- `coordinate`: (*str*) Coordinate string (e.g. `"A1"`).
- `row`: (*int*) 1-based row index.
- `column`: (*int*) 1-based column index.
- `column_letter`: (*str*) Column letter (e.g. `"A"`, `"BC"`).
- `data_type`: (*str*) Openpyxl single-character code (`'s'`, `'n'`, `'b'`, `'f'`, `'d'`, `'null'`).
- `is_date`: (*bool*) `True` if value is a date or datetime object.
- `parent`: (*Worksheet*) Reference back to containing `Worksheet`.
- `number_format`: (*str*) Number format string (e.g. `"yyyy-mm-dd"`, `"0.00%"`).
- `comment`: (*str*) Cell comment string.
- `hyperlink`: (*str*) Hyperlink URL target.

---

## 5. Streaming Workbooks

### `rpyxl.Workbook(write_only=True)`
- High-performance, streaming writer with constant $O(1)$ memory usage.
- `ws = wb.create_sheet(title)`: Creates streaming sheet.
- `ws.append(row)`: Flushes row directly to temporary XML buffer on disk.
- `wb.save(filename)`: Finalizes archive.

### `rpyxl.load_workbook(filename, read_only=True)`
- Lazy streaming reader powered by `calamine`.
- `ws.iter_rows(values_only=True)` yields rows on-the-fly without accumulating cells in RAM.

---

## 6. Utilities & Exceptions

### `rpyxl.utils`
- `get_column_letter(idx: int) -> str`
- `column_index_from_string(col: str) -> int`
- `coordinate_to_tuple(coord: str) -> tuple[int, int]`
- `range_boundaries(range_str: str) -> tuple[int, int, int, int]`
- `quote_sheetname(sheetname: str) -> str`

### `rpyxl.utils.exceptions`
- `OpenpyxlException` (base)
- `InvalidFileException`
- `SheetTitleException`
- `CellCoordinatesException`
- `NamedRangeException`

### `rpyxl.styles`
- `Font(name, size, bold, italic, color)`
- `PatternFill(fill_type, start_color, end_color)`
- `Alignment(horizontal, vertical, wrap_text)`
- `Border(left, right, top, bottom)`
- `Side(style, color)`
- `Color(rgb)`
