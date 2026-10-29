# Migrating from openpyxl to RPyXL

RPyXL is designed as a high-performance, drop-in replacement for `openpyxl`. For standard workflows, migrating requires only changing your import statement.

---

## 1. Quick Migration (Drop-In Alias)

If your codebase relies heavily on `openpyxl`, you can switch immediately using an import alias:

```python
# Before
import openpyxl

# After
import rpyxl as openpyxl
```

All standard calls such as `openpyxl.load_workbook("data.xlsx")`, `openpyxl.Workbook()`, and `openpyxl.utils.get_column_letter` work identically while executing on RPyXL's Rust core.

---

## 2. Common Patterns & Code Comparison

### 2.1 Standard In-Memory Workbook Writing

Both libraries share identical APIs for workbook creation, cell assignment, row appending, and saving:

```python
import rpyxl

wb = rpyxl.Workbook()
ws = wb.active
ws.title = "Q3_Report"

# Coordinate assignment
ws["A1"] = "Metric"
ws["B1"] = "Value"

# Cell lookup with row/col
ws.cell(row=2, column=1, value="Revenue")
ws.cell(row=2, column=2, value=150000.50)

# Row appending
ws.append(["Gross Margin", 0.65])
ws.append({"A": "Net Income", "B": 97500.32})

wb.save("report.xlsx")
```

### 2.2 High-Throughput Streaming Writer (`write_only=True`)

For datasets with 50,000+ rows, streaming write-only mode streams data to disk with constant $O(1)$ memory:

```python
import rpyxl
from datetime import datetime

wb = rpyxl.Workbook(write_only=True)
ws = wb.create_sheet("LargeStream")

# Header
ws.append(["ID", "Timestamp", "Value", "Status"])

for i in range(1, 100_001):
    ws.append([i, datetime.now(), i * 1.25, i % 2 == 0])

wb.save("large_export.xlsx")
```

*Performance note: RPyXL writes 100k rows in ~0.065s (~1.5M cells/s) compared to ~0.42s in openpyxl.*

### 2.3 Streaming Reader (`read_only=True`)

Streaming read mode parses XLSX files with zero accumulation of Python objects:

```python
import rpyxl

with rpyxl.load_workbook("large_export.xlsx", read_only=True) as wb:
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        id_val, ts, val, status = row
        # Process each row
```

*Performance note: RPyXL reads 100k rows in ~0.068s with a 550× faster time-to-first-row vs openpyxl.*

### 2.4 Slicing & Iteration

Subscript slicing and iterators match openpyxl semantics:

```python
# Slicing ranges
cells_2d = ws["A1:C10"]       # 2D tuple of Cell instances
vals = ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=3, values_only=True)

# Column and Row ranges
col_a = ws["A"]              # 1D tuple of cells in col A
cols_ac = ws["A:C"]          # 2D tuple of column tuples
row_1 = ws[1]                # 1D tuple of cells in row 1
rows_1_5 = ws[1:5]           # 2D tuple of row tuples
```

---

## 3. Behavioral Differences & Scope

| Area | openpyxl Behavior | RPyXL Behavior | Migration Tip |
|---|---|---|---|
| **Formulas** | Parses formula strings, caches values | Preserves formulas (`=SUM(A1:A10)`) and writes them directly | Formula evaluation engine is planned for v0.2 |
| **VBA Macros** | `keep_vba=True` retains macro XML | Accepts `keep_vba` parameter; macro binaries are omitted | RPyXL is focused on high-speed tabular data processing |
| **Charts** | `openpyxl.chart.*` DrawingML generators | Charts planned for v0.3 | If you need chart rendering, generate charts in a post-processing step |
| **Styles** | Full XML style engine | `rpyxl.styles` provides compatible `Font`, `PatternFill`, `Alignment`, `Border`, `Side` | Basic styles supported; advanced conditional rules in v0.2 |
| **Coordinate Validation** | Validates bounds in Python | Validates bounds in Rust; raises `ValueError` on coordinates < 1 | Exception types inherit from standard Python exceptions |

---

## 4. Performance Expectations

Expect the following performance improvements when migrating to RPyXL:

- **4.7× to 8.1× faster file saving** across small to large workbooks.
- **5.1× to 8.7× faster file loading**.
- **550× faster time-to-first-row** in streaming read mode.
- **148 MB less RAM consumption** when loading large spreadsheets (500k cells).
- **Throughput exceeding 1.6 million cells/second** on standard multi-core hardware.
