# RPyXL: A High-Performance, Rust-Powered Alternative to openpyxl

[![CI](https://github.com/franckoj/rpyxl/actions/workflows/ci.yml/badge.svg)](https://github.com/franckoj/rpyxl/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue.svg)](LICENSE)
[![Python Versions](https://img.shields.io/badge/python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://pypi.org/project/rpyxl/)
[![Rust](https://img.shields.io/badge/rust-2021%20edition-orange.svg)](https://www.rust-lang.org/)
[![ABI3](https://img.shields.io/badge/wheels-ABI3%20Universal-green.svg)](https://github.com/franckoj/rpyxl)

RPyXL is a high-performance Python spreadsheet library providing a drop-in **[openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)-compatible API**, powered by a memory-efficient Rust core built with [PyO3](https://pyo3.rs/) and [Maturin](https://www.maturin.rs/).

**Version 0.4.0** delivers high-leverage performance upgrades: native streaming row/column iterators, versioned cell value caching, zero-alloc batch row ingestion (`append_rows`), inline style XML caching, and configurable Deflate compression—delivering **5x to 9x speedups** over openpyxl 3.1.5 while maintaining **100% roundtrip data fidelity**.

---

## Performance Highlights: openpyxl 3.1.5 vs RPyXL 0.4.0

The following live head-to-head benchmark compares `openpyxl 3.1.5` against `rpyxl 0.4.0` on Python 3.12 (Apple Silicon, 8 CPU cores):

### 1. Ingestion, Export & Mutation Throughput

| Operation / Workload Scale | openpyxl (3.1.5) | RPyXL (0.4.0) | Speedup | Throughput (RPyXL) |
| :--- | :---: | :---: | :---: | :---: |
| **Standard Export / Write** (500k cells) | 15.94 s | **2.21 s** | **7.22x faster** | 226,338 cells/s |
| **Streaming Write-Only** (500k cells) | 13.17 s | **1.78 s** | **7.39x faster** | 280,537 cells/s |
| **Standard Read / Ingestion** (500k cells) | 10.94 s | **1.81 s** | **6.04x faster** | 275,971 cells/s |
| **Streaming Read-Only** (500k cells) | 7.97 s | **1.32 s** | **6.04x faster** | 379,250 cells/s |
| **Time to First Row (TTFR)** | 12.68 ms | **1.39 ms** | **9.11x faster** | Instantaneous |
| **Read-Modify-Write Cycle** (100k cells) | 4.67 s | **0.84 s** | **5.53x faster** | 118,425 cells/s |
| **Batch Row Ingestion** (`append_rows()`) | 3.68 s | **0.47 s** | **7.83x faster** | **2,127,428 cells/s** |
| **Streaming Value Iteration** (`iter_rows`) | 1.10 s | **0.23 s** | **4.77x faster** | **2,164,305 cells/s** |

### 2. Cross-Library Roundtrip Fidelity

| Serialization Direction | Cells Tested | Content Verified | Mismatches | Fidelity Status |
| :--- | :---: | :--- | :---: | :---: |
| **RPyXL Write $\rightarrow$ openpyxl Read** | 5,000 cells | String, Int, Float, Bool, Datetime | **0** | **100% PERFECT** |
| **openpyxl Write $\rightarrow$ RPyXL Read** | 5,000 cells | String, Int, Float, Bool, Datetime | **0** | **100% PERFECT** |

---

## What's New in Version 0.4.0 (vs 0.3.0)

| Capability / Subsystem | Prior Release (0.3.0) | Current Release (0.4.0) | Optimization Impact |
| :--- | :---: | :---: | :--- |
| **Cell Value Read** (`cell.value`) | 354,050 cells/s | **592,130 cells/s** | 🚀 **1.67x faster**: Versioned value cache skips sheet locks |
| **Cell Value Write** (`cell.value = x`) | 227,268 cells/s | **325,804 cells/s** | 🚀 **1.43x faster**: Direct integer coordinate bounds check |
| **Row Traversal** (`iter_rows`) | 548,026 cells/s | **817,828 cells/s** | 🚀 **1.49x faster**: Native streaming Rust iterators |
| **Styled XLSX File Size** (25k rows) | 1.066 MB | **0.705 MB** | 📉 **33.8% smaller archive**: Configurable level 6 Deflate |
| **Styled Save Serialization** | Per-cell map lookup | Inline style cache | ⚡ Consecutive cell styles bypass style table lookup |
| **Merged Cell Traversal** | Scanned merge ranges | `_may_have_merges` | ⚡ Non-merged sheets bypass merge searches entirely |

---

## Migration for Supported Workflows

For existing projects using `openpyxl`, simply swap the import:

```python
# Before
import openpyxl

# After
import rpyxl as openpyxl

# Standard workflow works identically
wb = openpyxl.load_workbook("financial_report.xlsx", data_only=True)
ws = wb.active
print(ws["A1"].value)
```

See the [Migration Guide](docs/migration_guide.md) for comprehensive patterns and edge cases.

---

## Quick Start

### 1. In-Memory Workbook Writing

```python
import rpyxl

wb = rpyxl.Workbook()
ws = wb.active
ws.title = "Sales"

ws["A1"] = "Employee"
ws["B1"] = "Revenue"
ws.append(["Alice", 15400.50])
ws.append(["Bob", 22900.75])

# Subscript access and coordinates
ws.cell(row=4, column=1, value="Charlie")
ws.cell(row=4, column=2, value=18300.00)

wb.save("sales_report.xlsx")
```

### 2. High-Speed Batch Row Ingestion (`append_rows`)

Ingest massive tabular datasets at over **2,000,000+ cells/sec** using zero-allocation batch buffering:

```python
import rpyxl

wb = rpyxl.Workbook()
ws = wb.active

records = [
    [f"ID_{i:06d}", i * 10, round(i * 1.25, 2), i % 2 == 0]
    for i in range(1, 100_001)
]

# Ingest all rows in a single native pass
ws.append_rows(records)  # Alias: ws.extend(records)

# Configurable compression level (1 for ultra-fast saves, 6 for standard compact files)
wb.save("bulk_data.xlsx", compression_level=6)
```

### 3. High-Throughput Streaming Writer (`write_only=True`)

Low-memory streaming writer for huge datasets:

```python
import rpyxl
from datetime import datetime

wb = rpyxl.Workbook(write_only=True)
ws = wb.create_sheet("Transactions")

ws.append(["Transaction_ID", "Timestamp", "Amount", "Verified"])

for i in range(1, 100_001):
    ws.append([f"TX_{i:06d}", datetime.now(), i * 19.99, i % 2 == 0])

wb.save("transactions.xlsx")
```

### 4. Streaming Reader (`read_only=True`)

Iterate worksheet rows on demand with minimal memory footprint:

```python
import rpyxl

with rpyxl.load_workbook("transactions.xlsx", read_only=True) as wb:
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        tx_id, ts, amount, verified = row
        # Process record
```

### 5. Advanced Slicing & Iteration

```python
# 2D tuple of Cell instances
grid = ws["A1:C10"]

# Native streaming row and column iterators
for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
    for cell in row:
        val = cell.value

# Column and Row ranges
col_a = ws["A"]       # 1D tuple of cells in column A
cols_ab = ws["A:B"]   # 2D tuple of column tuples
row_1 = ws[1]         # 1D tuple of cells in row 1
rows_1_5 = ws[1:5]    # 2D tuple of row tuples
```

---

## Installation

### From Pre-Built Wheels (ABI3)

```bash
pip install rpyxl
```

_Requires Python 3.9 or newer. ABI3 wheels target Linux (x86_64 and aarch64), macOS (Apple Silicon and Intel), and Windows x64._

### From Source

```bash
git clone https://github.com/franckoj/rpyxl.git
cd rpyxl
pip install maturin
maturin develop --release
```

---

## Testing & Quality Gates

Deterministic quality gates ensure 100% test passing, clean linting, and formatting:

```bash
# 1. Python Linting & Formatting (Ruff)
ruff check python/ tests/ benchmarks/
ruff format --check python/ tests/ benchmarks/

# 2. Rust Formatting & Clippy Linter
cargo fmt --all -- --check
cargo clippy --workspace --all-targets -- -D warnings

# 3. Full Automated Test Suite (498 tests: 432 unit + 66 compatibility)
pytest tests/unit/ tests/compatibility/

# 4. Comprehensive Benchmark Suite
python benchmarks/live_benchmark.py
python benchmarks/bench_batch_append.py
python benchmarks/bench_styled_sweep.py
```

---

## Architecture Overview

```mermaid
graph TD
    A[Python Application] -->|Import rpyxl| B[Python Compatibility Layer]
    B --> C[rpyxl._core PyO3 Bindings]
    C --> D[rpyxl-core: Sparse BTreeMap Cell Store]
    C --> E[rpyxl-ooxml: Native XML Reader and Writer]
    E --> F[Direct ZIP & XML Streaming]
    D -->|Rayon Multi-Threading| E
```

- **`rpyxl-core`**: Pure Rust workbook and sparse cell coordinate model (`BTreeMap<(row, col), Cell>`).
- **`rpyxl-ooxml`**: High-speed ECMA-376 XML parser/generator with `itoa`, `ryu`, and `rayon`.
- **`rpyxl-python`**: Thread-safe PyO3 bindings with ABI3 multi-Python version support and native streaming iterators.
- **`python/rpyxl`**: Python public layer replicating openpyxl classes, styles, and exceptions with versioned value caching.

---

## Documentation & Phase Reports

- [Release Changelog](CHANGELOG.md): Exhaustive release history and version notes.
- [Development Changelog](docs/development_progress.md): Chronological engineering milestones.
- [API Reference Manual](docs/api_reference.md): Detailed signatures and parameter reference.
- [openpyxl Migration Guide](docs/migration_guide.md): Drop-in patterns, edge cases, and best practices.
- [API Compatibility Matrix](docs/compatibility_matrix.md): Exhaustive openpyxl vs RPyXL feature matrix.
- [Security Architecture & Threat Model](docs/security.md): Hardening against ZIP bombs, XXE, and malformed files.
- [Future Backend Plan](docs/future_backend_plan.md): Staged writer and reader backend evaluation.

---

## License

Licensed under either of:

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE) or http://www.apache.org/licenses/LICENSE-2.0)
- MIT license ([LICENSE-MIT](LICENSE) or http://opensource.org/licenses/MIT)

at your option.
