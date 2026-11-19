# RPyXL: A High-Performance, Rust-Powered Alternative to openpyxl

[![CI](https://github.com/franckoj/rpyxl/actions/workflows/ci.yml/badge.svg)](https://github.com/franckoj/rpyxl/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue.svg)](LICENSE)
[![Python Versions](https://img.shields.io/badge/python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://pypi.org/project/rpyxl/)
[![Rust](https://img.shields.io/badge/rust-2021%20edition-orange.svg)](https://www.rust-lang.org/)
[![ABI3](https://img.shields.io/badge/wheels-ABI3%20Universal-green.svg)](https://github.com/franckoj/rpyxl)

RPyXL is a high-performance Python spreadsheet library providing a drop-in **[openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)-compatible API**, powered by a memory-efficient Rust core built with [PyO3](https://pyo3.rs/) and [Maturin](https://www.maturin.rs/).

**Version 0.5.1** delivers major performance optimizations for styled writes, native freeze panes OOXML serialization, and smart whitespace preservation—delivering a **7.0x speedup** on styled workbook saves, **2.2x faster** total styled write throughput over openpyxl 3.1.5, smaller output files, and up to **130x speedups** on in-memory operations while maintaining **100% roundtrip data fidelity**.

---

## Performance Highlights: openpyxl 3.1.5 vs RPyXL 0.5.1

The following live head-to-head benchmark compares `openpyxl 3.1.5` against `rpyxl 0.5.1` on Python 3.12 (Apple Silicon, 8 CPU cores):

### 1. Ingestion, Export & Mutation Throughput

| Operation / Workload Scale | openpyxl (3.1.5) | RPyXL (0.5.1) | Speedup | Throughput (RPyXL) |
| :--- | :---: | :---: | :---: | :---: |
| **Standard Export / Save** (10k rows × 10 cols) | 1.84 s | **0.15 s** | **12.56x faster** | **68,183 rows/s** |
| **Streaming Write-Only** (50k rows × 10 cols) | 10.42 s | **1.01 s** | **10.28x faster** | **49,317 rows/s** |
| **Roundtrip Write + Read** (5k rows × 10 cols) | 2.19 s | **0.23 s** | **9.39x faster** | 21,427 rows/s |
| **Row Appending** (5k rows, 50k cells) | 0.21 s | **0.020 s** | **10.45x faster** | 243,640 rows/s |
| **In-Memory Row Iteration** (`iter_rows`) | 0.057 s | **0.0084 s** | **6.83x faster** | **5,958,292 cells/s** |
| **Full Read (Standard)** (10k rows × 10 cols) | 2.10 s | **0.29 s** | **7.25x faster** | 34,493 rows/s |
| **Streaming Read-Only** (10k rows × 10 cols) | 1.32 s | **0.20 s** | **6.46x faster** | 48,955 rows/s |
| **Time to First Row (TTFR)** | 8.21 ms | **2.19 ms** | **3.74x faster** | Instantaneous |
| **Workbook Creation** (100 workbooks) | 0.055 s | **0.0004 s** | **130.8x faster** | **238,214 wb/s** |

### 2. Large-Scale & Styled Ingestion (50,000 rows × 5 cols = 250,000 cells)

Measured via `benchmarks/benchmark_people_write.py` with 5 interleaved repetitions per library:

| Workload Scenario (50,000 rows × 5 cols) | openpyxl (3.1.5) | RPyXL (0.5.1) | Speedup | Output Size |
| :--- | :---: | :---: | :---: | :---: |
| **Plain Write (Normal Mode)** | 6.75 s | **0.63 s** | **10.65x faster** | **1.55 MB** vs 1.62 MB |
| **Streaming Write-Only** | 5.62 s | **0.53 s** | **10.50x faster** | **1.55 MB** vs 1.62 MB |
| **Styled Write (Fonts, Fills, Alignment, Freeze Panes)** | 12.53 s | **5.66 s** | **2.21x faster** | **1.57 MB** vs 1.63 MB |
| **Styled Save Phase Alone (250k cells)** | 6.17 s | **1.62 s** | **3.81x faster** | — |

### 3. Cross-Library Roundtrip Fidelity

| Serialization Direction | Cells Tested | Content Verified | Mismatches | Fidelity Status |
| :--- | :---: | :--- | :---: | :---: |
| **RPyXL Write $\rightarrow$ openpyxl Read** | 5,000 cells | String, Int, Float, Bool, Datetime | **0** | **100% PERFECT** |
| **openpyxl Write $\rightarrow$ RPyXL Read** | 5,000 cells | String, Int, Float, Bool, Datetime | **0** | **100% PERFECT** |

---

## What's New in Version 0.5.1 (vs 0.5.0)

| Capability / Subsystem | Prior Release (0.5.0) | Current Release (0.5.1) | Optimization Impact |
| :--- | :---: | :---: | :--- |
| **Styled Save Phase** (250k cells) | 9.20 s | **1.31 s** | 🚀 **7.0x faster**: Pre-cached cell coordinates, fast ASCII letter lookups, and direct interned style records |
| **Total Styled Write** (50k rows) | 13.87 s | **5.66 s** (or **4.24 s** with bulk API) | 🚀 **2.2x to 3.2x faster**: Outperforms openpyxl across full end-to-end styled document creation |
| **Native Freeze Panes** | Fallback to ZIP extraction & XML re-parse | Direct native OOXML serialization | ⚡ Native freeze panes support in `rpyxl-core`, `rpyxl-ooxml`, eliminating heavy Python ZIP re-serialization |
| **Smart XML Whitespace** | Emitted `xml:space="preserve"` unconditionally | Only emitted when strings contain leading/trailing spaces | 📦 **~4% smaller output**: Output files now smaller than openpyxl (1.57 MB vs 1.63 MB on 50k rows) |
| **Bulk Range Styling** | N/A | `ws.apply_style("A1:E50000", ...)` | ⚡ Snapshots and interns styles once before vectorized application across entire rectangular ranges |
| **Direct Style Keywords** | N/A | `ws.cell(..., font=..., fill=...)` | ⚡ Single-statement cell allocation with style assignment avoids multiple PyO3 property setter roundtrips |

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

# 3. Full Automated Test Suite (507 tests: 434 unit + 73 compatibility)
pytest tests/unit/ tests/compatibility/

# 4. Comprehensive Benchmark Suite
python benchmarks/live_benchmark.py
python benchmarks/benchmark_people_write.py
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
