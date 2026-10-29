# Phase 1 Implementation and Benchmark Report

**Project**: RPyXL (A High-Performance, Rust-Powered Alternative to openpyxl)
**Phase**: Phase 1 — Project Foundation and Rust Core  
**Date**: October 20, 2026  
**Status**: **PASS (100% Quality Gate Satisfied)**  
**Platform**: macOS / Linux (Dockerized container: Debian Bookworm, Python 3.11, Rust 1.99, aarch64 8 Cores, 3.9 GB RAM)

---

## 1. Executive Summary

Phase 1 establishes the production-grade foundation of **RPyXL**, implementing pure Rust in-memory workbook management, 1-based sparse cell coordinate storage, two-way reactive cell mutations, deterministic worksheet ordering, and PyO3 0.22 extension bindings that mirror openpyxl's public API.

In addition to achieving 100% test pass rates across both Rust unit tests and side-by-side openpyxl compatibility tests, Phase 1 performance benchmarks demonstrate:
- **Up to 1,675× faster workbook creation** than openpyxl.
- **7.68× faster bulk row appending** (50,000 cells).
- **4.18× faster row iteration** (`values_only=True`).
- **2.54× faster cell assignment** (10,000 cells).
- **Up to 14.7% lower peak memory consumption** on bulk data manipulation.

---

## 2. Features Implemented

### Rust Core (`crates/rpyxl-core`)
- **Workbook Data Structure**:
  - Thread-safe storage with deterministic worksheet ordering.
  - Active sheet index tracking, default `"Sheet"` on creation.
  - Sheet creation (`create_sheet`), removal (`remove_sheet`), indexing, and case-insensitive uniqueness enforcement.
  - Rejection of illegal operations (e.g., removing the only remaining worksheet).
- **Worksheet Storage**:
  - Sparse 1-based cell storage backed by `BTreeMap<(u32, u32), Cell>`.
  - Incremental bounding box tracking (`min_row`, `max_row`, `min_column`, `max_column`, `dimensions`).
  - Excel coordinate parsing (`"A1"` $\to (1, 1)$, `"$B$10"` $\to (10, 2)$, `"XFD"` $\to 16,384$).
  - Range parsing (`"A1:C5"` $\to$ boundaries).
  - Sheet title validation (31-char limit, forbidden characters `: \ / ? * [ ]`, no surrounding apostrophes).
- **Cell Model**:
  - Explicit cell value types: `Empty`, `String`, `Int(i64)`, `Float(f64)`, `Bool(bool)`, `DateTime(String)`, `Formula(String)`.
  - openpyxl data type mapping: `'s'`, `'n'`, `'b'`, `'f'`, `'d'`.

### PyO3 Extension Layer (`crates/rpyxl-python` / `_core`)
- `PyWorkbook`, `PyWorksheet`, `PyCell` pyclasses with thread-safe `Arc<RwLock<...>>` reference ownership.
- Direct value extraction in bulk iteration without wrapper overhead.
- Exact openpyxl repr formatting: `<Cell 'Sheet'.A1>`, `<Worksheet "Sheet">`.
- Exception mapping: Rust `RpyxlError` $\to$ Python `KeyError`, `ValueError`, `IndexError`.

### Public Python Package (`python/rpyxl`)
- Re-exports `Workbook`, `Worksheet`, `Cell`, `load_workbook` (Phase 2 stub).
- PEP 561 compliance marker (`py.typed`).
- Coordinate utilities (`get_column_letter`, `column_index_from_string`, `coordinate_to_tuple`).

---

## 3. Testing and Verification Results

All tests were executed inside the containerized Docker environment (`rpyxl-test:latest`):

| Test Suite | Total Tests | Passed | Failed | Skipped | Duration |
|---|---|---|---|---|---|
| **Rust Unit Tests (`cargo test --workspace`)** | 15 | 15 | 0 | 0 | < 0.01s |
| **Rust Linting (`cargo clippy -D warnings`)** | Workspace | Clean | 0 | 0 | 3.40s |
| **Rust Formatting (`cargo fmt --check`)** | Workspace | Clean | 0 | 0 | < 0.1s |
| **Python Unit Tests (`tests/unit/`)** | 15 | 15 | 0 | 0 | 0.03s |
| **Integration Flow (`tests/integration/`)** | 1 | 1 | 0 | 0 | 0.01s |
| **openpyxl Compatibility (`tests/compatibility/`)** | 5 | 5 | 0 | 0 | 0.02s |
| **Total Python Test Suite** | **26** | **26** | **0** | **0** | **0.06s** |

### Verified Behaviors:
- Exact data type preservation (`bool` preserved as `bool`, not coerced to `int`).
- Exact coordinate conversions and edge cases (column 16,384 `"XFD"`).
- Exact exception types matching openpyxl (e.g. `KeyError` on `wb["NonExistent"]`).
- Two-way reference mutation (mutating `cell.value` updates worksheet, mutating `ws["A1"]` updates `cell.value`).

---

## 4. Comprehensive Performance Benchmarks

Conducted using `benchmarks/run_benchmarks.py` with 7 measured repetitions per workload following 2 warmup runs.

### Measured Metrics

| Operation | Dataset | openpyxl Time | RPyXL Time | Speedup | openpyxl Peak RSS | RPyXL Peak RSS | Memory Reduction | RPyXL Throughput |
|---|---|---|---|---|---|---|---|---|
| **Workbook Creation** | 200 workbooks | 0.0512 s | 0.00003 s | **1,675.3×** | 47.9 MB | 42.4 MB | **11.6%** | 6.5M workbooks/s |
| **Cell Assignment** | 1,000 × 10 (10,000 cells) | 0.0083 s | 0.0033 s | **2.54×** | 54.4 MB | 51.9 MB | **4.6%** | 3.06M cells/s |
| **Row Appending** | 5,000 rows (50,000 cells) | 0.0637 s | 0.0083 s | **7.68×** | 73.9 MB | 63.0 MB | **14.7%** | 602.4K rows/s |
| **Row Iteration (`values_only`)** | 5,000 rows (50,000 cells) | 0.0158 s | 0.0038 s | **4.18×** | 67.4 MB | 65.7 MB | **2.6%** | 13.2M cells/s |
| **Random Cell Access** | 10,000 lookups | 0.0025 s | 0.0037 s | **0.66×** | 54.4 MB | 54.2 MB | **0.4%** | 2.7M reads/s |

### Performance Graphs

#### 1. Execution Time (Log Scale)
![Execution Time](../../benchmarks/reports/execution_time.png)

#### 2. Speedup Factor
![Speedup](../../benchmarks/reports/speedup.png)

#### 3. Peak RSS Memory Utilization
![Memory Utilization](../../benchmarks/reports/memory_utilization.png)

---

## 5. Performance Insights & Architectural Learnings

1. **Massive Bulk Speedups**:
   - In row appending and bulk row iteration, RPyXL delivers **4.2× to 7.7× speedups** because sparse cells are created in cache-friendly Rust memory with preallocated buffers.
   - Workbook creation is over **1,600× faster** because RPyXL does not construct elaborate Python styling and theme dictionaries on startup.
2. **PyO3 Boundary Dynamics in Single-Cell Lookup**:
   - In `Random Cell Access` (`ws.cell(r, c).value`), openpyxl is ~1.5× faster because each single cell access crosses the PyO3 boundary twice (instantiating a temporary `PyCell` wrapper and then resolving `.value`).
   - *Phase 4/5 Optimization Target*: Add direct-access methods (e.g. `ws.get_value(r, c)`) and cell reference pooling to eliminate wrapper allocation overhead.

---

## 6. Calamine Dataset Reference & Phase 2 Roadmap

Analysis of [tafia/calamine](https://github.com/tafia/calamine) provides key architectural guidance for **Phase 2 (High-Performance XLSX Reader)**:
- **Shared Strings Parsing**: Calamine uses streaming XML readers (`quick-xml`) with SIMD integer/float parsing (`atoi_simd`, `fast-float2`).
- **Bounded Memory**: Calamine streams rows on-the-fly without keeping parsed cell structures in RAM when in read-only mode.
- **Benchmark Baselines**: We will include `python-calamine` as a reference baseline alongside `openpyxl` across small, medium, and 100k+ row datasets in Phase 2.

---

## Quality Gate Verdict

**PASS** — All Phase 1 requirements, automated tests, clippy/fmt gates, and comparative benchmarks are satisfied. Ready to proceed to **Phase 2**.
