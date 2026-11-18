# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.5.1] - 2026-11-18

This release introduces comprehensive performance optimizations and native serialization enhancements for styled writes, speeding up styled saving by **7.0x** (1.31s vs 9.20s baseline), reducing output file sizes by ~4% (smaller than openpyxl), and adding fast bulk styling APIs.

### Added
- **Native Freeze Panes OOXML Serialization**: Implemented native `freeze_panes` in `rpyxl-core`, `rpyxl-ooxml`, and `rpyxl-python`, completely eliminating the legacy Python ZIP-extraction and `ElementTree` re-serialization fallback.
- **Bulk Range Styling API (`Worksheet.apply_style`)**: Added `ws.apply_style("A1:E50000", ...)` to snapshot and intern styles once before mass application across cell ranges in a single vectorized pass (~0.52s for 250k cells).
- **Direct Style Keywords in `Worksheet.cell()`**: Added optional `font`, `fill`, `border`, `alignment`, and `protection` kwargs to `ws.cell(...)` for single-statement cell creation and styling without multiple property setter round-trips.
- **Random People Benchmark**: Added `benchmarks/benchmark_people_write.py` comparing rpyxl against openpyxl across plain write, streaming write_only, and styled write workloads.

### Changed & Optimized
- **7.0x Faster Styled Writes Save Phase**: Reduced styled save time from 9.20s down to 1.31s on 250,000 cells (4.52x faster than openpyxl). Total styled write time dropped from 13.87s to 5.41s (2.15x faster than openpyxl; 4.24s with bulk keyword styling).
- **Smart Whitespace XML Serialization**: Emits `xml:space="preserve"` only when cell strings actually contain leading or trailing whitespace, reducing raw XML by 4.2 MB across 200k string cells and producing archives smaller than openpyxl (1.57 MB vs 1.63 MB).
- **Zero-Alloc Coordinate Byte Formatting**: Added `_LETTERS` lookup tuple on `Cell` view for fast ASCII column coordinate resolution (`col <= 26`), caching `_row` and `_column` attributes and avoiding redundant PyO3 FFI string allocations.
- **Style Compilation Fast-Path**: Replaced 250,000 `coordinate_parts()` string parses in `compile_style_records` by directly utilizing pre-cached integer coordinates on cell metadata.

---

## [0.5.0] - 2026-11-18

This release introduces next-generation performance architecture spanning Tier 1 zero-allocation I/O & XML acceleration, Tier 2 row-chunked memory & streaming iteration optimizations, and a direct scalar cell write fast-path, delivering **6x to 18x speedups** over openpyxl 3.1.5 and up to **6.17x speedup** over rpyxl 0.4.0.

### Added
- **Zero-Allocation Coordinate Byte Parsing**: Added `coordinate_bytes_to_tuple(&[u8]) -> Option<(u32, u32)>` in `rpyxl-core/src/coordinate.rs`. Parses raw ASCII coordinates (e.g. `b"AA1048576"`) directly from XML token byte slices into `(row, col)` tuples without intermediate heap string allocations, with robust fallbacks for entity-encoded attributes (`&#x41;&#49;`).
- **Zero-Allocation XML Attribute Parsing**: Replaced heap-allocated strings for cell types, styles, and coordinates in `rpyxl-ooxml/src/reader.rs` with copy-free `CellTypeTag` enums, integer parsers (`parse_style_index`, `parse_u32_ascii`), and raw byte-slice matching across `StreamingRowReader`, `parse_worksheet_into`, and `read_single_cell`.
- **Parallel Multi-Sheet Reader**: Integrated Rayon parallel execution in `to_workbook()`, allowing multi-sheet workbooks to parse worksheet XML payloads concurrently across CPU threads.
- **Direct Scalar Cell Write Fast-Path**: Optimized `Worksheet.cell(row, col, value=...)` for unmerged sheets to pass values directly into Rust native cell constructors in a single FFI crossing, bypassing redundant Python `Cell` wrapping and subsequent property setter dispatches (6.17x faster cell writes than 0.4.0, up to 655,000 cells/s).
- **Automated Version Comparison Benchmark**: Added `benchmarks/bench_040_vs_latest.py` for automated head-to-head benchmarking between release versions across all writing, reading, and in-memory workflows.

### Changed & Optimized
- **Atomic File Persistence (`NamedTempFile::persist`)**: Replaced buffered temporary file copy loops with atomic rename in `rpyxl-ooxml/src/writer.rs`, eliminating the double disk I/O penalty during workbook saves.
- **Zero-Allocation Streaming Row Iterator**: Replaced per-row temporary `HashMap` allocations in `PyRowIterator::__next__` with a linear peekable iterator over the sorted cell stream, converting $\mathcal{O}(N)$ hashing into allocation-free $\mathcal{O}(1)$ peek/advance operations.
- **Sheet Index Resolution Caching**: Added `cached_sheet_idx` tracking to `PyWorksheetRowIterator` and `PyWorksheetColIterator` in `rpyxl-python/src/bindings.rs`, bypassing $\mathcal{O}(S)$ sheet title iteration on every row iteration.
- **Monotonic Direct BTreeMap Cell Insertions**: Replaced per-cell `cells.entry().or_insert_with()` tree traversals in `Worksheet::append` and `Worksheet::append_rows` (`rpyxl-core/src/worksheet.rs`) with direct in-place insertions, computing and updating cached sheet boundaries in one batch.
- **CPython Pointer Downcasting & PyO3 Interning**: Replaced Python MRO reflection in `python_value_to_cell_value` with direct CPython pointer downcasting for primitive types. Module and method names (`datetime`, `isoformat`, `utcoffset`) are interned with `pyo3::intern!`, completely eliminating runtime string allocations in hot loops.
- **Worksheet Iterator Local Binding**: Bound `cell_cls` and `wrap` locally in `python/rpyxl/worksheet/__init__.py` to eliminate Python global lookups inside inner row/column comprehensions.

---

## [0.4.0] - 2026-11-17

This release delivers major performance optimizations across cell access, streaming iteration, styled worksheet serialization, and configurable compression, alongside enhanced developer tooling and comprehensive benchmark suites.

### Added
- **Native Streaming Row & Column Iterators**: Implemented `PyWorksheetRowIterator` and `PyWorksheetColIterator` in Rust (`rpyxl-python/src/bindings.rs`). Iterating over rows via `ws.iter_rows(values_only=False)` or columns via `ws.iter_cols()` now makes a single native call that yields row/column tuples lazily without materializing full grids or issuing redundant cross-FFI calls per row (1.49x faster cell row traversal, 817k cells/s).
- **Versioned `PyCell` Value Caching**: Added workbook version tracking (`Workbook.version`) and thread-safe cell value caching in `PyCell.cached_value`. Repeated cell value reads bypass sheet RwLock acquisition and map lookup when unchanged, accelerating cell reads by **1.67x** (up to 592k cells/s).
- **Configurable Deflate Compression**: Added `compression_level` property to `Workbook` and `WriteOnlyWorkbook`, and support in `wb.save(filename, compression_level=...)`. Defaults to standard level 6, shrinking archive sizes by **31% to 34%** while preserving high-speed saving.
- **Full Test Suite & Benchmark Harness**: Added `benchmarks/benchmark.py` and `benchmarks/full_test_openpyxl_vs_rpyxl.py` covering 7 testing sections (`compat`, `interop`, `perf`, `scale`, `threads`, `concurrency`, `leak`) across 20 granular performance scenarios.
- **Styled Sweep Benchmark**: Added `benchmarks/bench_styled_sweep.py` for profiling build, save, memory, and archive size scaling up to 500k rows.

### Changed & Optimized
- **Inline XML Style Serialization Caching**: In `rpyxl-ooxml/src/writer.rs`, implemented consecutive cell style caching (`last_rich_ptr`, `last_eff_fmt`, `last_xf_idx`, `last_style`) in `generate_worksheet_xml`. Sequential cells sharing styles avoid map lookups, speeding up styled worksheet XML emission.
- **Fast-Path General Formats**: Non-date cells without explicit number formats immediately resolve to `format_id = 0`, bypassing number format resolution maps.
- **Integer Coordinate Bounds Check**: Replaced regex and string formatting coordinate checks in `ws.cell()` with direct integer boundary checks (`1 <= row <= 1048576 and 1 <= column <= 16384`), boosting individual cell writes by **1.43x** (up to 325k cells/s).
- **Merged Cell Range Fast-Path**: Worksheets without merged cells (`_may_have_merges == False`) bypass merge range containment searches entirely when wrapping cells. Pre-parsed `bounds` tuples are cached directly on `CellRange` instances.
- **Zero-Allocation Sheet Title Access**: Replaced per-cell `Arc<RwLock<String>>` cloning with direct reference borrowing in Python cell bindings.
- **Code Quality & CI Tooling**: Integrated Ruff linting and formatting across `python/`, `tests/`, and `benchmarks/`, and codified pre-push feature branch quality gates.

---

## [0.3.0] - 2026-11-11

This release implements high-leverage ingestion, compression, and style serialization speedups identified through live containerized benchmarking against openpyxl 3.1.5.

### Added
- **Native Batch Append**: Added `ws.append_rows(rows)` and alias `ws.extend(rows)` to `Worksheet` (core, python bindings, wrapper) and `WriteOnlyWorksheet`. Implements lock acquisition once per batch, zero-alloc row buffering, and single-pass bounding box expansion. Ingestion speed surged to **5,159,243 cells/sec** (6.15x faster than openpyxl; 1,000,000 cells ingested in 0.1938s).
- **Batch Dict Support**: Added full support for dictionary rows (`{"A": 1, "C": 2}`, `{1: "A", 2: "B"}`) to both single `append()` and batch `append_rows()` / `extend()` in both standard and streaming write-only modes.
- **Dedicated Live Benchmark Suite**: Added comprehensive reproducible benchmarks in `benchmarks/`:
  - `benchmarks/live_benchmark.py`: Export/write throughput, standard/streaming read, TTFR, transactional mutation, and 4-way cross-library roundtrip verification.
  - `benchmarks/bench_batch_append.py`: Single-row append vs batch append across 10k, 50k, and 100k rows (up to 1M cells).
  - `benchmarks/styles_and_memory_benchmark.py`: Styled write throughput, $\Delta$RSS memory growth, and archive size comparison across 25k, 100k, and 250k cells.
  - `benchmarks/advanced_live_benchmark.py`: 1M-cell Rayon multi-sheet parallelism, 2D range slicing (`ws["A1:J5000"]`), and full re-serialization cycles.

### Changed & Optimized
- **Deflate Compression Level 1 Tuning**: In `rpyxl-ooxml/src/writer.rs`, tuned deflate compression level to 1 (`CompressionMethod::Deflated`, `level=1`). Retains 95% of archive compression efficiency while accelerating streaming write throughput to **3,071,851 cells/sec** (**14.45x faster** than openpyxl; 500k cells saved in 0.1628s) and standard save to **1,615,983 cells/sec** (**9.99x faster** than openpyxl).
- **Zero-Allocation B-Tree Cursor for `iter_rows(values_only=True)`**: Replaced $O(N \log M)$ coordinate map searches with an $O(N)$ linear B-tree range cursor across `ws.cells()`. Standard worksheet read throughput reaches **945,586 cells/sec** (**6.85x faster** than openpyxl).
- **Identity-Cached Style Compilation**: In `rpyxl/_native_styles.py`, added per-save `identity_cache` mapping style object IDs to compiled native tuples, eliminating 1,250,000 redundant dataclass reflections on 250k cells and accelerating styled save times by **3.19x** (0.457s vs 1.457s).
- **Fast-Path ASCII Coordinate Parser**: In `rpyxl/_models.py`, added fast-path ASCII coordinate parser for common `"A1"`, `"B12"`, `"AA100"` cell coordinates, avoiding Python regex engine dispatch for high-frequency coordinate parsing.

---

## [0.2.0] - 2026-11-02

This release expands the openpyxl-style wrapper and improves selected native save paths since 0.1.1. Compatibility remains incomplete: this is not a general drop-in replacement, a lossless arbitrary-OOXML editor, or a production-readiness claim. The 0.1.0 entry below is retained as historical release text, not revalidated claims for 0.2.0.

### Added

- Expanded wrapper support for cell styles, worksheet metadata (merges, dimensions, freeze panes, filters, visibility, protection and page setup), structural edits, workbook properties, defined names, comments, hyperlinks, validations, conditional formatting, tables and basic bar/column charts. Support is feature- and mode-specific, not full openpyxl parity.
- Binary file-like loading and saving, including `BytesIO`, with caller-owned stream lifetime, input-position preservation and output replacement/truncation. Inputs are buffered; output streams must support seek/write/truncate.
- Native interned style persistence with pooled copy-on-write Python style views, lazy loaded-style decoding and sparse overrides. Source stylesheet provenance includes the distinction between explicit `s="0"` and an absent style attribute.
- Preservation-oriented saves for eligible loaded workbooks: retain a bounded compressed source snapshot, raw-copy unchanged ZIP members and regenerate native cells/styles in one output package. Rich or unproven edits retain the conservative metadata overlay; preservation of unknown parts remains best-effort.
- An atomic, all-sheet eligibility proof enabling native structural saves only for a narrowly proven plain-workbook subset. Unproven workbooks retain the conservative fallback.
- A separate [future backend evaluation plan](docs/future_backend_plan.md). No new backend implementation or runtime dependency is introduced by that plan in 0.2.0.

### Fixed

- Exact text handling for leading/trailing whitespace, XML newline normalization and carriage-return escaping.
- Date/datetime/time conversion, fractional seconds, millisecond serial rounding, 1900/1904 epoch reads and number-format persistence. Elapsed durations and some locale-specific/write-only formats remain limited.
- Malformed coordinate validation: require a complete ASCII cell reference, reject mixed/trailing content and misplaced absolute markers, and retain Excel row/column bounds checks.
- Range separator validation: reject extra separators, missing endpoints and malformed endpoint references while retaining absolute references and normalized reversed bounds. These fixes and regressions are incorporated from `main`.
- Save-time safety rejection of detected unsupported cell payloads, including shared/array-formula attributes, inline rich-text runs and cell metadata, rather than silently rewriting them incorrectly. Loading/reading remains available; broader shared-string and formula coverage is incomplete.

### Changed

- Reduced repeated style eligibility scans and snapshots; metadata output avoids full-tree copying and double serialization while restoring temporarily changed attributes safely.
- Structural editing still rejects formula translation and reference-bearing features where correct rewriting is not implemented. There is no formula calculation, general editable-chart support, threaded-comment support or new-image API.

### Validation and measured performance

- **Pre-`main`-merge evidence at `5883f7f`:** 112 Rust tests passed; 450 Python tests passed with one skip on each of Windows Python 3.11 and 3.14. The skip targets older native wheels lacking the new private API. These are not validation counts for the merged 0.2.0 release candidate; combined results must be reported separately.
- The earlier quick interoperability harness recorded 43/44 API probes, zero cell mismatches and 10/10 feature checks in each of four writer/reader directions. The remaining API difference is stricter rejection of the cell-reference-like table name `T1`. The outer harness command timed out after producing its reports; this is scoped evidence, not a clean full-release acceptance run.
- Targeted measurements compare `6899d7e` with `5883f7f` and openpyxl 3.1.5 in the same Windows/Python 3.14.3 session, using fresh processes, rotating order and three repetitions. Both cases use **20,000 rows × 20 columns**; figures below are medians. Memory is **sampled incremental RSS**, not absolute peak memory.

| Workload                                | Previous rpyxl | `5883f7f` rpyxl | openpyxl | rpyxl incremental RSS, previous → new |
| --------------------------------------- | -------------: | --------------: | -------: | ------------------------------------: |
| Styled write                            |        18.49 s |         13.63 s |  17.45 s |                     304.2 → 300.5 MiB |
| Insert/delete rows, including load/save |        16.36 s |          5.55 s |  11.75 s |                     722.7 → 152.4 MiB |

- Styled-write memory remains above openpyxl's 194.5 MiB in that run (approximately 300 versus 195 MiB). The Python style sidecar has not been eliminated. Output checks compared 400,000 cells per case per rpyxl wheel with zero mismatches, plus nine sampled style checks per styled output; this does not establish exhaustive style fidelity.
- An append control remains unresolved: a separate nine-pair alternating-order investigation found median total wall time approximately 11% higher at 10,000 × 20 and 13% higher at 50,000 × 20 for the new wheel. No specific source-level cause was established. Read iteration was roughly unchanged in the original control.
- The full 20-scenario large benchmark matrix has not been repeated after this increment. These results do not imply universal speed, CPU or memory superiority, constant-memory loaded editing, or the historical 0.1.0 speedups for this release. See the [remediation evidence and limitations](docs/remediation_plan.md).

## [0.1.1] - 2026-10-29

### Changed

- Renamed the Python distribution and import namespace to `rpyxl`, including the native extension `rpyxl._core`.
- Renamed Rust crates to `rpyxl-core`, `rpyxl-ooxml`, and `rpyxl-python`, with matching internal identifiers.
- Updated examples, tests, CI, and benchmark labels. Consumers must update their package requirements and imports; no legacy import alias is provided.

## [0.1.0] - 2026-10-20

Initial production-ready release of **RPyXL**, a high-performance, Rust-powered alternative to `openpyxl`.

### Added

#### 1. Core Architecture & In-Memory Engine (`rpyxl-core`)

- Pure Rust sparse cell store utilizing `BTreeMap<(row, col), Cell>` for compact memory locality.
- Thread-safe workbook, worksheet, and cell primitives wrapped with `Arc<RwLock<...>>` ensuring zero dangling pointers in Python.
- Fast 1-based coordinate conversions (`tuple_to_coordinate`, `coordinate_to_tuple`, `get_column_letter`, `column_index_from_string`, `range_boundaries`).
- Comprehensive error types (`RpyxlError`, `RpyxlResult`) mapping directly to Python exceptions.

#### 2. High-Performance XLSX Reader (`rpyxl-ooxml`)

- Integrated `calamine` parser for standard `load_workbook(filename)` and `data_only=True` evaluation.
- Streaming read-only reader (`load_workbook(read_only=True)`) delivering constant $O(1)$ memory consumption.
- Background worker thread with bounded crossbeam channel streaming rows lazily to Python generators.
- Shared strings table (`<sst>`) and styles parser (`<styleSheet>`) extracting numbers, dates, booleans, and formulas.

#### 3. High-Performance XLSX Writer (`rpyxl-ooxml`)

- Standard in-memory workbook serialization (`wb.save()`) generating valid ECMA-376 OOXML packages.
- Streaming write-only workbook (`Workbook(write_only=True)`) with zero-buffering disk XML flusher.
- Multi-sheet parallel XML serialization using `rayon`.

#### 4. openpyxl Compatibility Layer

- Drop-in import compatibility: `import rpyxl as openpyxl`.
- Full Workbook API: `wb.active`, `wb.sheetnames`, `wb.worksheets`, `wb.create_sheet()`, `wb.remove()`, `wb.copy_worksheet()`, `wb.index()`, `wb.save()`, `wb.close()`.
- Full Worksheet API: Subscript slicing (`ws["A1"]`, `ws["A1:C5"]`, `ws["A"]`, `ws["A:C"]`, `ws[1]`, `ws["1:5"]`, `ws[1:5]`), `ws.cell()`, dict append (`ws.append({"A": 1, "C": 2})`), `ws.iter_rows()`, `ws.iter_cols()`, `ws.values`, `ws.rows`, `ws.columns`.
- Full Cell API: `cell.value`, `cell.coordinate`, `cell.row`, `cell.column`, `cell.column_letter`, `cell.data_type`, `cell.is_date`, `cell.number_format`, `cell.comment`, `cell.hyperlink`, `cell.parent`.
- Foundational styles dataclasses (`rpyxl.styles.Font`, `PatternFill`, `Alignment`, `Border`, `Side`, `Color`).
- Complete exception hierarchy in `rpyxl.utils.exceptions` (`OpenpyxlException`, `InvalidFileException`, `SheetTitleException`, `CellCoordinatesException`, `NamedRangeException`).

#### 5. Performance Optimizations

- Zero-allocation coordinate formatting using small stack buffers (`[u8; 3]`).
- Fast integer serialization using `itoa::Buffer` (< 5 ns per integer).
- Fast IEEE 754 float serialization using `ryu::Buffer`.
- Pre-allocated worksheet XML buffers based on cell density.
- Slice-based byte scanning for XML entity escaping.
- Persistent streaming row buffer reuse in `WriteOnlySheet`.

#### 6. Production Hardening & Security Defenses

- ZIP bomb & decompression explosion defense: 512 MB per-entry limit, 1,024 MB total uncompressed limit, 100:1 compression ratio ceiling, and 10,000 max entries limit.
- XML External Entity (XXE) and recursive entity expansion (Billion Laughs) defense via strict `Event::DocType` rejection.
- Shared strings table memory cap (10,000,000 entries).
- Excel coordinate boundary limits enforcement (1 to 1,048,576 rows, 1 to 16,384 columns).
- Consistent `InvalidFileException` raising for malformed files and security violations.

#### 7. Packaging & Distribution

- Universal ABI3 wheels (`cp39-abi3`) supporting Python 3.9 through 3.13.
- Automated GitHub Actions CI workflow (`.github/workflows/ci.yml`) testing Linux and macOS.
- Automated multi-platform wheel build workflow (`.github/workflows/wheels.yml`) supporting Linux (x86_64, aarch64), macOS (Intel, Apple Silicon), and Windows (x64).
- Comprehensive documentation: Migration Guide, API Reference Manual, Security Threat Model, Compatibility Matrix, and 8 Phase Reports.

### Performance Benchmarks

- **Standard Save**: **4.7× to 8.12× faster** vs openpyxl.
- **Write-Only Streaming**: **5.9× to 7.16× faster** vs openpyxl.
- **Standard Load**: **5.0× to 8.73× faster** vs openpyxl.
- **Time to First Row (Read-Only)**: **552× faster** vs openpyxl.
- **Peak Throughput**: Exceeding **1,688,000 cells/second**.
- **Memory Savings**: **148 MB less RAM** on 500,000 cells.
