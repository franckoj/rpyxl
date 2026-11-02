# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
