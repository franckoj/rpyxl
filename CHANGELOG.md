# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

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
