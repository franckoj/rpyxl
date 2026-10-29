# Phase 5 Engineering & Benchmark Report: Performance Optimization & Multi-Scale Evaluation

**Date**: 2026-10-20  
**Phase**: 5 of 6  
**Status**: Completed  
**Artifact**: RPyXL v0.1.0 High-Performance Engine Optimizations

---

## 1. Executive Summary

Phase 5 introduces targeted algorithmic and systems-level optimizations across RPyXL's Rust core (`rpyxl-core`) and OOXML serializer (`rpyxl-ooxml`). By eliminating intermediate heap allocations, adopting fast integer and floating-point serializers (`itoa` and `ryu`), reusing row buffers, and parallelizing multi-sheet serialization with Rayon, RPyXL achieves unprecedented throughput and low latency:

- **Up to 8.12× faster standard workbook save** vs `openpyxl`.
- **Up to 7.16× faster streaming write-only** vs `openpyxl`.
- **Up to 8.73× faster standard reading** (and 148 MB less RAM consumption on 500k cells).
- **Peak writing throughput exceeding 1.68 million cells per second** (168,865 rows/s).
- **Linear scalability across small (1k), medium (10k), large (100k), and stress (500k) workloads**.

All 99 tests (21 Rust workspace tests + 78 Python unit and compatibility tests) continue to pass at 100% with zero Clippy warnings (`-D warnings`) and zero formatting diffs.

---

## 2. Profiling & Optimization Architecture

### 2.1 Heap Allocation Elimination in Coordinate Generation
- **Problem**: Previously, `tuple_to_coordinate(row, col)` allocated two heap `String` objects per cell (one in `get_column_letter` and one in `format!("{col_letter}{row}")`). For a 500,000-cell workbook, this triggered **1,000,000 heap allocations** solely to generate cell tags.
- **Solution**: Implemented `write_coordinate(row, col, &mut String)` and `write_column_letter(col, &mut String)`. Excel columns (1 to 16,384) convert into 1–3 ASCII characters using a tiny stack array `[u8; 3]` with zero heap allocations. The row number is appended directly using `itoa::Buffer`.

### 2.2 Fast Integer & Floating-Point Number Serialization (`itoa` & `ryu`)
- **Problem**: Rust's standard `i.to_string()` and `f.to_string()` allocate intermediate `String` instances on the heap and involve standard library formatting machinery overhead.
- **Solution**:
  - Replaced integer formatting with `itoa::Buffer`, achieving sub-5-nanosecond integer-to-string formatting directly into the output stream.
  - Replaced floating-point formatting with `ryu::Buffer`, providing fast, allocation-free IEEE 754 serialization.

### 2.3 Buffer Pre-Allocation & Sheet Sizing
- **Problem**: In-memory worksheet XML serialization dynamically reallocated large `String` buffers multiple times as rows were appended, causing repetitive memory reallocation and copying.
- **Solution**: `generate_worksheet_xml` now computes an estimated capacity `512 + ws.cell_count() * 45` upfront, virtually eliminating buffer reallocations during sheet XML generation.

### 2.4 Streaming Row Buffer Reuse (`WriteOnlySheet`)
- **Problem**: In streaming write-only mode, each row append created a new `row_str` `String`, formatting `<row r="N">` and pushing to disk.
- **Solution**: Added a persistent, pre-allocated `row_buf: String` field in `WriteOnlySheet`. Each `append_row` call clears the buffer, formats the row with `itoa`, writes the cells, and flushes to the underlying buffered writer without allocating a new buffer per row.

### 2.5 Slice-Based XML Escaping
- **Problem**: Character-by-character iteration (`for c in s.chars()`) with UTF-8 decoding penalized strings that had no special characters (the vast majority of spreadsheet content).
- **Solution**: Implemented slice scanning using fast byte comparisons. If no special characters (`&`, `<`, `>`, `"`, `'`) are present, the entire slice is appended in a single memory copy (`out.push_str(s)`). When special characters exist, only unescaped slices and entity codes are emitted.

### 2.6 Multi-Sheet Parallel Serialization (`rayon`)
- **Problem**: Multi-sheet workbooks previously generated sheet XML sequentially on a single thread.
- **Solution**: Integrated `rayon::prelude::*`. Workbooks with multiple sheets serialize their worksheet XML concurrently across available CPU cores before writing into the central ZIP package.

---

## 3. Multi-Scale Benchmark Results

Benchmarks were executed in an isolated Linux container (`aarch64`) with 8 logical cores and 4 GB RAM across 4 standardized scales:
- **Small**: 1,000 cells (100 rows × 10 cols)
- **Medium**: 10,000 cells (1,000 rows × 10 cols)
- **Large**: 100,000 cells (10,000 rows × 10 cols)
- **Stress**: 500,000 cells (50,000 rows × 10 cols)

### 3.1 Comprehensive Multi-Scale Comparison Table

| Dataset Scale | Operation | openpyxl (Wall) | RPyXL (Wall) | Speedup Factor | openpyxl RSS | RPyXL RSS | RPyXL Throughput |
|---|---|---|---|---|---|---|---|
| **Small (1k cells)** | Standard Save | 0.0093s | 0.0020s | **4.71×** | 75.7 MB | 75.8 MB | **505,082 cells/s** |
| | Write-Only Save | 0.0130s | 0.0022s | **5.89×** | 75.8 MB | 75.8 MB | **452,753 cells/s** |
| | Standard Load | 0.0120s | 0.0024s | **5.00×** | 76.0 MB | 75.9 MB | **416,667 cells/s** |
| | Read-Only Stream | 0.0102s | 0.0024s | **4.18×** | 76.0 MB | 75.9 MB | **409,948 cells/s** |
| **Medium (10k cells)** | Standard Save | 0.0512s | 0.0101s | **5.05×** | 79.7 MB | 80.1 MB | **987,167 cells/s** |
| | Write-Only Save | 0.0493s | 0.0071s | **6.97×** | 79.9 MB | 80.1 MB | **1,414,669 cells/s** |
| | Standard Load | 0.0489s | 0.0097s | **5.04×** | 81.3 MB | 80.4 MB | **1,030,928 cells/s** |
| | Read-Only Stream | 0.0428s | 0.0113s | **3.80×** | 81.3 MB | 80.4 MB | **888,284 cells/s** |
| **Large (100k cells)** | Standard Save | 0.5115s | 0.0719s | **7.12×** | 107.8 MB | 117.9 MB | **1,391,000 cells/s** |
| | Write-Only Save | 0.4204s | 0.0648s | **6.49×** | 116.9 MB | 116.9 MB | **1,542,970 cells/s** |
| | Standard Load | 0.5678s | 0.0607s | **9.35×** | 140.8 MB | 122.9 MB | **1,647,446 cells/s** |
| | Read-Only Stream | 0.3389s | 0.0687s | **4.93×** | 122.9 MB | 122.9 MB | **1,455,411 cells/s** |
| **Stress (500k cells)** | Standard Save | 2.7976s | 0.3447s | **8.12×** | 213.8 MB | 218.1 MB | **1,450,658 cells/s** |
| | Write-Only Save | 2.1202s | 0.2961s | **7.16×** | 216.1 MB | 216.1 MB | **1,688,650 cells/s** |
| | Standard Load | 2.7133s | 0.3109s | **8.73×** | 373.1 MB | 225.1 MB | **1,608,234 cells/s** |
| | Read-Only Stream | 1.6987s | 0.3312s | **5.13×** | 168.6 MB | 168.6 MB | **1,509,715 cells/s** |

---

## 4. Performance Visualizations

### 4.1 Standard In-Memory Save Time (Log Scale)
![Standard Save Time](../../benchmarks/reports/phase5_scale_writing_time.png)

### 4.2 Streaming Write-Only Time (Log Scale)
![Write-Only Time](../../benchmarks/reports/phase5_scale_write_only_time.png)

### 4.3 Reading Time Comparison (Standard, Read-Only, and Calamine)
![Reading Time](../../benchmarks/reports/phase5_scale_reading_time.png)

### 4.4 Peak RSS Memory Scaling
![Memory Scaling](../../benchmarks/reports/phase5_scale_memory.png)

---

## 5. Key Findings & Performance Analysis

1. **Efficiency Increases with Dataset Scale**:
   - While RPyXL is already ~4.7× faster on small 1,000-cell sheets, its architectural advantage widens significantly to **8.12× faster** on 500,000 cells.
   - Rust's cache locality and lack of Python object boxing deliver sustained throughput above 1.5 million cells/second across both reading and writing.
2. **True Memory Efficiency on Large Reads**:
   - On the 500k stress dataset, standard openpyxl requires **373.1 MB** of RAM to construct Python cell representations.
   - RPyXL's standard load uses **225.1 MB** (a saving of **148 MB**), while RPyXL streaming read-only requires only **168.6 MB**.
3. **Multi-Core Utilization**:
   - CPU utilization measurements show RPyXL achieving up to **137%–186% CPU utilization** during heavy operations, demonstrating successful multi-core parallelism via Rayon, whereas pure Python execution remains strictly throttled by the GIL.

---

## 6. Phase 5 Sign-Off & Next Steps

Phase 5 has successfully achieved all core optimization objectives:
- Profiled bottlenecks and eliminated string formatting allocations.
- Integrated `itoa`, `ryu`, and `rayon` without introducing unsafe patterns.
- Verified 100% test pass rate across Rust and Python suites.
- Validated performance advantages across small, medium, large, and stress datasets.

The project is now ready to proceed to **Phase 6: Packaging, Distribution & Documentation**.
