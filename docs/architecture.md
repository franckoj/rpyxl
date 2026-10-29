# RPyXL Architecture

## Overview

RPyXL is a high-performance alternative to openpyxl designed for Python applications needing speed and bounded memory consumption when working with Excel workbooks.

The library is organized in a three-tier architecture:

```
┌────────────────────────────────────────────────────────┐
│  Python Public API Layer (python/rpyxl)                 │
│  - openpyxl-compatible signatures & docstrings         │
│  - Type definitions (py.typed)                         │
└───────────────────────────▲────────────────────────────┘
                            │ PyO3 ABI Boundary
┌───────────────────────────▼────────────────────────────┐
│  PyO3 Extension Module (crates/rpyxl-python / _core)    │
│  - Python-to-Rust type conversion                      │
│  - Reference management via Arc<RwLock<...>>           │
│  - Error translation to Python exception hierarchy     │
└───────────────────────────▲────────────────────────────┘
                            │ Pure Rust calls
┌───────────────────────────▼────────────────────────────┐
│  Core Engine (crates/rpyxl-core)                        │
│  - In-memory Workbook, Worksheet, Cell models          │
│  - Sparse cell indexing (BTreeMap<(row, col), Cell>)   │
│  - Coordinate conversions (1-based, letters <-> idx)   │
└────────────────────────────────────────────────────────┘
```

## Core Rust Design Principles

1. **Rust Ownership**: All workbook structures, worksheet collections, and cell data are owned in Rust. Python wrappers hold thread-safe `Arc<RwLock<Workbook>>` references.
2. **Sparse Storage**: Cells are indexed using `BTreeMap<(u32, u32), Cell>`, ensuring minimal memory overhead when worksheets are sparse. Iterating rows or ranges leverages natural key order.
3. **Zero Dangling References**: Mutating a cell via `cell.value = ...` updates the underlying workbook immediately; mutating a worksheet via `ws["A1"] = ...` is immediately reflected when reading `cell.value`.
4. **Boundary Minimization**: In bulk iteration (`iter_rows(values_only=True)`), primitive values are extracted directly without creating intermediate `Cell` wrappers.
