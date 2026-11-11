#!/usr/bin/env python3
"""
Benchmark: Single-Row Append vs Native Batch Append (append_rows)
Comparing:
1. openpyxl.append() (single-row loop)
2. rpyxl.append() (single-row loop)
3. rpyxl.append_rows() (batch engine)
"""

import os
import sys
import time
import gc
import psutil
from tabulate import tabulate

import openpyxl
import rpyxl

print("====================================================================")
print(f" BENCHMARK: SINGLE-ROW APPEND vs BATCH APPEND_ROWS")
print(f" Python: {sys.version.split()[0]} | openpyxl: {openpyxl.__version__} | rpyxl: {rpyxl.__version__}")
print(f" CPU Cores: {psutil.cpu_count(logical=True)}")
print("====================================================================\n")

OUTPUT_DIR = "/tmp/batch_bench_data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

SCALES = [
    ("10,000 rows", 10000, 10),      # 100,000 cells
    ("50,000 rows", 50000, 10),      # 500,000 cells
    ("100,000 rows", 100000, 10),    # 1,000,000 cells
]

def make_rows(n_rows, n_cols):
    return [[f"R{r}_C{c}" if c % 3 == 0 else (r * c if c % 2 == 0 else r * 1.5) for c in range(1, n_cols + 1)] for r in range(1, n_rows + 1)]

results = []

for label, n_rows, n_cols in SCALES:
    total_cells = n_rows * n_cols
    print(f"--- Running {label} ({total_cells:,} cells) ---")
    rows = make_rows(n_rows, n_cols)

    # 1. openpyxl row-by-row
    gc.collect()
    t0 = time.perf_counter()
    wb_ox = openpyxl.Workbook()
    ws_ox = wb_ox.active
    for r in rows:
        ws_ox.append(r)
    t_append_ox = time.perf_counter() - t0
    wb_ox.close()

    # 2. rpyxl row-by-row
    gc.collect()
    t0 = time.perf_counter()
    wb_px = rpyxl.Workbook()
    ws_px = wb_px.active
    for r in rows:
        ws_px.append(r)
    t_append_px = time.perf_counter() - t0
    wb_px.close()

    # 3. rpyxl batch append_rows
    gc.collect()
    t0 = time.perf_counter()
    wb_batch = rpyxl.Workbook()
    ws_batch = wb_batch.active
    ws_batch.append_rows(rows)
    t_batch_px = time.perf_counter() - t0
    wb_batch.close()

    speedup_vs_ox = t_append_ox / t_batch_px if t_batch_px > 0 else 0
    speedup_vs_single = t_append_px / t_batch_px if t_batch_px > 0 else 0
    throughput = total_cells / t_batch_px if t_batch_px > 0 else 0

    results.append([
        f"{label} ({total_cells:,} cells)",
        f"{t_append_ox:.4f}s",
        f"{t_append_px:.4f}s",
        f"{t_batch_px:.4f}s",
        f"{speedup_vs_single:.2f}x faster",
        f"{speedup_vs_ox:.2f}x faster",
        f"{throughput:,.0f} cells/s"
    ])

print("\n" + "="*80)
print(" INGESTION SPEED RESULTS: SINGLE-ROW vs BATCH APPEND")
print("="*80)
print(tabulate(
    results,
    headers=[
        "Workload Scale",
        "openpyxl (Single)",
        "rpyxl (Single)",
        "rpyxl (Batch append_rows)",
        "Batch vs Single rpyxl",
        "Batch vs openpyxl",
        "Batch Ingestion Throughput"
    ],
    tablefmt="github"
))

# Test Write-Only Streaming with batch append
print("\n--- Testing Streaming Write-Only with Batch Append (1,000,000 cells) ---")
wo_rows = make_rows(100000, 10)
out_wo_single = os.path.join(OUTPUT_DIR, "wo_single.xlsx")
out_wo_batch = os.path.join(OUTPUT_DIR, "wo_batch.xlsx")

# Single append loop
t0 = time.perf_counter()
wb = rpyxl.Workbook(write_only=True)
ws = wb.create_sheet()
for r in wo_rows:
    ws.append(r)
wb.save(out_wo_single)
wb.close()
t_wo_single = time.perf_counter() - t0

# Batch append
t0 = time.perf_counter()
wb = rpyxl.Workbook(write_only=True)
ws = wb.create_sheet()
ws.append_rows(wo_rows)
wb.save(out_wo_batch)
wb.close()
t_wo_batch = time.perf_counter() - t0

print(tabulate([
    ["Write-Only Single Append Loop", f"{t_wo_single:.4f}s", f"{1000000 / t_wo_single:,.0f} cells/s", "1.00x"],
    ["Write-Only Batch append_rows()", f"{t_wo_batch:.4f}s", f"{1000000 / t_wo_batch:,.0f} cells/s", f"{t_wo_single / t_wo_batch:.2f}x faster"]
], headers=["Streaming Write Mode (1M cells)", "Time", "Throughput", "Gain"], tablefmt="github"))

