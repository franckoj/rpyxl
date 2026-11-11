#!/usr/bin/env python3
"""
Live Head-to-Head Benchmark: openpyxl vs rpyxl (v0.2.0)
Executed directly in the production environment.
"""

import os
import sys
import time
import datetime
import gc
import psutil
from tabulate import tabulate

import openpyxl
import rpyxl

print(f"================================================================")
print(f" LIVE BENCHMARK: openpyxl ({openpyxl.__version__}) vs rpyxl ({rpyxl.__version__})")
print(f" Python: {sys.version.split()[0]} | Platform: {sys.platform} | CPU cores: {psutil.cpu_count()}")
print(f"================================================================\n")

OUTPUT_DIR = "/tmp/live_bench_data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_rss_mb():
    gc.collect()
    return psutil.Process().memory_info().rss / (1024 * 1024)

# Multi-type test row generator
def generate_sample_rows(n_rows):
    rows = []
    base_dt = datetime.datetime(2026, 1, 1, 12, 0, 0)
    for i in range(1, n_rows + 1):
        row = [
            f"ID_{i:06d}",                              # Col 1: Formatted String
            i * 10,                                     # Col 2: Integer
            round(i * 1.25, 4),                         # Col 3: Float
            (i % 2 == 0),                               # Col 4: Boolean
            base_dt + datetime.timedelta(minutes=i),    # Col 5: Datetime
            f"Category_{i % 50}",                       # Col 6: Repeated String
            i * 100,                                    # Col 7: Integer
            round(i * 0.05, 2),                         # Col 8: Currency Float
            (i % 3 != 0),                               # Col 9: Boolean
            "  Padded & Text <special>  "               # Col 10: Special / Padded String
        ]
        rows.append(row)
    return rows

# ==============================================================================
# 1. WRITE BENCHMARK (STANDARD & WRITE-ONLY STREAMING)
# ==============================================================================
SCALES = [
    ("Small", 1000),      # 1,000 rows x 10 cols = 10,000 cells
    ("Medium", 10000),    # 10,000 rows x 10 cols = 100,000 cells
    ("Large", 50000),     # 50,000 rows x 10 cols = 500,000 cells
]

print("--- 1. BENCHMARKING EXPORT / WRITE PERFORMANCE ---")
write_results = []

for label, n_rows in SCALES:
    n_cells = n_rows * 10
    rows = generate_sample_rows(n_rows)
    
    # --- Standard DOM Write ---
    # openpyxl
    file_ox = os.path.join(OUTPUT_DIR, f"std_ox_{n_rows}.xlsx")
    rss_before = get_rss_mb()
    t0 = time.perf_counter()
    wb_ox = openpyxl.Workbook()
    ws_ox = wb_ox.active
    for r in rows:
        ws_ox.append(r)
    wb_ox.save(file_ox)
    wb_ox.close()
    t_std_ox = time.perf_counter() - t0
    rss_std_ox = get_rss_mb() - rss_before

    # rpyxl
    file_px = os.path.join(OUTPUT_DIR, f"std_px_{n_rows}.xlsx")
    rss_before = get_rss_mb()
    t0 = time.perf_counter()
    wb_px = rpyxl.Workbook()
    ws_px = wb_px.active
    for r in rows:
        ws_px.append(r)
    wb_px.save(file_px)
    wb_px.close()
    t_std_px = time.perf_counter() - t0
    rss_std_px = get_rss_mb() - rss_before

    std_speedup = t_std_ox / t_std_px if t_std_px > 0 else 0
    std_throughput = n_cells / t_std_px if t_std_px > 0 else 0

    # --- Streaming Write-Only ---
    # openpyxl
    file_wo_ox = os.path.join(OUTPUT_DIR, f"wo_ox_{n_rows}.xlsx")
    t0 = time.perf_counter()
    wb_wo_ox = openpyxl.Workbook(write_only=True)
    ws_wo_ox = wb_wo_ox.create_sheet()
    for r in rows:
        ws_wo_ox.append(r)
    wb_wo_ox.save(file_wo_ox)
    wb_wo_ox.close()
    t_wo_ox = time.perf_counter() - t0

    # rpyxl
    file_wo_px = os.path.join(OUTPUT_DIR, f"wo_px_{n_rows}.xlsx")
    t0 = time.perf_counter()
    wb_wo_px = rpyxl.Workbook(write_only=True)
    ws_wo_px = wb_wo_px.create_sheet()
    for r in rows:
        ws_wo_px.append(r)
    wb_wo_px.save(file_wo_px)
    wb_wo_px.close()
    t_wo_px = time.perf_counter() - t0

    wo_speedup = t_wo_ox / t_wo_px if t_wo_px > 0 else 0
    wo_throughput = n_cells / t_wo_px if t_wo_px > 0 else 0

    write_results.append([
        f"{label} ({n_cells:,} cells)",
        f"{t_std_ox:.4f}s", f"{t_std_px:.4f}s", f"{std_speedup:.2f}x", f"{std_throughput:,.0f} c/s",
        f"{t_wo_ox:.4f}s", f"{t_wo_px:.4f}s", f"{wo_speedup:.2f}x", f"{wo_throughput:,.0f} c/s"
    ])

print(tabulate(
    write_results,
    headers=["Scale", "Std openpyxl", "Std rpyxl", "Std Speedup", "Std Throughput",
             "Stream openpyxl", "Stream rpyxl", "Stream Speedup", "Stream Throughput"],
    tablefmt="github"
))
print()

# ==============================================================================
# 2. READ BENCHMARK (STANDARD & STREAMING READ-ONLY)
# ==============================================================================
print("--- 2. BENCHMARKING INGESTION / READ PERFORMANCE ---")
read_results = []

for label, n_rows in SCALES:
    n_cells = n_rows * 10
    target_file = os.path.join(OUTPUT_DIR, f"std_ox_{n_rows}.xlsx")

    # --- Standard Load ---
    # openpyxl
    t0 = time.perf_counter()
    wb_ox = openpyxl.load_workbook(target_file, data_only=True)
    rows_ox = list(wb_ox.active.iter_rows(values_only=True))
    wb_ox.close()
    t_read_std_ox = time.perf_counter() - t0

    # rpyxl
    t0 = time.perf_counter()
    wb_px = rpyxl.load_workbook(target_file, data_only=True)
    rows_px = list(wb_px.active.iter_rows(values_only=True))
    wb_px.close()
    t_read_std_px = time.perf_counter() - t0

    read_std_speedup = t_read_std_ox / t_read_std_px if t_read_std_px > 0 else 0
    read_std_throughput = n_cells / t_read_std_px if t_read_std_px > 0 else 0

    # --- Streaming Read-Only ---
    # openpyxl
    t0 = time.perf_counter()
    wb_ro_ox = openpyxl.load_workbook(target_file, read_only=True, data_only=True)
    rows_ro_ox = list(wb_ro_ox.active.iter_rows(values_only=True))
    wb_ro_ox.close()
    t_read_ro_ox = time.perf_counter() - t0

    # rpyxl
    t0 = time.perf_counter()
    wb_ro_px = rpyxl.load_workbook(target_file, read_only=True, data_only=True)
    rows_ro_px = list(wb_ro_px.active.iter_rows(values_only=True))
    wb_ro_px.close()
    t_read_ro_px = time.perf_counter() - t0

    read_ro_speedup = t_read_ro_ox / t_read_ro_px if t_read_ro_px > 0 else 0

    # --- Time to First Row (TTFR) ---
    t0 = time.perf_counter()
    wb_ttfr_ox = openpyxl.load_workbook(target_file, read_only=True, data_only=True)
    first_row_ox = next(wb_ttfr_ox.active.iter_rows(values_only=True))
    wb_ttfr_ox.close()
    ttfr_ox = time.perf_counter() - t0

    t0 = time.perf_counter()
    wb_ttfr_px = rpyxl.load_workbook(target_file, read_only=True, data_only=True)
    first_row_px = next(wb_ttfr_px.active.iter_rows(values_only=True))
    wb_ttfr_px.close()
    ttfr_px = time.perf_counter() - t0

    ttfr_speedup = ttfr_ox / ttfr_px if ttfr_px > 0 else 0

    read_results.append([
        f"{label} ({n_cells:,} cells)",
        f"{t_read_std_ox:.4f}s", f"{t_read_std_px:.4f}s", f"{read_std_speedup:.2f}x", f"{read_std_throughput:,.0f} c/s",
        f"{t_read_ro_ox:.4f}s", f"{t_read_ro_px:.4f}s", f"{read_ro_speedup:.2f}x",
        f"{ttfr_ox*1000:.2f}ms", f"{ttfr_px*1000:.2f}ms", f"{ttfr_speedup:.2f}x"
    ])

print(tabulate(
    read_results,
    headers=["Scale", "Std openpyxl", "Std rpyxl", "Std Speedup", "Std Throughput",
             "RO openpyxl", "RO rpyxl", "RO Speedup", "TTFR openpyxl", "TTFR rpyxl", "TTFR Speedup"],
    tablefmt="github"
))
print()

# ==============================================================================
# 3. TRANSACTIONAL MUTATION (READ-MODIFY-WRITE)
# ==============================================================================
print("--- 3. BENCHMARKING TRANSACTIONAL MUTATION (100,000 CELLS) ---")
mut_file_src = os.path.join(OUTPUT_DIR, "std_ox_10000.xlsx")

# openpyxl mutation
t0 = time.perf_counter()
wb = openpyxl.load_workbook(mut_file_src)
ws = wb.active
for r in range(1, 10001):
    ws.cell(row=r, column=2, value=r * 20)  # Double the integer
    ws.cell(row=r, column=4, value=False)   # Flip boolean
wb.save(os.path.join(OUTPUT_DIR, "mut_ox.xlsx"))
wb.close()
t_mut_ox = time.perf_counter() - t0

# rpyxl mutation
t0 = time.perf_counter()
wb = rpyxl.load_workbook(mut_file_src)
ws = wb.active
for r in range(1, 10001):
    ws.cell(row=r, column=2, value=r * 20)
    ws.cell(row=r, column=4, value=False)
wb.save(os.path.join(OUTPUT_DIR, "mut_px.xlsx"))
wb.close()
t_mut_px = time.perf_counter() - t0

mut_speedup = t_mut_ox / t_mut_px if t_mut_px > 0 else 0
print(tabulate([
    ["Read-Modify-Write (100k cells)", f"{t_mut_ox:.4f}s", f"{t_mut_px:.4f}s", f"{mut_speedup:.2f}x faster"]
], headers=["Operation", "openpyxl", "rpyxl", "Speedup"], tablefmt="github"))
print()

# ==============================================================================
# 4. CROSS-LIBRARY DATA INTEGRITY VALIDATION
# ==============================================================================
print("--- 4. CROSS-LIBRARY DATA INTEGRITY VERIFICATION ---")
check_rows = 500
test_rows = generate_sample_rows(check_rows)

# Case A: rpyxl writes -> openpyxl reads
path_a = os.path.join(OUTPUT_DIR, "interop_px_to_ox.xlsx")
wb = rpyxl.Workbook()
ws = wb.active
for r in test_rows:
    ws.append(r)
wb.save(path_a)
wb.close()

wb_in = openpyxl.load_workbook(path_a, data_only=True)
read_back_a = list(wb_in.active.iter_rows(values_only=True))
wb_in.close()

# Case B: openpyxl writes -> rpyxl reads
path_b = os.path.join(OUTPUT_DIR, "interop_ox_to_px.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
for r in test_rows:
    ws.append(r)
wb.save(path_b)
wb.close()

wb_in = rpyxl.load_workbook(path_b, data_only=True)
read_back_b = list(wb_in.active.iter_rows(values_only=True))
wb_in.close()

def compare_matrices(original, loaded, direction):
    mismatches = 0
    total = len(original) * len(original[0])
    for r_idx, (orig_row, load_row) in enumerate(zip(original, loaded)):
        for c_idx, (v_orig, v_load) in enumerate(zip(orig_row, load_row)):
            if isinstance(v_orig, float):
                match = abs(v_orig - v_load) < 1e-4
            else:
                match = (v_orig == v_load)
            if not match:
                mismatches += 1
                if mismatches <= 3:
                    print(f"  [Mismatch in {direction}] Row {r_idx+1}, Col {c_idx+1}: expected {v_orig!r}, got {v_load!r}")
    return mismatches, total

mismatches_a, total_a = compare_matrices(test_rows, read_back_a, "rpyxl -> openpyxl")
mismatches_b, total_b = compare_matrices(test_rows, read_back_b, "openpyxl -> rpyxl")

print(tabulate([
    ["rpyxl Write -> openpyxl Read", f"{total_a:,} cells", f"{mismatches_a} mismatches", "100% PERFECT" if mismatches_a == 0 else "MISMATCH FOUND"],
    ["openpyxl Write -> rpyxl Read", f"{total_b:,} cells", f"{mismatches_b} mismatches", "100% PERFECT" if mismatches_b == 0 else "MISMATCH FOUND"]
], headers=["Direction", "Total Cells Tested", "Mismatches", "Fidelity Status"], tablefmt="github"))

print("\n================================================================")
print(" LIVE BENCHMARK COMPLETE")
print("================================================================")

