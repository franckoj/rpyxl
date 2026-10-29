import time
import os
import psutil
import gc
from tabulate import tabulate
import openpyxl
import rpyxl
import python_calamine

def get_process_rss():
    gc.collect()
    return psutil.Process().memory_info().rss / (1024 * 1024)

files = [
    ("Small (1k cells)", "benchmarks/data/scale_100r_10c.xlsx", 1000),
    ("Medium (10k cells)", "benchmarks/data/scale_1000r_10c.xlsx", 10000),
    ("Large (100k cells)", "benchmarks/data/scale_10000r_10c.xlsx", 100000),
    ("Stress (500k cells)", "benchmarks/data/scale_50000r_10c.xlsx", 500000),
]

load_results = []
ttfr_results = []

for label, path, cells in files:
    # 1. openpyxl standard load
    t0 = time.perf_counter()
    wb_ox = openpyxl.load_workbook(path, data_only=True)
    ws_ox = wb_ox.active
    rows_ox = list(ws_ox.iter_rows(values_only=True))
    wb_ox.close()
    t_ox = time.perf_counter() - t0

    # 2. rpyxl standard load
    t0 = time.perf_counter()
    wb_px = rpyxl.load_workbook(path, data_only=True)
    ws_px = wb_px.active
    rows_px = list(ws_px.iter_rows(values_only=True))
    wb_px.close()
    t_px = time.perf_counter() - t0

    # 3. calamine load
    t0 = time.perf_counter()
    wb_cal = python_calamine.CalamineWorkbook.from_path(path)
    sheet_cal = wb_cal.get_sheet_by_name(wb_cal.sheet_names[0])
    rows_cal = sheet_cal.to_python()
    t_cal = time.perf_counter() - t0

    # Speedups
    speedup_px_vs_ox = t_ox / t_px if t_px > 0 else 0
    speedup_cal_vs_ox = t_ox / t_cal if t_cal > 0 else 0
    ratio_px_vs_cal = t_px / t_cal if t_cal > 0 else 0

    load_results.append({
        "Dataset": label,
        "Cells": f"{cells:,}",
        "openpyxl (Python)": f"{t_ox:.4f}s",
        "RPyXL (Rust+PyO3)": f"{t_px:.4f}s",
        "calamine (Pure Rust)": f"{t_cal:.4f}s",
        "RPyXL Speedup vs openpyxl": f"{speedup_px_vs_ox:.2f}x",
        "calamine Speedup vs openpyxl": f"{speedup_cal_vs_ox:.2f}x",
        "RPyXL overhead vs Calamine": f"{ratio_px_vs_cal:.2f}x",
    })

    # Time to First Row (TTFR) on Large & Stress
    if cells >= 100000:
        # openpyxl read_only TTFR
        t0 = time.perf_counter()
        wb_ro_ox = openpyxl.load_workbook(path, read_only=True)
        ws_ro_ox = wb_ro_ox.active
        gen_ox = ws_ro_ox.iter_rows(values_only=True)
        first_row_ox = next(gen_ox)
        ttfr_ox = time.perf_counter() - t0
        wb_ro_ox.close()

        # rpyxl read_only TTFR
        t0 = time.perf_counter()
        wb_ro_px = rpyxl.load_workbook(path, read_only=True)
        ws_ro_px = wb_ro_px.active
        gen_px = ws_ro_px.iter_rows(values_only=True)
        first_row_px = next(gen_px)
        ttfr_px = time.perf_counter() - t0
        wb_ro_px.close()

        ttfr_results.append({
            "Dataset": label,
            "openpyxl TTFR": f"{ttfr_ox:.4f}s",
            "RPyXL TTFR": f"{ttfr_px*1000:.2f}ms",
            "TTFR Latency Advantage": f"{ttfr_ox/ttfr_px:.1f}x faster",
        })

print("\n=== STANDARD LOAD & FULL MATRIX ITERATION ===")
print(tabulate(load_results, headers="keys", tablefmt="github"))

print("\n=== TIME-TO-FIRST-ROW (STREAMING LATENCY) ===")
print(tabulate(ttfr_results, headers="keys", tablefmt="github"))

# 3. Writing Benchmarks (openpyxl vs xlsxwriter vs RPyXL Standard vs RPyXL Write-Only vs rust_xlsxwriter)
import subprocess
import json

try:
    import xlsxwriter
    has_xlsxwriter = True
except ImportError:
    has_xlsxwriter = False

write_results = []
for label, _, cells in files:
    rows = cells // 10
    cols = 10
    
    # openpyxl standard write
    t0 = time.perf_counter()
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c, value=r * 10 + c)
    out_ox = f"/tmp/ox_{cells}.xlsx"
    wb.save(out_ox)
    wb.close()
    t_write_ox = time.perf_counter() - t0
    if os.path.exists(out_ox): os.remove(out_ox)

    # python xlsxwriter
    if has_xlsxwriter:
        out_xw = f"/tmp/xw_{cells}.xlsx"
        t0 = time.perf_counter()
        wb_xw = xlsxwriter.Workbook(out_xw)
        ws_xw = wb_xw.add_worksheet()
        for r in range(rows):
            for c in range(cols):
                ws_xw.write_number(r, c, r * 10 + c)
        wb_xw.close()
        t_write_xw = time.perf_counter() - t0
        if os.path.exists(out_xw): os.remove(out_xw)
        xw_str = f"{t_write_xw:.4f}s"
    else:
        xw_str = "N/A"

    # rpyxl standard write
    t0 = time.perf_counter()
    wb = rpyxl.Workbook()
    ws = wb.active
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c, value=r * 10 + c)
    out_px = f"/tmp/px_{cells}.xlsx"
    wb.save(out_px)
    wb.close()
    t_write_px = time.perf_counter() - t0
    if os.path.exists(out_px): os.remove(out_px)

    # rpyxl streaming write-only
    t0 = time.perf_counter()
    wb_wo = rpyxl.Workbook(write_only=True)
    ws_wo = wb_wo.create_sheet()
    for r in range(rows):
        ws_wo.append([r * 10 + c for c in range(cols)])
    out_wo = f"/tmp/wo_{cells}.xlsx"
    wb_wo.save(out_wo)
    wb_wo.close()
    t_write_wo = time.perf_counter() - t0
    if os.path.exists(out_wo): os.remove(out_wo)

    # rust_xlsxwriter (pure Rust binary)
    try:
        proc = subprocess.run(
            ["./target/release/examples/bench_rust_xlsxwriter", str(rows), str(cols), "standard"],
            capture_output=True, text=True, check=True
        )
        data = json.loads(proc.stdout.strip())
        t_rust_xlsx = data["seconds"]
        rust_str = f"{t_rust_xlsx:.4f}s"
    except Exception:
        rust_str = "N/A"

    speedup_px_std = t_write_ox / t_write_px if t_write_px > 0 else 0
    speedup_px_wo = t_write_ox / t_write_wo if t_write_wo > 0 else 0

    write_results.append({
        "Dataset": label,
        "Cells": f"{cells:,}",
        "openpyxl": f"{t_write_ox:.4f}s",
        "xlsxwriter (Py)": xw_str,
        "RPyXL Std": f"{t_write_px:.4f}s ({speedup_px_std:.1f}x)",
        "RPyXL Write-Only": f"{t_write_wo:.4f}s ({speedup_px_wo:.1f}x)",
        "rust_xlsxwriter (Rust)": rust_str,
        "calamine": "N/A (Read-Only)",
    })

print("\n=== COMPREHENSIVE WRITING BENCHMARKS ===")
print(tabulate(write_results, headers="keys", tablefmt="github"))


