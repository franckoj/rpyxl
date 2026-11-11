#!/usr/bin/env python3
"""
Deep-Dive Benchmark: Styles Write Speed & Memory Consumption
openpyxl (3.1.5) vs rpyxl (0.2.0)
"""

import os
import sys
import time
import gc
import psutil
from tabulate import tabulate

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import rpyxl
from rpyxl.styles import (
    Font as RFont,
    PatternFill as RPatternFill,
    Alignment as RAlignment,
    Border as RBorder,
    Side as RSide,
)

print("====================================================================")
print(f" STYLES WRITE SPEED & MEMORY CONSUMPTION BENCHMARK")
print(f" Python: {sys.version.split()[0]} | openpyxl: {openpyxl.__version__} | rpyxl: {rpyxl.__version__}")
print(f" Platform: {sys.platform} | CPU Cores: {psutil.cpu_count(logical=True)}")
print("====================================================================\n")

OUTPUT_DIR = "/tmp/styles_bench_data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def measure_rss_mb():
    gc.collect()
    return psutil.Process().memory_info().rss / (1024 * 1024)

# ==============================================================================
# Workload Configurations
# ==============================================================================
STYLE_SCALES = [
    ("Small", 5000),      # 5,000 rows x 5 cols = 25,000 styled cells
    ("Medium", 20000),    # 20,000 rows x 5 cols = 100,000 styled cells
    ("Large", 50000),     # 50,000 rows x 5 cols = 250,000 styled cells
]

headers = ["Account Code", "Posting Date", "Debit ($)", "Credit ($)", "Net Position ($)"]

# Reusable openpyxl styles
ox_thin = Side(style="thin", color="D3D3D3")
ox_border = Border(left=ox_thin, right=ox_thin, top=ox_thin, bottom=ox_thin)
ox_header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
ox_header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
ox_header_align = Alignment(horizontal="center", vertical="center")

ox_data_font = Font(name="Calibri", size=10)
ox_zebra_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
ox_align_right = Alignment(horizontal="right")
ox_align_center = Alignment(horizontal="center")

# Reusable rpyxl styles
r_thin = RSide(style="thin", color="D3D3D3")
r_border = RBorder(left=r_thin, right=r_thin, top=r_thin, bottom=r_thin)
r_header_font = RFont(name="Arial", size=11, bold=True, color="FFFFFF")
r_header_fill = RPatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
r_header_align = RAlignment(horizontal="center", vertical="center")

r_data_font = RFont(name="Calibri", size=10)
r_zebra_fill = RPatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
r_align_right = RAlignment(horizontal="right")
r_align_center = RAlignment(horizontal="center")

results = []

for label, n_rows in STYLE_SCALES:
    n_cells = n_rows * 5
    print(f"--- Running {label} Scale ({n_rows:,} rows = {n_cells:,} styled cells) ---")

    # --------------------------------------------------------------------------
    # 1. openpyxl Styled Run
    # --------------------------------------------------------------------------
    gc.collect()
    rss_before_ox = measure_rss_mb()
    t0 = time.perf_counter()

    wb_ox = openpyxl.Workbook()
    ws_ox = wb_ox.active
    ws_ox.title = "Ledger"

    # Header Row
    ws_ox.append(headers)
    for col_idx in range(1, 6):
        c = ws_ox.cell(row=1, column=col_idx)
        c.font = ox_header_font
        c.fill = ox_header_fill
        c.alignment = ox_header_align
        c.border = ox_border

    # Data Rows
    for r in range(2, n_rows + 2):
        ws_ox.append([
            f"ACC-{r:06d}",
            "2026-11-11",
            round(r * 12.5, 2),
            round(r * 8.25, 2),
            round(r * 4.25, 2)
        ])
        is_even = (r % 2 == 0)
        # Apply styles to row
        for col_idx in range(1, 6):
            c = ws_ox.cell(row=r, column=col_idx)
            c.font = ox_data_font
            c.border = ox_border
            if is_even:
                c.fill = ox_zebra_fill
            if col_idx == 1:
                c.alignment = ox_align_center
            elif col_idx == 2:
                c.alignment = ox_align_center
                c.number_format = "YYYY-MM-DD"
            else:
                c.alignment = ox_align_right
                c.number_format = "$#,##0.00"

    t_build_ox = time.perf_counter() - t0
    rss_peak_build_ox = measure_rss_mb()

    # Save phase
    t_save_start = time.perf_counter()
    ox_file = os.path.join(OUTPUT_DIR, f"styled_ox_{n_rows}.xlsx")
    wb_ox.save(ox_file)
    wb_ox.close()
    t_save_ox = time.perf_counter() - t_save_start
    t_total_ox = time.perf_counter() - t0
    rss_final_ox = measure_rss_mb()
    delta_rss_ox = rss_peak_build_ox - rss_before_ox
    ox_size_kb = os.path.getsize(ox_file) / 1024

    # --------------------------------------------------------------------------
    # 2. rpyxl Styled Run
    # --------------------------------------------------------------------------
    gc.collect()
    rss_before_px = measure_rss_mb()
    t0 = time.perf_counter()

    wb_px = rpyxl.Workbook()
    ws_px = wb_px.active
    ws_px.title = "Ledger"

    # Header Row
    ws_px.append(headers)
    for col_idx in range(1, 6):
        c = ws_px.cell(row=1, column=col_idx)
        c.font = r_header_font
        c.fill = r_header_fill
        c.alignment = r_header_align
        c.border = r_border

    # Data Rows
    for r in range(2, n_rows + 2):
        ws_px.append([
            f"ACC-{r:06d}",
            "2026-11-11",
            round(r * 12.5, 2),
            round(r * 8.25, 2),
            round(r * 4.25, 2)
        ])
        is_even = (r % 2 == 0)
        # Apply styles to row
        for col_idx in range(1, 6):
            c = ws_px.cell(row=r, column=col_idx)
            c.font = r_data_font
            c.border = r_border
            if is_even:
                c.fill = r_zebra_fill
            if col_idx == 1:
                c.alignment = r_align_center
            elif col_idx == 2:
                c.alignment = r_align_center
                c.number_format = "YYYY-MM-DD"
            else:
                c.alignment = r_align_right
                c.number_format = "$#,##0.00"

    t_build_px = time.perf_counter() - t0
    rss_peak_build_px = measure_rss_mb()

    # Save phase
    t_save_start = time.perf_counter()
    px_file = os.path.join(OUTPUT_DIR, f"styled_px_{n_rows}.xlsx")
    wb_px.save(px_file)
    wb_px.close()
    t_save_px = time.perf_counter() - t_save_start
    t_total_px = time.perf_counter() - t0
    rss_final_px = measure_rss_mb()
    delta_rss_px = rss_peak_build_px - rss_before_px
    px_size_kb = os.path.getsize(px_file) / 1024

    # Calculate metrics
    total_speedup = t_total_ox / t_total_px if t_total_px > 0 else 0
    save_speedup = t_save_ox / t_save_px if t_save_px > 0 else 0
    styled_throughput = n_cells / t_total_px if t_total_px > 0 else 0

    results.append({
        "scale": f"{label} ({n_cells:,} cells)",
        "ox_total": t_total_ox,
        "px_total": t_total_px,
        "speedup": total_speedup,
        "ox_save": t_save_ox,
        "px_save": t_save_px,
        "save_speedup": save_speedup,
        "ox_rss": delta_rss_ox,
        "px_rss": delta_rss_px,
        "throughput": styled_throughput,
        "ox_size": ox_size_kb,
        "px_size": px_size_kb,
    })

print("\n" + "="*80)
print(" STYLES BENCHMARK SUMMARY TABLE")
print("="*80)

table_data = []
for r in results:
    table_data.append([
        r["scale"],
        f"{r['ox_total']:.3f}s",
        f"{r['px_total']:.3f}s",
        f"{r['speedup']:.2f}x",
        f"{r['ox_save']:.3f}s",
        f"{r['px_save']:.3f}s",
        f"{r['save_speedup']:.2f}x",
        f"{r['ox_rss']:+.1f} MB",
        f"{r['px_rss']:+.1f} MB",
        f"{r['throughput']:,.0f} c/s"
    ])

print(tabulate(
    table_data,
    headers=[
        "Scale",
        "openpyxl Total", "rpyxl Total", "Total Speedup",
        "openpyxl Save", "rpyxl Save", "Save Speedup",
        "openpyxl ΔRSS", "rpyxl ΔRSS",
        "rpyxl Throughput"
    ],
    tablefmt="github"
))

print("\n" + "="*80)
print(" FILE SIZE & FORMATTING FIDELITY CHECK")
print("="*80)

size_table = []
for r in results:
    size_table.append([
        r["scale"],
        f"{r['ox_size']:,.1f} KB",
        f"{r['px_size']:,.1f} KB",
        f"{(r['px_size'] / r['ox_size'] * 100):.1f}% of openpyxl"
    ])

print(tabulate(
    size_table,
    headers=["Scale", "openpyxl Size", "rpyxl Size", "Ratio (rpyxl / openpyxl)"],
    tablefmt="github"
))

