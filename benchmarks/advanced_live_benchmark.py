#!/usr/bin/env python3
"""
Advanced Live Benchmark: rpyxl vs openpyxl
Tests:
1. Multi-Sheet Rayon Concurrency (5 sheets x 20,000 rows = 1,000,000 cells)
2. Heavy Cell Styling (Fonts, Fills, Borders, Number Formats)
3. Subscript Slicing & Indexing Operations (2D Ranges, Column Iteration)
4. Memory RSS Tracking Under Sustained Load
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
from rpyxl.styles import Font as RFont, PatternFill as RPatternFill, Alignment as RAlignment, Border as RBorder, Side as RSide

print("================================================================")
print(f" ADVANCED LIVE BENCHMARK: Multi-Sheet, Styling, Slicing & Memory")
print(f" Python: {sys.version.split()[0]} | openpyxl: {openpyxl.__version__} | rpyxl: {rpyxl.__version__}")
print(f" CPU Cores: {psutil.cpu_count(logical=True)} logical ({psutil.cpu_count(logical=False)} physical)")
print("================================================================\n")

OUTPUT_DIR = "/tmp/advanced_bench_data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_rss_mb():
    gc.collect()
    return psutil.Process().memory_info().rss / (1024 * 1024)

# ==============================================================================
# 1. MULTI-SHEET PARALLELISM (1,000,000 CELLS ACROSS 5 SHEETS)
# ==============================================================================
print("--- 1. MULTI-SHEET PARALLELISM (5 SHEETS x 20,000 ROWS = 1,000,000 CELLS) ---")
N_SHEETS = 5
ROWS_PER_SHEET = 20000
COLS = 10
TOTAL_CELLS = N_SHEETS * ROWS_PER_SHEET * COLS

# openpyxl multi-sheet
file_ox_multi = os.path.join(OUTPUT_DIR, "multi_ox.xlsx")
rss_start = get_rss_mb()
t0 = time.perf_counter()
wb_ox = openpyxl.Workbook()
for s in range(N_SHEETS):
    title = f"Department_{s+1}"
    ws = wb_ox.active if s == 0 else wb_ox.create_sheet(title=title)
    if s == 0:
        ws.title = title
    for r in range(1, ROWS_PER_SHEET + 1):
        ws.append([f"S{s+1}_R{r}_C{c}" for c in range(1, COLS + 1)])
wb_ox.save(file_ox_multi)
wb_ox.close()
t_multi_ox = time.perf_counter() - t0
rss_multi_ox = get_rss_mb() - rss_start

# rpyxl multi-sheet (Rayon parallelism in Rust)
file_px_multi = os.path.join(OUTPUT_DIR, "multi_px.xlsx")
rss_start = get_rss_mb()
t0 = time.perf_counter()
wb_px = rpyxl.Workbook()
for s in range(N_SHEETS):
    title = f"Department_{s+1}"
    ws = wb_px.active if s == 0 else wb_px.create_sheet(title=title)
    if s == 0:
        ws.title = title
    for r in range(1, ROWS_PER_SHEET + 1):
        ws.append([f"S{s+1}_R{r}_C{c}" for c in range(1, COLS + 1)])
wb_px.save(file_px_multi)
wb_px.close()
t_multi_px = time.perf_counter() - t0
rss_multi_px = get_rss_mb() - rss_start

multi_speedup = t_multi_ox / t_multi_px if t_multi_px > 0 else 0
multi_throughput = TOTAL_CELLS / t_multi_px if t_multi_px > 0 else 0

print(tabulate([
    [
        f"Multi-Sheet Save ({TOTAL_CELLS:,} cells)",
        f"{t_multi_ox:.4f}s", f"{t_multi_px:.4f}s",
        f"{multi_speedup:.2f}x faster",
        f"{multi_throughput:,.0f} cells/s",
        f"{rss_multi_ox:+.1f} MB", f"{rss_multi_px:+.1f} MB"
    ]
], headers=["Workload", "openpyxl (Wall)", "rpyxl (Wall)", "Speedup", "rpyxl Throughput", "openpyxl RSS Δ", "rpyxl RSS Δ"], tablefmt="github"))
print()

# ==============================================================================
# 2. HEAVY STYLING & NUMBER FORMATTING BENCHMARK
# ==============================================================================
print("--- 2. HEAVY STYLING BENCHMARK (10,000 ROWS x 5 COLUMNS = 50,000 STYLED CELLS) ---")
N_STYLE_ROWS = 10000

# Styles definition
thin_border = Side(style="thin", color="000000")
ox_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
ox_font_head = Font(name="Arial", size=12, bold=True, color="FFFFFF")
ox_fill_head = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
ox_font_data = Font(name="Calibri", size=10, italic=False)
ox_fill_alt = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

r_thin_border = RSide(style="thin", color="000000")
r_border = RBorder(left=r_thin_border, right=r_thin_border, top=r_thin_border, bottom=r_thin_border)
r_font_head = RFont(name="Arial", size=12, bold=True, color="FFFFFF")
r_fill_head = RPatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
r_font_data = RFont(name="Calibri", size=10, italic=False)
r_fill_alt = RPatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

# openpyxl styled write
file_ox_styled = os.path.join(OUTPUT_DIR, "styled_ox.xlsx")
rss_start = get_rss_mb()
t0 = time.perf_counter()
wb = openpyxl.Workbook()
ws = wb.active
headers = ["Account ID", "Transaction Date", "Debit Amount", "Credit Amount", "Net Balance"]
ws.append(headers)
for col_idx in range(1, 6):
    cell = ws.cell(row=1, column=col_idx)
    cell.font = ox_font_head
    cell.fill = ox_fill_head
    cell.border = ox_border

for r in range(2, N_STYLE_ROWS + 2):
    ws.append([f"ACC-{r:05d}", "2026-11-11", r * 15.5, r * 12.0, r * 3.5])
    fill_to_use = ox_fill_alt if r % 2 == 0 else None
    for c in range(1, 6):
        cell = ws.cell(row=r, column=c)
        cell.font = ox_font_data
        cell.border = ox_border
        if fill_to_use:
            cell.fill = fill_to_use
        if c in (3, 4, 5):
            cell.number_format = "$#,##0.00"

wb.save(file_ox_styled)
wb.close()
t_style_ox = time.perf_counter() - t0
rss_style_ox = get_rss_mb() - rss_start

# rpyxl styled write
file_px_styled = os.path.join(OUTPUT_DIR, "styled_px.xlsx")
rss_start = get_rss_mb()
t0 = time.perf_counter()
wb = rpyxl.Workbook()
ws = wb.active
ws.append(headers)
for col_idx in range(1, 6):
    cell = ws.cell(row=1, column=col_idx)
    cell.font = r_font_head
    cell.fill = r_fill_head
    cell.border = r_border

for r in range(2, N_STYLE_ROWS + 2):
    ws.append([f"ACC-{r:05d}", "2026-11-11", r * 15.5, r * 12.0, r * 3.5])
    fill_to_use = r_fill_alt if r % 2 == 0 else None
    for c in range(1, 6):
        cell = ws.cell(row=r, column=c)
        cell.font = r_font_data
        cell.border = r_border
        if fill_to_use:
            cell.fill = fill_to_use
        if c in (3, 4, 5):
            cell.number_format = "$#,##0.00"

wb.save(file_px_styled)
wb.close()
t_style_px = time.perf_counter() - t0
rss_style_px = get_rss_mb() - rss_start

style_speedup = t_style_ox / t_style_px if t_style_px > 0 else 0
print(tabulate([
    [
        "Styled Sheet (50k styled cells)",
        f"{t_style_ox:.4f}s", f"{t_style_px:.4f}s",
        f"{style_speedup:.2f}x faster",
        f"{rss_style_ox:+.1f} MB", f"{rss_style_px:+.1f} MB"
    ]
], headers=["Workload", "openpyxl (Wall)", "rpyxl (Wall)", "Speedup", "openpyxl RSS Δ", "rpyxl RSS Δ"], tablefmt="github"))
print()

# ==============================================================================
# 3. 2D RANGE SLICING & COLUMN ITERATION
# ==============================================================================
print("--- 3. 2D RANGE SLICING & CELL ACCESS PATTERNS ---")

# Test file with 10,000 rows x 10 cols
test_file = os.path.join(OUTPUT_DIR, "slice_test.xlsx")
wb = rpyxl.Workbook()
ws = wb.active
for r in range(1, 10001):
    ws.append([r * c for c in range(1, 11)])
wb.save(test_file)
wb.close()

# A. 2D Block Slicing ws["A1:J5000"] (50,000 cells)
t0 = time.perf_counter()
wb = openpyxl.load_workbook(test_file)
block_ox = [[cell.value for cell in row] for row in wb.active["A1:J5000"]]
wb.close()
t_slice_ox = time.perf_counter() - t0

t0 = time.perf_counter()
wb = rpyxl.load_workbook(test_file)
block_px = [[cell.value for cell in row] for row in wb.active["A1:J5000"]]
wb.close()
t_slice_px = time.perf_counter() - t0

slice_speedup = t_slice_ox / t_slice_px if t_slice_px > 0 else 0

# B. Column-major Iteration: iter_cols()
t0 = time.perf_counter()
wb = openpyxl.load_workbook(test_file)
cols_ox = [[cell.value for cell in col] for col in wb.active.iter_cols(min_col=1, max_col=10, min_row=1, max_row=5000)]
wb.close()
t_cols_ox = time.perf_counter() - t0

t0 = time.perf_counter()
wb = rpyxl.load_workbook(test_file)
cols_px = [[cell.value for cell in col] for col in wb.active.iter_cols(min_col=1, max_col=10, min_row=1, max_row=5000)]
wb.close()
t_cols_px = time.perf_counter() - t0

cols_speedup = t_cols_ox / t_cols_px if t_cols_px > 0 else 0

print(tabulate([
    ["2D Block Slicing ws['A1:J5000']", f"{t_slice_ox:.4f}s", f"{t_slice_px:.4f}s", f"{slice_speedup:.2f}x faster"],
    ["Column Iteration iter_cols()", f"{t_cols_ox:.4f}s", f"{t_cols_px:.4f}s", f"{cols_speedup:.2f}x faster"]
], headers=["Access Pattern", "openpyxl", "rpyxl", "Speedup"], tablefmt="github"))
print()

# ==============================================================================
# 4. LARGE WORKBOOK LOAD & SAVE RE-SERIALIZATION
# ==============================================================================
print("--- 4. FULL RE-SERIALIZATION CYCLE (LOAD -> UNCHANGED SAVE) ---")
# Using the 1,000,000-cell multi-sheet file from Section 1
t0 = time.perf_counter()
wb = openpyxl.load_workbook(file_ox_multi)
wb.save(os.path.join(OUTPUT_DIR, "resave_ox.xlsx"))
wb.close()
t_resave_ox = time.perf_counter() - t0

t0 = time.perf_counter()
wb = rpyxl.load_workbook(file_px_multi)
wb.save(os.path.join(OUTPUT_DIR, "resave_px.xlsx"))
wb.close()
t_resave_px = time.perf_counter() - t0

resave_speedup = t_resave_ox / t_resave_px if t_resave_px > 0 else 0

print(tabulate([
    ["Full 1M Cell Re-save Cycle", f"{t_resave_ox:.4f}s", f"{t_resave_px:.4f}s", f"{resave_speedup:.2f}x faster"]
], headers=["Operation", "openpyxl", "rpyxl", "Speedup"], tablefmt="github"))

print("\n================================================================")
print(" ADVANCED BENCHMARK COMPLETE")
print("================================================================")
