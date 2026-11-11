#!/usr/bin/env python3
"""Styled write size sweep benchmark.

Measures styled write performance across row counts (10k, 50k, 100k, 250k, 500k):
- Build time (seconds)
- Save time (seconds)
- Total time (seconds)
- Peak RSS memory growth (MB)
- Output file size (MB)
"""

import argparse
import gc
import os
import tempfile
import time

try:
    import psutil
except ImportError:
    psutil = None


def get_rss_mb():
    if psutil is None:
        return 0.0
    return psutil.Process().memory_info().rss / (1024 * 1024)


def benchmark_styled_write(lib_name, n_rows, n_cols=5):
    if lib_name == "rpyxl":
        import rpyxl as lib
        from rpyxl.styles import Alignment, Border, Font, PatternFill, Side
    elif lib_name == "openpyxl":
        import openpyxl as lib
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    else:
        raise ValueError(f"Unknown library: {lib_name}")

    gc.collect()
    rss_start = get_rss_mb()

    # Pre-create style objects
    thin = Side(style="thin", color="D3D3D3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="Calibri", size=10)
    fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    align_center = Alignment(horizontal="center")
    align_right = Alignment(horizontal="right")

    t0 = time.perf_counter()
    wb = lib.Workbook()
    ws = wb.active
    ws.title = "Ledger"

    for r in range(1, n_rows + 1):
        ws.append([f"ACC-{r:06d}", "2026-11-11", r * 12.5, r * 8.25, r * 4.25])
        is_even = r % 2 == 0
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = font
            cell.border = border
            if is_even:
                cell.fill = fill
            if c <= 2:
                cell.alignment = align_center
            else:
                cell.alignment = align_right

    t_build = time.perf_counter() - t0
    rss_mid = get_rss_mb()

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        t1 = time.perf_counter()
        wb.save(tmp_path)
        t_save = time.perf_counter() - t1
        t_total = t_build + t_save
        rss_end = get_rss_mb()
        peak_rss_growth = max(rss_mid - rss_start, rss_end - rss_start)
        file_size_mb = os.path.getsize(tmp_path) / (1024 * 1024)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    return {
        "lib": lib_name,
        "rows": n_rows,
        "cols": n_cols,
        "build_s": t_build,
        "save_s": t_save,
        "total_s": t_total,
        "peak_rss_mb": peak_rss_growth,
        "file_size_mb": file_size_mb,
    }


def main():
    parser = argparse.ArgumentParser(description="Sweep styled write scaling")
    parser.add_argument(
        "--rows",
        nargs="+",
        type=int,
        default=[10000, 50000, 100000, 250000, 500000],
        help="Row counts to sweep",
    )
    parser.add_argument(
        "--libs",
        nargs="+",
        default=["rpyxl"],
        help="Libraries to benchmark (e.g. rpyxl openpyxl)",
    )
    args = parser.parse_args()

    print(
        f"{'Lib':<10} | {'Rows':<8} | {'Build (s)':<10} | {'Save (s)':<10} | {'Total (s)':<10} | {'Peak RSS (MB)':<14} | {'File (MB)':<10}"
    )
    print("-" * 85)

    for n in args.rows:
        for lib in args.libs:
            try:
                res = benchmark_styled_write(lib, n)
                print(
                    f"{res['lib']:<10} | {res['rows']:<8,d} | {res['build_s']:<10.3f} | {res['save_s']:<10.3f} | "
                    f"{res['total_s']:<10.3f} | {res['peak_rss_mb']:<14.1f} | {res['file_size_mb']:<10.2f}"
                )
            except Exception as e:
                print(f"{lib:<10} | {n:<8,d} | ERROR: {e}")


if __name__ == "__main__":
    main()
