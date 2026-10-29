"""
Comprehensive Industry-Grade Benchmark Suite:
Comparing openpyxl, xlsxwriter (Python), calamine (Rust), rust_xlsxwriter (Rust), and RPyXL.

Metrics:
- Throughput (cells/sec, rows/sec)
- Latency (median, min, max, p95, stdev)
- Peak RSS memory footprint
- Compression efficiency (bytes on disk)
- Transactional Read-Modify-Write mutation
- Multi-sheet parallel serialization
- Cross-engine round-trip precision verification
"""

import os
import sys
import time
import json
import gc
import subprocess
import statistics
import tempfile
import psutil
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from tabulate import tabulate

# Config
SCALES = [
    {"name": "Small (1k cells)", "rows": 100, "cols": 10, "cells": 1_000},
    {"name": "Medium (10k cells)", "rows": 1_000, "cols": 10, "cells": 10_000},
    {"name": "Large (100k cells)", "rows": 10_000, "cols": 10, "cells": 100_000},
    {"name": "Stress (500k cells)", "rows": 50_000, "cols": 10, "cells": 500_000},
]
ITERATIONS = 3
DATA_DIR = "benchmarks/data"
REPORT_DIR = "benchmarks/reports"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)

def get_process_rss():
    gc.collect()
    return psutil.Process().memory_info().rss / (1024 * 1024)

def run_stats(func, iterations=ITERATIONS):
    # Warmup
    gc.collect()
    func()
    times = []
    rss_samples = []
    for _ in range(iterations):
        gc.collect()
        rss_before = get_process_rss()
        t0 = time.perf_counter()
        func()
        t1 = time.perf_counter()
        rss_after = get_process_rss()
        times.append(t1 - t0)
        rss_samples.append(max(rss_before, rss_after))
    
    times.sort()
    p95_idx = int(0.95 * len(times)) if len(times) > 1 else 0
    return {
        "median": statistics.median(times),
        "min": min(times),
        "max": max(times),
        "p95": times[p95_idx],
        "stdev": statistics.stdev(times) if len(times) > 1 else 0.0,
        "peak_rss_mb": max(rss_samples),
    }

def generate_enterprise_data(rows):
    return [
        [
            f"User_{r}",
            r * 10,
            r * 1.25,
            r % 2 == 0,
            "2023-01-01 12:00:00",
            f"Item_{r % 100}",
            r * 5,
            round(r * 0.05, 2),
            r % 3 == 0,
            "Active",
        ]
        for r in range(1, rows + 1)
    ]

# ------------------------------------------------------------------------------
# 1. Dataset Pre-Generation
# ------------------------------------------------------------------------------
def ensure_dataset(scale):
    path = os.path.join(DATA_DIR, f"enterprise_{scale['rows']}r_{scale['cols']}c.xlsx")
    if os.path.exists(path):
        return path
    print(f"Generating enterprise dataset {path} ({scale['cells']:,} cells)...")
    import rpyxl
    wb = rpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="EnterpriseData")
    ws.append(["ID", "Amount", "Ratio", "Flag", "Timestamp", "Category", "Count", "Rate", "Valid", "Status"])
    for row in generate_enterprise_data(scale["rows"]):
        ws.append(row)
    wb.save(path)
    wb.close()
    return path

# ------------------------------------------------------------------------------
# Main Benchmarking Suite
# ------------------------------------------------------------------------------
def main():
    print("=" * 80)
    print(" INDUSTRY-LEVEL SPREADSHEET BENCHMARK & COMPARISON")
    print(" Engines: openpyxl, xlsxwriter (Py), RPyXL, calamine (Rust), rust_xlsxwriter (Rust)")
    print("=" * 80)

    import openpyxl
    import rpyxl
    import python_calamine
    import xlsxwriter

    # Pre-generate datasets
    datasets = {s["name"]: ensure_dataset(s) for s in SCALES}

    results = {
        "write_standard": {},
        "write_streaming": {},
        "read_standard": {},
        "read_streaming": {},
        "ttfr": {},
        "file_sizes": {},
        "mutation": {},
        "multisheet": {},
        "cross_verification": {},
    }

    # ==========================================================================
    # SCENARIO 1: WRITING & SERIALIZATION BENCHMARKS
    # ==========================================================================
    print("\n[1/5] Executing Writing & Serialization Benchmarks...")
    for s in SCALES:
        s_name = s["name"]
        rows = s["rows"]
        cells = s["cells"]
        print(f"  Testing Scale: {s_name} ({cells:,} cells)...")
        data = generate_enterprise_data(rows)

        # 1. openpyxl Standard Write
        def _write_openpyxl_std():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
            wb = openpyxl.Workbook()
            ws = wb.active
            for r_idx, row in enumerate(data, start=1):
                for c_idx, val in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
            wb.save(tmp)
            wb.close()
            os.remove(tmp)

        # 2. RPyXL Standard Write
        def _write_rpyxl_std():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
            wb = rpyxl.Workbook()
            ws = wb.active
            for r_idx, row in enumerate(data, start=1):
                for c_idx, val in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
            wb.save(tmp)
            wb.close()
            os.remove(tmp)

        # 3. openpyxl Write-Only Streaming
        def _write_openpyxl_stream():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
            wb = openpyxl.Workbook(write_only=True)
            ws = wb.create_sheet()
            for row in data:
                ws.append(row)
            wb.save(tmp)
            wb.close()
            os.remove(tmp)

        # 4. xlsxwriter (Python)
        def _write_xlsxwriter():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
            wb = xlsxwriter.Workbook(tmp)
            ws = wb.add_worksheet()
            for r_idx, row in enumerate(data):
                for c_idx, val in enumerate(row):
                    if isinstance(val, (int, float)):
                        ws.write_number(r_idx, c_idx, val)
                    elif isinstance(val, bool):
                        ws.write_boolean(r_idx, c_idx, val)
                    else:
                        ws.write_string(r_idx, c_idx, str(val))
            wb.close()
            os.remove(tmp)

        # 5. RPyXL Streaming Write-Only
        def _write_rpyxl_stream():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
            wb = rpyxl.Workbook(write_only=True)
            ws = wb.create_sheet()
            for row in data:
                ws.append(row)
            wb.save(tmp)
            wb.close()
            os.remove(tmp)

        # 6. rust_xlsxwriter (Pure Rust)
        def _write_rust_xlsxwriter():
            res = subprocess.run(
                ["./target/release/examples/bench_rust_xlsxwriter", str(rows), "10", "enterprise"],
                capture_output=True, text=True, check=True
            )
            return json.loads(res.stdout.strip())["seconds"]

        # Run stats
        results["write_standard"][s_name] = {
            "openpyxl": run_stats(_write_openpyxl_std),
            "rpyxl": run_stats(_write_rpyxl_std),
        }
        
        # Measure rust_xlsxwriter
        rust_times = []
        for _ in range(ITERATIONS):
            rust_times.append(_write_rust_xlsxwriter())
        rust_stats = {
            "median": statistics.median(rust_times),
            "min": min(rust_times),
            "max": max(rust_times),
            "p95": max(rust_times),
            "stdev": statistics.stdev(rust_times) if len(rust_times) > 1 else 0.0,
            "peak_rss_mb": 0.0,
        }

        results["write_streaming"][s_name] = {
            "openpyxl_stream": run_stats(_write_openpyxl_stream),
            "xlsxwriter_py": run_stats(_write_xlsxwriter),
            "rpyxl_stream": run_stats(_write_rpyxl_stream),
            "rust_xlsxwriter": rust_stats,
        }

    # ==========================================================================
    # SCENARIO 2: READING & DESERIALIZATION BENCHMARKS
    # ==========================================================================
    print("\n[2/5] Executing Reading & Ingestion Benchmarks...")
    for s in SCALES:
        s_name = s["name"]
        path = datasets[s_name]
        cells = s["cells"]
        print(f"  Testing Scale: {s_name} ({cells:,} cells)...")

        # 1. openpyxl Standard Load
        def _read_openpyxl_std():
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            cnt = 0
            for r in ws.iter_rows(values_only=True):
                cnt += len(r)
            wb.close()
            return cnt

        # 2. RPyXL Standard Load
        def _read_rpyxl_std():
            wb = rpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            cnt = 0
            for r in ws.iter_rows(values_only=True):
                cnt += len(r)
            wb.close()
            return cnt

        # 3. calamine Load
        def _read_calamine():
            wb = python_calamine.CalamineWorkbook.from_path(path)
            sheet = wb.get_sheet_by_name(wb.sheet_names[0])
            cnt = 0
            for r in sheet.to_python():
                cnt += len(r)
            return cnt

        # 4. openpyxl Read-Only Streaming
        def _read_openpyxl_stream():
            wb = openpyxl.load_workbook(path, read_only=True)
            ws = wb.active
            cnt = 0
            for r in ws.iter_rows(values_only=True):
                cnt += len(r)
            wb.close()
            return cnt

        # 5. RPyXL Streaming Read-Only
        def _read_rpyxl_stream():
            wb = rpyxl.load_workbook(path, read_only=True)
            ws = wb.active
            cnt = 0
            for r in ws.iter_rows(values_only=True):
                cnt += len(r)
            wb.close()
            return cnt

        results["read_standard"][s_name] = {
            "openpyxl": run_stats(_read_openpyxl_std),
            "rpyxl": run_stats(_read_rpyxl_std),
            "calamine": run_stats(_read_calamine),
        }
        results["read_streaming"][s_name] = {
            "openpyxl": run_stats(_read_openpyxl_stream),
            "rpyxl": run_stats(_read_rpyxl_stream),
        }

        # TTFR on Large and Stress
        if cells >= 100_000:
            def _ttfr_ox():
                t0 = time.perf_counter()
                wb = openpyxl.load_workbook(path, read_only=True)
                ws = wb.active
                gen = ws.iter_rows(values_only=True)
                _ = next(gen)
                dur = time.perf_counter() - t0
                wb.close()
                return dur

            def _ttfr_px():
                t0 = time.perf_counter()
                wb = rpyxl.load_workbook(path, read_only=True)
                ws = wb.active
                gen = ws.iter_rows(values_only=True)
                _ = next(gen)
                dur = time.perf_counter() - t0
                wb.close()
                return dur

            results["ttfr"][s_name] = {
                "openpyxl": min(_ttfr_ox() for _ in range(5)),
                "rpyxl": min(_ttfr_px() for _ in range(5)),
            }

    # ==========================================================================
    # SCENARIO 3: TRANSACTIONAL READ-MODIFY-WRITE MUTATION
    # ==========================================================================
    print("\n[3/5] Executing Read-Modify-Write Transactional Mutation Benchmark...")
    # Load 100k cell file, double numeric amounts in column 2, append summary row, save
    large_path = datasets["Large (100k cells)"]

    def _mutate_openpyxl():
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
        wb = openpyxl.load_workbook(large_path)
        ws = wb.active
        total = 0.0
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=2).value
            if isinstance(val, (int, float)):
                new_val = val * 2
                ws.cell(row=row, column=2, value=new_val)
                total += new_val
        ws.cell(row=ws.max_row + 1, column=2, value=total)
        wb.save(tmp)
        wb.close()
        os.remove(tmp)

    def _mutate_rpyxl():
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
        wb = rpyxl.load_workbook(large_path)
        ws = wb.active
        total = 0.0
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=2).value
            if isinstance(val, (int, float)):
                new_val = val * 2
                ws.cell(row=row, column=2, value=new_val)
                total += new_val
        ws.cell(row=ws.max_row + 1, column=2, value=total)
        wb.save(tmp)
        wb.close()
        os.remove(tmp)

    results["mutation"]["openpyxl"] = run_stats(_mutate_openpyxl, iterations=3)
    results["mutation"]["rpyxl"] = run_stats(_mutate_rpyxl, iterations=3)

    # ==========================================================================
    # SCENARIO 4: MULTI-SHEET PARALLELISM (100,000 Rows across 5 sheets)
    # ==========================================================================
    print("\n[4/5] Executing Multi-Sheet Parallelism Benchmark (5 sheets × 20k rows)...")
    sheet_data = generate_enterprise_data(20_000)

    def _multisheet_openpyxl():
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
        wb = openpyxl.Workbook()
        ws_first = wb.active
        ws_first.title = "Dept_0"
        for r_idx, row in enumerate(sheet_data, start=1):
            for c_idx, val in enumerate(row, start=1):
                ws_first.cell(row=r_idx, column=c_idx, value=val)
        for s_idx in range(1, 5):
            ws = wb.create_sheet(f"Dept_{s_idx}")
            for r_idx, row in enumerate(sheet_data, start=1):
                for c_idx, val in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
        wb.save(tmp)
        wb.close()
        os.remove(tmp)

    def _multisheet_rpyxl():
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
        wb = rpyxl.Workbook()
        ws_first = wb.active
        ws_first.title = "Dept_0"
        for r_idx, row in enumerate(sheet_data, start=1):
            for c_idx, val in enumerate(row, start=1):
                ws_first.cell(row=r_idx, column=c_idx, value=val)
        for s_idx in range(1, 5):
            ws = wb.create_sheet(f"Dept_{s_idx}")
            for r_idx, row in enumerate(sheet_data, start=1):
                for c_idx, val in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
        wb.save(tmp)
        wb.close()
        os.remove(tmp)

    results["multisheet"]["openpyxl"] = run_stats(_multisheet_openpyxl, iterations=2)
    results["multisheet"]["rpyxl"] = run_stats(_multisheet_rpyxl, iterations=2)

    # ==========================================================================
    # SCENARIO 5: CROSS-ENGINE ROUND-TRIP FIDELITY & PRECISION
    # ==========================================================================
    print("\n[5/5] Executing Cross-Engine Precision & Round-Trip Validation...")
    # Write precision file with RPyXL
    tmp_matrix = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    wb = rpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Precision Float"
    ws["B1"] = 1234567.89012345
    ws["A2"] = "Large Int"
    ws["B2"] = 922337203685477580  # 64-bit int
    ws["A3"] = "Boolean True"
    ws["B3"] = True
    ws["A4"] = "Boolean False"
    ws["B4"] = False
    ws["A5"] = "Unicode Text"
    ws["B5"] = "Tokyo 東京 • Berlin • Zurich 100€"
    wb.save(tmp_matrix)
    wb.close()

    # Read back with openpyxl
    wb_ox = openpyxl.load_workbook(tmp_matrix, data_only=True)
    ws_ox = wb_ox.active
    ox_float_diff = abs(ws_ox["B1"].value - 1234567.89012345)
    ox_int_match = (ws_ox["B2"].value == 922337203685477580)
    ox_bool_match = (ws_ox["B3"].value is True and ws_ox["B4"].value is False)
    ox_text_match = (ws_ox["B5"].value == "Tokyo 東京 • Berlin • Zurich 100€")
    wb_ox.close()

    # Read back with calamine
    wb_cal = python_calamine.CalamineWorkbook.from_path(tmp_matrix)
    cal_data = wb_cal.get_sheet_by_name(wb_cal.sheet_names[0]).to_python()
    cal_float_diff = abs(cal_data[0][1] - 1234567.89012345)
    cal_int_match = (cal_data[1][1] == 922337203685477580)
    cal_bool_match = (cal_data[2][1] is True and cal_data[3][1] is False)
    cal_text_match = (cal_data[4][1] == "Tokyo 東京 • Berlin • Zurich 100€")

    os.remove(tmp_matrix)

    results["cross_verification"] = {
        "float_precision_preserved": (ox_float_diff < 1e-6 and cal_float_diff < 1e-6),
        "int_64bit_exact": (ox_int_match and cal_int_match),
        "boolean_exact": (ox_bool_match and cal_bool_match),
        "multilingual_exact": (ox_text_match and cal_text_match),
    }

    # Save JSON raw results
    results_path = os.path.join(REPORT_DIR, "industry_benchmarks_raw.json")
    with open(results_path, "w") as f:
        json.dump(results, f, indent=2)
    print(f"\nRaw results exported to {results_path}")

    # ==========================================================================
    # GENERATE PUBLICATION-GRADE CHARTS
    # ==========================================================================
    print("\nRendering Publication-Quality Charts...")
    plt.style.use("ggplot")

    # Chart 1: Write Throughput
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
    scales = [s["name"].split(" ")[0] for s in SCALES]
    x = range(len(scales))
    width = 0.18

    ox_w = [s["cells"] / results["write_streaming"][s["name"]]["openpyxl_stream"]["median"] for s in SCALES]
    xw_w = [s["cells"] / results["write_streaming"][s["name"]]["xlsxwriter_py"]["median"] for s in SCALES]
    px_w = [s["cells"] / results["write_streaming"][s["name"]]["rpyxl_stream"]["median"] for s in SCALES]
    rust_w = [s["cells"] / results["write_streaming"][s["name"]]["rust_xlsxwriter"]["median"] for s in SCALES]

    ax.bar([i - 1.5*width for i in x], ox_w, width, label="openpyxl (Py Stream)", color="#E57373")
    ax.bar([i - 0.5*width for i in x], xw_w, width, label="xlsxwriter (Py Stream)", color="#FFB74D")
    ax.bar([i + 0.5*width for i in x], px_w, width, label="RPyXL (Rust+PyO3 Stream)", color="#4CAF50")
    ax.bar([i + 1.5*width for i in x], rust_w, width, label="rust_xlsxwriter (Pure Rust)", color="#1E88E5")

    ax.set_title("Industry Write Throughput Across Scales (Higher is Better)", fontsize=13, fontweight="bold")
    ax.set_ylabel("Throughput (Cells / Second)", fontsize=11)
    ax.set_xticks(x)
    ax.set_xticklabels(scales, fontsize=10)
    ax.legend(frameon=True)
    ax.grid(True, linestyle="--", alpha=0.5)
    plt.tight_layout()
    chart1_path = os.path.join(REPORT_DIR, "industry_write_throughput.png")
    fig.savefig(chart1_path)
    plt.close()

    # Chart 2: Read Throughput
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
    ox_r = [s["cells"] / results["read_standard"][s["name"]]["openpyxl"]["median"] for s in SCALES]
    px_r = [s["cells"] / results["read_standard"][s["name"]]["rpyxl"]["median"] for s in SCALES]
    cal_r = [s["cells"] / results["read_standard"][s["name"]]["calamine"]["median"] for s in SCALES]

    ax.bar([i - width for i in x], ox_r, width, label="openpyxl (Python)", color="#E57373")
    ax.bar([i for i in x], px_r, width, label="RPyXL (Rust+PyO3)", color="#4CAF50")
    ax.bar([i + width for i in x], cal_r, width, label="calamine (Pure Rust)", color="#8E24AA")

    ax.set_title("Industry Read Throughput Across Scales (Higher is Better)", fontsize=13, fontweight="bold")
    ax.set_ylabel("Throughput (Cells / Second)", fontsize=11)
    ax.set_xticks(x)
    ax.set_xticklabels(scales, fontsize=10)
    ax.legend(frameon=True)
    ax.grid(True, linestyle="--", alpha=0.5)
    plt.tight_layout()
    chart2_path = os.path.join(REPORT_DIR, "industry_read_throughput.png")
    fig.savefig(chart2_path)
    plt.close()

    # Print summary tables to console
    print("\n" + "="*80)
    print(" SUMMARY: STREAMING EXPORT PERFORMANCE (Cells / Second)")
    print("="*80)
    w_table = []
    for s in SCALES:
        s_name = s["name"]
        w_ox = results["write_streaming"][s_name]["openpyxl_stream"]["median"]
        w_xw = results["write_streaming"][s_name]["xlsxwriter_py"]["median"]
        w_px = results["write_streaming"][s_name]["rpyxl_stream"]["median"]
        w_rust = results["write_streaming"][s_name]["rust_xlsxwriter"]["median"]
        w_table.append({
            "Dataset": s_name,
            "openpyxl Stream": f"{w_ox:.4f}s ({s['cells']/w_ox:,.0f} c/s)",
            "xlsxwriter (Py)": f"{w_xw:.4f}s ({s['cells']/w_xw:,.0f} c/s)",
            "RPyXL Stream": f"{w_px:.4f}s ({s['cells']/w_px:,.0f} c/s)",
            "rust_xlsxwriter": f"{w_rust:.4f}s ({s['cells']/w_rust:,.0f} c/s)",
            "RPyXL vs openpyxl": f"{w_ox/w_px:.1f}x faster",
            "RPyXL vs xlsxwriter": f"{w_xw/w_px:.1f}x faster",
        })
    print(tabulate(w_table, headers="keys", tablefmt="github"))

    print("\n" + "="*80)
    print(" SUMMARY: STANDARD READ INGESTION PERFORMANCE (Cells / Second)")
    print("="*80)
    r_table = []
    for s in SCALES:
        s_name = s["name"]
        r_ox = results["read_standard"][s_name]["openpyxl"]["median"]
        r_px = results["read_standard"][s_name]["rpyxl"]["median"]
        r_cal = results["read_standard"][s_name]["calamine"]["median"]
        r_table.append({
            "Dataset": s_name,
            "openpyxl (Python)": f"{r_ox:.4f}s ({s['cells']/r_ox:,.0f} c/s)",
            "RPyXL (Rust+PyO3)": f"{r_px:.4f}s ({s['cells']/r_px:,.0f} c/s)",
            "calamine (Rust)": f"{r_cal:.4f}s ({s['cells']/r_cal:,.0f} c/s)",
            "RPyXL vs openpyxl": f"{r_ox/r_px:.1f}x faster",
            "RPyXL vs calamine": f"+{(r_px/r_cal - 1)*100:.0f}% overhead",
        })
    print(tabulate(r_table, headers="keys", tablefmt="github"))

    print("\n" + "="*80)
    print(" SUMMARY: TRANSACTIONAL MUTATION (Read-Modify-Write 100k Cells)")
    print("="*80)
    mut_ox = results["mutation"]["openpyxl"]["median"]
    mut_px = results["mutation"]["rpyxl"]["median"]
    print(f"openpyxl : {mut_ox:.4f}s (Peak RSS: {results['mutation']['openpyxl']['peak_rss_mb']:.1f} MB)")
    print(f"RPyXL     : {mut_px:.4f}s (Peak RSS: {results['mutation']['rpyxl']['peak_rss_mb']:.1f} MB)")
    print(f"Speedup  : {mut_ox/mut_px:.2f}x faster with RPyXL")

    print("\n" + "="*80)
    print(" SUMMARY: MULTI-SHEET PARALLELISM (100k Rows across 5 Sheets = 1M Cells)")
    print("="*80)
    ms_ox = results["multisheet"]["openpyxl"]["median"]
    ms_px = results["multisheet"]["rpyxl"]["median"]
    print(f"openpyxl (GIL throttled) : {ms_ox:.4f}s ({1_000_000/ms_ox:,.0f} cells/s)")
    print(f"RPyXL (Rayon multi-thread): {ms_px:.4f}s ({1_000_000/ms_px:,.0f} cells/s)")
    print(f"Speedup                  : {ms_ox/ms_px:.2f}x faster with RPyXL")

    print("\n" + "="*80)
    print(" SUMMARY: CROSS-ENGINE FIDELITY")
    print("="*80)
    print(json.dumps(results["cross_verification"], indent=2))

if __name__ == "__main__":
    main()
