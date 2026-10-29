"""Phase 5 Multi-Scale Benchmark Runner: RPyXL vs openpyxl vs calamine.

Scales:
- Small: 1k cells (100 rows x 10 cols)
- Medium: 10k cells (1,000 rows x 10 cols)
- Large: 100k cells (10,000 rows x 10 cols)
- Stress: 500k cells (50,000 rows x 10 cols)

Measures:
- Standard In-Memory Workbook Save
- Streaming Write-Only Workbook
- Standard In-Memory Workbook Load
- Streaming Read-Only Workbook
- CPU Utilization (%), CPU Time (s), Peak RSS (MB), Wall Time (s), Throughput, Speedup
- Generates publication-ready PNG charts and Markdown/JSON comparison reports
"""

import gc
import json
import os
import platform
import statistics
import sys
import tempfile
import threading
import time
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import psutil

REPETITIONS = 3
WARMUP_RUNS = 1

SCALES = [
    {"name": "Small (1k cells)", "rows": 100, "cols": 10, "cells": 1_000},
    {"name": "Medium (10k cells)", "rows": 1_000, "cols": 10, "cells": 10_000},
    {"name": "Large (100k cells)", "rows": 10_000, "cols": 10, "cells": 100_000},
    {"name": "Stress (500k cells)", "rows": 50_000, "cols": 10, "cells": 500_000},
]


def get_environment_info():
    uname = platform.uname()
    return {
        "timestamp": datetime.utcnow().isoformat(),
        "os": uname.system,
        "os_release": uname.release,
        "machine": uname.machine,
        "processor": uname.processor,
        "cpu_count_logical": psutil.cpu_count(logical=True),
        "cpu_count_physical": psutil.cpu_count(logical=False),
        "total_ram_mb": round(psutil.virtual_memory().total / (1024 * 1024), 2),
        "python_version": sys.version,
    }


class CPUMonitor:
    def __init__(self, interval=0.01):
        self.interval = interval
        self.samples = []
        self._stop = False
        self._thread = None
        self.proc = psutil.Process()

    def start(self):
        self.samples = []
        self._stop = False
        try:
            self.proc.cpu_percent()
        except Exception:
            pass

        def _poll():
            while not self._stop:
                try:
                    s = self.proc.cpu_percent(interval=self.interval)
                    if s > 0:
                        self.samples.append(s)
                except Exception:
                    break

        self._thread = threading.Thread(target=_poll, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop = True
        if self._thread:
            self._thread.join(timeout=0.3)
        return statistics.mean(self.samples) if self.samples else (self.proc.cpu_percent() or 100.0)


def measure_benchmark(func, repetitions=REPETITIONS, warmup=WARMUP_RUNS):
    # Warmup
    for _ in range(warmup):
        gc.collect()
        func()

    wall_times = []
    cpu_times = []
    cpu_utils = []
    peak_rss_list = []
    proc = psutil.Process()

    for _ in range(repetitions):
        gc.collect()
        time.sleep(0.05)

        rss_before = proc.memory_info().rss
        cpu_t_before = proc.cpu_times()
        cpu_monitor = CPUMonitor(interval=0.01)

        cpu_monitor.start()
        t0 = time.perf_counter()

        func()

        t1 = time.perf_counter()
        avg_cpu_pct = cpu_monitor.stop()
        cpu_t_after = proc.cpu_times()
        rss_after = proc.memory_info().rss

        wall_times.append(t1 - t0)
        user_diff = cpu_t_after.user - cpu_t_before.user
        sys_diff = cpu_t_after.system - cpu_t_before.system
        cpu_times.append(user_diff + sys_diff)
        cpu_utils.append(avg_cpu_pct)
        peak_rss_list.append(max(rss_before, rss_after) / (1024 * 1024))

    return {
        "median_sec": statistics.median(wall_times),
        "mean_sec": statistics.mean(wall_times),
        "min_sec": min(wall_times),
        "max_sec": max(wall_times),
        "stdev_sec": statistics.stdev(wall_times) if len(wall_times) > 1 else 0.0,
        "median_cpu_sec": statistics.median(cpu_times),
        "cpu_util_pct": round(statistics.median(cpu_utils), 1),
        "peak_rss_mb": round(statistics.median(peak_rss_list), 2),
    }


def ensure_scale_file(rows, cols):
    os.makedirs("benchmarks/data", exist_ok=True)
    filepath = f"benchmarks/data/scale_{rows}r_{cols}c.xlsx"
    if os.path.exists(filepath):
        return filepath

    print(f"  Generating test file ({rows:,} rows x {cols} cols) -> {filepath}...")
    import rpyxl

    wb = rpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    headers = [f"Col_{c}" for c in range(1, cols + 1)]
    ws.append(headers)

    dt_str = "2026-05-15 14:30:00"
    for r in range(1, rows + 1):
        row = [
            f"User_{r}",
            r * 10,
            r * 1.5,
            r % 2 == 0,
            dt_str,
            f"Note_{r % 100}",
            r * 100,
            (r * 2.5),
            r % 3 == 0,
            f"Tag_{r % 50}",
        ]
        ws.append(row)

    wb.save(filepath)
    return filepath


# -----------------------------------------------------------------------------
# Benchmark Routines
# -----------------------------------------------------------------------------

def bench_standard_save(lib_name, rows, cols):
    if lib_name == "openpyxl":
        import openpyxl

        def _run():
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=True) as tmp:
                wb = openpyxl.Workbook()
                ws = wb.active
                for r in range(1, rows + 1):
                    ws.append([f"Item_{r}", r * 10, r * 1.5, r % 2 == 0, "Test string"] * (cols // 5))
                wb.save(tmp.name)
        return _run
    else:
        import rpyxl

        def _run():
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=True) as tmp:
                wb = rpyxl.Workbook()
                ws = wb.active
                for r in range(1, rows + 1):
                    ws.append([f"Item_{r}", r * 10, r * 1.5, r % 2 == 0, "Test string"] * (cols // 5))
                wb.save(tmp.name)
        return _run


def bench_write_only_save(lib_name, rows, cols):
    if lib_name == "openpyxl":
        import openpyxl

        def _run():
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=True) as tmp:
                wb = openpyxl.Workbook(write_only=True)
                ws = wb.create_sheet("Stream")
                for r in range(1, rows + 1):
                    ws.append([f"Item_{r}", r * 10, r * 1.5, r % 2 == 0, "Test string"] * (cols // 5))
                wb.save(tmp.name)
        return _run
    else:
        import rpyxl

        def _run():
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=True) as tmp:
                wb = rpyxl.Workbook(write_only=True)
                ws = wb.create_sheet("Stream")
                for r in range(1, rows + 1):
                    ws.append([f"Item_{r}", r * 10, r * 1.5, r % 2 == 0, "Test string"] * (cols // 5))
                wb.save(tmp.name)
        return _run


def bench_standard_load(lib_name, filepath):
    if lib_name == "openpyxl":
        import openpyxl

        def _run():
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            cnt = 0
            for row in ws.iter_rows(values_only=True):
                cnt += len(row)
            wb.close()
            return cnt
        return _run
    elif lib_name == "rpyxl":
        import rpyxl

        def _run():
            wb = rpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            cnt = 0
            for row in ws.iter_rows(values_only=True):
                cnt += len(row)
            wb.close()
            return cnt
        return _run
    elif lib_name == "calamine":
        import python_calamine

        def _run():
            wb = python_calamine.CalamineWorkbook.from_path(filepath)
            sheet = wb.get_sheet_by_name(wb.sheet_names[0])
            cnt = 0
            for row in sheet.to_python():
                cnt += len(row)
            return cnt
        return _run


def bench_read_only_stream(lib_name, filepath):
    if lib_name == "openpyxl":
        import openpyxl

        def _run():
            wb = openpyxl.load_workbook(filepath, read_only=True)
            ws = wb.active
            cnt = 0
            for row in ws.iter_rows(values_only=True):
                cnt += len(row)
            wb.close()
            return cnt
        return _run
    elif lib_name == "rpyxl":
        import rpyxl

        def _run():
            with rpyxl.load_workbook(filepath, read_only=True) as wb:
                ws = wb.active
                cnt = 0
                for row in ws.iter_rows(values_only=True):
                    cnt += len(row)
                return cnt
        return _run


# -----------------------------------------------------------------------------
# Chart Generation
# -----------------------------------------------------------------------------

def generate_scale_charts(results_by_scale, output_dir="benchmarks/reports"):
    os.makedirs(output_dir, exist_ok=True)
    scale_labels = [s["name"] for s in SCALES]

    # 1. Standard Save Wall Time Chart
    plt.figure(figsize=(10, 6))
    op_std_times = [results_by_scale[s["name"]]["std_save_openpyxl"]["median_sec"] for s in SCALES]
    px_std_times = [results_by_scale[s["name"]]["std_save_rpyxl"]["median_sec"] for s in SCALES]

    x = range(len(scale_labels))
    width = 0.35

    plt.bar([i - width/2 for i in x], op_std_times, width, label="openpyxl (Standard)", color="#E53935")
    plt.bar([i + width/2 for i in x], px_std_times, width, label="RPyXL (Standard)", color="#1E88E5")

    plt.yscale("log")
    plt.ylabel("Wall Time in Seconds (Log Scale)", fontsize=11)
    plt.title("Standard In-Memory Save Time Across Scales (Log Scale: Lower is Better)", fontsize=13, fontweight="bold")
    plt.xticks(x, scale_labels, fontsize=10)
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.legend(fontsize=10)

    for i in x:
        sp = op_std_times[i] / px_std_times[i] if px_std_times[i] > 0 else 1.0
        plt.text(i + width/2, px_std_times[i] * 1.15, f"{sp:.1f}x\nfaster", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#0D47A1")

    plt.tight_layout()
    plt.savefig(f"{output_dir}/phase5_scale_writing_time.png", dpi=200)
    plt.close()

    # 2. Write-Only Streaming Time Chart
    plt.figure(figsize=(10, 6))
    op_wo_times = [results_by_scale[s["name"]]["write_only_openpyxl"]["median_sec"] for s in SCALES]
    px_wo_times = [results_by_scale[s["name"]]["write_only_rpyxl"]["median_sec"] for s in SCALES]

    plt.bar([i - width/2 for i in x], op_wo_times, width, label="openpyxl (write_only)", color="#FB8C00")
    plt.bar([i + width/2 for i in x], px_wo_times, width, label="RPyXL (write_only)", color="#43A047")

    plt.yscale("log")
    plt.ylabel("Wall Time in Seconds (Log Scale)", fontsize=11)
    plt.title("Streaming Write-Only Save Time Across Scales (Log Scale: Lower is Better)", fontsize=13, fontweight="bold")
    plt.xticks(x, scale_labels, fontsize=10)
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.legend(fontsize=10)

    for i in x:
        sp = op_wo_times[i] / px_wo_times[i] if px_wo_times[i] > 0 else 1.0
        plt.text(i + width/2, px_wo_times[i] * 1.15, f"{sp:.1f}x\nfaster", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#1B5E20")

    plt.tight_layout()
    plt.savefig(f"{output_dir}/phase5_scale_write_only_time.png", dpi=200)
    plt.close()

    # 3. Reading Time Comparison Chart (Standard & Read-Only)
    plt.figure(figsize=(11, 6))
    op_ro_times = [results_by_scale[s["name"]]["read_ro_openpyxl"]["median_sec"] for s in SCALES]
    px_ro_times = [results_by_scale[s["name"]]["read_ro_rpyxl"]["median_sec"] for s in SCALES]
    cal_times = [results_by_scale[s["name"]]["read_std_calamine"]["median_sec"] for s in SCALES]

    w3 = 0.25
    plt.bar([i - w3 for i in x], op_ro_times, w3, label="openpyxl (read_only)", color="#E53935")
    plt.bar([i for i in x], px_ro_times, w3, label="RPyXL (read_only)", color="#1E88E5")
    plt.bar([i + w3 for i in x], cal_times, w3, label="python-calamine", color="#8E24AA")

    plt.yscale("log")
    plt.ylabel("Wall Time in Seconds (Log Scale)", fontsize=11)
    plt.title("Reading Execution Time Across Scales (Lower is Better)", fontsize=13, fontweight="bold")
    plt.xticks(x, scale_labels, fontsize=10)
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.legend(fontsize=10)

    for i in x:
        sp = op_ro_times[i] / px_ro_times[i] if px_ro_times[i] > 0 else 1.0
        plt.text(i, px_ro_times[i] * 1.15, f"{sp:.1f}x", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#0D47A1")

    plt.tight_layout()
    plt.savefig(f"{output_dir}/phase5_scale_reading_time.png", dpi=200)
    plt.close()

    # 4. Memory Footprint Chart (Peak RSS in MB)
    plt.figure(figsize=(10, 6))
    op_rss = [results_by_scale[s["name"]]["std_save_openpyxl"]["peak_rss_mb"] for s in SCALES]
    px_rss = [results_by_scale[s["name"]]["std_save_rpyxl"]["peak_rss_mb"] for s in SCALES]

    plt.plot(scale_labels, op_rss, marker="o", linewidth=2.5, color="#E53935", label="openpyxl Peak RSS (MB)")
    plt.plot(scale_labels, px_rss, marker="s", linewidth=2.5, color="#1E88E5", label="RPyXL Peak RSS (MB)")

    plt.ylabel("Peak RSS Memory (MB)", fontsize=11)
    plt.title("Peak RAM Usage Scaling (Lower is Better)", fontsize=13, fontweight="bold")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend(fontsize=10)

    for i, s in enumerate(scale_labels):
        plt.annotate(f"{op_rss[i]:.1f} MB", (s, op_rss[i]), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=9, color="#E53935")
        plt.annotate(f"{px_rss[i]:.1f} MB", (s, px_rss[i]), textcoords="offset points", xytext=(0, -14), ha="center", fontsize=9, color="#1E88E5")

    plt.tight_layout()
    plt.savefig(f"{output_dir}/phase5_scale_memory.png", dpi=200)
    plt.close()


def run_all_scales():
    print("================================================================================")
    print("Starting Phase 5 Multi-Scale Comparative Benchmark Suite")
    print("================================================================================")

    env = get_environment_info()
    print(f"OS: {env['os']} ({env['machine']}) | Logical Cores: {env['cpu_count_logical']} | RAM: {env['total_ram_mb']} MB")

    results_by_scale = {}
    summary_rows = []

    for s in SCALES:
        scale_name = s["name"]
        rows = s["rows"]
        cols = s["cols"]
        cells = s["cells"]

        print(f"\n--------------------------------------------------------------------------------")
        print(f">>> Scale: {scale_name} ({rows:,} rows x {cols} cols = {cells:,} cells)")
        print(f"--------------------------------------------------------------------------------")

        # 1. Prepare benchmark XLSX file for reader tests
        bench_file = ensure_scale_file(rows, cols)

        # 2. Standard Save
        print("  Running Standard Save: openpyxl vs RPyXL...")
        m_save_op = measure_benchmark(bench_standard_save("openpyxl", rows, cols))
        m_save_px = measure_benchmark(bench_standard_save("rpyxl", rows, cols))

        # 3. Write-Only Streaming
        print("  Running Write-Only Streaming: openpyxl vs RPyXL...")
        m_wo_op = measure_benchmark(bench_write_only_save("openpyxl", rows, cols))
        m_wo_px = measure_benchmark(bench_write_only_save("rpyxl", rows, cols))

        # 4. Standard Load
        print("  Running Standard Load: openpyxl vs RPyXL vs calamine...")
        m_load_op = measure_benchmark(bench_standard_load("openpyxl", bench_file))
        m_load_px = measure_benchmark(bench_standard_load("rpyxl", bench_file))
        m_load_cal = measure_benchmark(bench_standard_load("calamine", bench_file))

        # 5. Read-Only Streaming
        print("  Running Read-Only Streaming: openpyxl vs RPyXL...")
        m_ro_op = measure_benchmark(bench_read_only_stream("openpyxl", bench_file))
        m_ro_px = measure_benchmark(bench_read_only_stream("rpyxl", bench_file))

        results_by_scale[scale_name] = {
            "scale": s,
            "std_save_openpyxl": m_save_op,
            "std_save_rpyxl": m_save_px,
            "write_only_openpyxl": m_wo_op,
            "write_only_rpyxl": m_wo_px,
            "read_std_openpyxl": m_load_op,
            "read_std_rpyxl": m_load_px,
            "read_std_calamine": m_load_cal,
            "read_ro_openpyxl": m_ro_op,
            "read_ro_rpyxl": m_ro_px,
        }

        # Compute speedups
        save_sp = m_save_op["median_sec"] / m_save_px["median_sec"] if m_save_px["median_sec"] > 0 else 1.0
        wo_sp = m_wo_op["median_sec"] / m_wo_px["median_sec"] if m_wo_px["median_sec"] > 0 else 1.0
        load_sp = m_load_op["median_sec"] / m_load_px["median_sec"] if m_load_px["median_sec"] > 0 else 1.0
        ro_sp = m_ro_op["median_sec"] / m_ro_px["median_sec"] if m_ro_px["median_sec"] > 0 else 1.0

        px_save_throughput = cells / m_save_px["median_sec"] if m_save_px["median_sec"] > 0 else 0
        px_wo_throughput = cells / m_wo_px["median_sec"] if m_wo_px["median_sec"] > 0 else 0
        px_ro_throughput = cells / m_ro_px["median_sec"] if m_ro_px["median_sec"] > 0 else 0

        summary_rows.append({
            "Scale": scale_name,
            "Cells": f"{cells:,}",
            "Save_openpyxl": f"{m_save_op['median_sec']:.4f}s",
            "Save_rpyxl": f"{m_save_px['median_sec']:.4f}s",
            "Save_Speedup": f"{save_sp:.2f}x",
            "Save_Throughput": f"{px_save_throughput:,.0f} cells/s",
            "WO_openpyxl": f"{m_wo_op['median_sec']:.4f}s",
            "WO_rpyxl": f"{m_wo_px['median_sec']:.4f}s",
            "WO_Speedup": f"{wo_sp:.2f}x",
            "WO_Throughput": f"{px_wo_throughput:,.0f} cells/s",
            "Load_openpyxl": f"{m_load_op['median_sec']:.4f}s",
            "Load_rpyxl": f"{m_load_px['median_sec']:.4f}s",
            "Load_Speedup": f"{load_sp:.2f}x",
            "RO_openpyxl": f"{m_ro_op['median_sec']:.4f}s",
            "RO_rpyxl": f"{m_ro_px['median_sec']:.4f}s",
            "RO_Speedup": f"{ro_sp:.2f}x",
            "RO_Throughput": f"{px_ro_throughput:,.0f} cells/s",
            "RAM_openpyxl": f"{m_save_op['peak_rss_mb']:.1f} MB",
            "RAM_rpyxl": f"{m_save_px['peak_rss_mb']:.1f} MB",
            "CPU_util_rpyxl": f"{m_save_px['cpu_util_pct']:.0f}%",
        })

    # Generate charts
    print("\nGenerating Phase 5 comparison charts...")
    generate_scale_charts(results_by_scale)

    # Save Markdown report
    os.makedirs("benchmarks/reports", exist_ok=True)
    md_path = "benchmarks/reports/phase5_multiscale_report.md"
    with open(md_path, "w") as f:
        f.write("# RPyXL Phase 5 Multi-Scale Benchmark Report\n\n")
        f.write(f"Generated at: {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')}\n\n")
        f.write(f"**Environment**: {env['os']} ({env['machine']}), {env['cpu_count_logical']} Logical Cores, {env['total_ram_mb']} MB RAM\n\n")

        f.write("## 1. Multi-Scale Execution Summary\n\n")
        f.write("| Scale | Cells | Std Save openpyxl | Std Save RPyXL | Save Speedup | Write-Only openpyxl | Write-Only RPyXL | WO Speedup | Read openpyxl | Read RPyXL | Read Speedup | RPyXL Peak RAM | openpyxl Peak RAM |\n")
        f.write("|---|---|---|---|---|---|---|---|---|---|---|---|---|\n")
        for r in summary_rows:
            f.write(f"| {r['Scale']} | {r['Cells']} | {r['Save_openpyxl']} | {r['Save_rpyxl']} | **{r['Save_Speedup']}** | {r['WO_openpyxl']} | {r['WO_rpyxl']} | **{r['WO_Speedup']}** | {r['RO_openpyxl']} | {r['RO_rpyxl']} | **{r['RO_Speedup']}** | {r['RAM_rpyxl']} | {r['RAM_openpyxl']} |\n")

        f.write("\n## 2. Detailed Throughput & Performance\n\n")
        f.write("| Scale | Standard Save Throughput | Write-Only Throughput | Read-Only Streaming Throughput | RPyXL CPU Utilization |\n")
        f.write("|---|---|---|---|---|\n")
        for r in summary_rows:
            f.write(f"| {r['Scale']} | **{r['Save_Throughput']}** | **{r['WO_Throughput']}** | **{r['RO_Throughput']}** | {r['CPU_util_rpyxl']} |\n")

        f.write("\n## 3. Visual Performance Comparisons\n\n")
        f.write("### 3.1 Standard In-Memory Save Time Across Scales\n")
        f.write("![Standard Save Time](phase5_scale_writing_time.png)\n\n")
        f.write("### 3.2 Streaming Write-Only Time Across Scales\n")
        f.write("![Write-Only Time](phase5_scale_write_only_time.png)\n\n")
        f.write("### 3.3 Reading Time Across Scales\n")
        f.write("![Reading Time](phase5_scale_reading_time.png)\n\n")
        f.write("### 3.4 Peak Memory Footprint (RAM)\n")
        f.write("![Memory Scaling](phase5_scale_memory.png)\n\n")

    # Save JSON report
    with open("benchmarks/reports/phase5_multiscale_results.json", "w") as f:
        json.dump({"environment": env, "results": results_by_scale, "summary": summary_rows}, f, indent=2)

    print("\n================================================================================")
    print(f"Phase 5 Multi-Scale Benchmarks Completed Successfully!")
    print(f"Report: {md_path}")
    print("Charts: benchmarks/reports/phase5_scale_*.png")
    print("JSON:   benchmarks/reports/phase5_multiscale_results.json")
    print("================================================================================")


if __name__ == "__main__":
    run_all_scales()
