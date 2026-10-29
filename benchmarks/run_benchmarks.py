"""Comprehensive benchmark runner for RPyXL vs openpyxl (Phase 1, Phase 2, and Phase 3).

Measures:
- Phase 1: In-Memory Operations (Workbook creation, cell access, append, iteration)
- Phase 2: XLSX File Reading (Standard loading, streaming read_only, time to first row)
- Phase 3: High-Performance XLSX Writing (Standard save, streaming write-only, 10k & 50k rows, round-trip)
- Baselines: openpyxl (standard & write_only), RPyXL (standard & write_only), python-calamine
- Metrics: Wall-clock time (median, mean, min, max, stdev), CPU time (user + sys), CPU utilization (%), Peak RSS, Throughput, Speedup
- Generates visual plots (PNG) and Markdown/JSON reports
"""

import csv
import gc
import json
import os
import platform
import statistics
import sys
import threading
import time
from datetime import datetime
import psutil
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

REPETITIONS = 5
WARMUP_RUNS = 1


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
    """Samples CPU utilization percentage during a benchmark run."""
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


def ensure_benchmark_file(rows=10000, cols=10):
    os.makedirs("benchmarks/data", exist_ok=True)
    filepath = f"benchmarks/data/bench_{rows // 1000}k.xlsx"
    if os.path.exists(filepath):
        return filepath

    print(f"Generating benchmark file with {rows:,} rows x {cols} cols: {filepath}...")
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DataSheet"

    headers = [f"Col_{c}" for c in range(1, cols + 1)]
    ws.append(headers)

    for r in range(1, rows + 1):
        row = [
            f"User_{r}",
            r * 10,
            r * 1.25,
            r % 2 == 0,
            datetime(2023, 1, 1, 12, 0, 0),
            f"Item_{r % 100}",
            r * 5,
            round(r * 0.05, 2),
            r % 3 == 0,
            "Active",
        ]
        ws.append(row)

    wb.save(filepath)
    print(f"Generated {filepath} ({round(os.path.getsize(filepath) / 1024, 1)} KB)")
    return filepath


# ==============================================================================
# Phase 1 Benchmarks (In-Memory)
# ==============================================================================

def bench_create_workbook(lib_name, iterations=200):
    mod = __import__(lib_name)
    Workbook = mod.Workbook

    for _ in range(WARMUP_RUNS):
        wb = Workbook()
        wb.close()

    times = []
    cpu_times = []
    mem_before = psutil.Process().memory_info().rss

    for _ in range(REPETITIONS):
        gc.collect()
        t0 = time.perf_counter()
        c0 = time.process_time()

        for _ in range(iterations):
            wb = Workbook()
            wb.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    mem_after = psutil.Process().memory_info().rss
    peak_rss = max(0, mem_after - mem_before)

    return {
        "operation": "Workbook Creation",
        "dataset": f"{iterations} instances",
        "total_units": iterations,
        "unit_name": "workbooks",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_cell_access_write(lib_name, rows=100, cols=100):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols

    wb = Workbook()
    ws = wb.active
    times = []
    cpu_times = []
    proc = psutil.Process()

    for _ in range(REPETITIONS):
        gc.collect()
        t0 = time.perf_counter()
        c0 = time.process_time()

        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                ws.cell(row=r, column=c, value=r * c)

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    peak_rss = proc.memory_info().rss
    wb.close()

    return {
        "operation": "Cell Access (Write)",
        "dataset": f"{rows}x{cols} ({total_cells:,} cells)",
        "total_units": total_cells,
        "unit_name": "cells",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_append_rows(lib_name, rows=5000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols

    sample_row = [100.5, "Alpha", True, 42, None, "Beta", 99.9, False, 12, "Gamma"]
    times = []
    cpu_times = []
    proc = psutil.Process()

    for _ in range(REPETITIONS):
        gc.collect()
        wb = Workbook()
        ws = wb.active

        t0 = time.perf_counter()
        c0 = time.process_time()

        for _ in range(rows):
            ws.append(sample_row)

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)
        wb.close()

    peak_rss = proc.memory_info().rss

    return {
        "operation": "Row Appending",
        "dataset": f"{rows:,} rows ({total_cells:,} cells)",
        "total_units": rows,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_iteration_values(lib_name, rows=5000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols

    wb = Workbook()
    ws = wb.active
    sample_row = [100.5, "Alpha", True, 42, None, "Beta", 99.9, False, 12, "Gamma"]
    for _ in range(rows):
        ws.append(sample_row)

    times = []
    cpu_times = []
    proc = psutil.Process()

    for _ in range(REPETITIONS):
        gc.collect()
        t0 = time.perf_counter()
        c0 = time.process_time()

        count = 0
        for row in ws.iter_rows(values_only=True):
            count += len(row)

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    peak_rss = proc.memory_info().rss
    wb.close()

    return {
        "operation": "In-Memory Row Iteration",
        "dataset": f"{rows:,} rows ({total_cells:,} cells)",
        "total_units": total_cells,
        "unit_name": "cells",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


# ==============================================================================
# Phase 2 Benchmarks (XLSX Reader)
# ==============================================================================

def bench_load_full_workbook(lib_name, filepath, read_only=False):
    mod = __import__(lib_name)
    load_workbook = mod.load_workbook

    for _ in range(WARMUP_RUNS):
        wb = load_workbook(filepath, read_only=read_only)
        wb.close()

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = load_workbook(filepath, read_only=read_only)
        count = 0
        for row in wb.active.iter_rows(values_only=True):
            count += len(row)
        wb.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    mode = "read_only=True" if read_only else "Standard"
    return {
        "operation": f"Full Read ({mode})",
        "dataset": "10,000 rows x 10 cols",
        "total_units": 10000,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_calamine_full_read(filepath):
    from python_calamine import CalamineWorkbook

    for _ in range(WARMUP_RUNS):
        wb = CalamineWorkbook.from_path(filepath)
        sheet_name = wb.sheet_names[0]
        _ = wb.get_sheet_by_name(sheet_name).to_python()

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = CalamineWorkbook.from_path(filepath)
        sheet_name = wb.sheet_names[0]
        rows = wb.get_sheet_by_name(sheet_name).to_python()
        count = len(rows)

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    return {
        "operation": "Calamine Full Read",
        "dataset": "10,000 rows x 10 cols",
        "total_units": 10000,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_time_to_first_row(lib_name, filepath, read_only=True):
    mod = __import__(lib_name)
    load_workbook = mod.load_workbook

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = load_workbook(filepath, read_only=read_only)
        row1 = next(iter(wb.active.iter_rows(values_only=True)))
        wb.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    mode = "read_only=True" if read_only else "Standard"
    return {
        "operation": f"Time to First Row ({mode})",
        "dataset": "10,000 rows x 10 cols",
        "total_units": 1,
        "unit_name": "first_row",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


# ==============================================================================
# Phase 3 Benchmarks (High-Performance XLSX Writer)
# ==============================================================================

def bench_standard_write(lib_name, rows=10000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    tmp_out = f"/tmp/bench_std_save_{lib_name}.xlsx"
    sample_row = ["User_X", 100, 3.1415, True, datetime(2023, 1, 1), "Item_A", 50, 0.05, False, "Active"]

    def build_wb():
        wb = Workbook()
        ws = wb.active
        for _ in range(rows):
            ws.append(sample_row)
        return wb

    # Warmup
    for _ in range(WARMUP_RUNS):
        wb = build_wb()
        wb.save(tmp_out)
        wb.close()
        if os.path.exists(tmp_out):
            os.remove(tmp_out)

    times = []
    cpu_times = []
    cpu_utils = []
    peak_rss = 0
    proc = psutil.Process()
    monitor = CPUMonitor(interval=0.01)

    for _ in range(REPETITIONS):
        wb = build_wb()
        gc.collect()
        monitor.start()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb.save(tmp_out)

        c1 = time.process_time()
        t1 = time.perf_counter()
        cpu_util = monitor.stop()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)
        wb.close()

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)
        cpu_utils.append(cpu_util)

        if os.path.exists(tmp_out):
            os.remove(tmp_out)

    return {
        "operation": "Standard XLSX Save",
        "dataset": f"{rows:,} rows x {cols} cols",
        "total_units": rows,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "cpu_util_pct": round(statistics.median(cpu_utils), 1) if cpu_utils else 100.0,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_write_only_streaming(lib_name, rows=10000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    tmp_out = f"/tmp/bench_wo_{lib_name}_{rows}.xlsx"
    sample_row = ["User_X", 100, 3.1415, True, datetime(2023, 1, 1), "Item_A", 50, 0.05, False, "Active"]

    # Warmup
    for _ in range(WARMUP_RUNS):
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Data")
        for _ in range(min(rows, 1000)):
            ws.append(sample_row)
        wb.save(tmp_out)
        wb.close()
        if os.path.exists(tmp_out):
            os.remove(tmp_out)

    times = []
    cpu_times = []
    cpu_utils = []
    peak_rss = 0
    proc = psutil.Process()
    monitor = CPUMonitor(interval=0.01)

    reps = REPETITIONS if rows <= 10000 else 3
    for _ in range(reps):
        gc.collect()
        monitor.start()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Data")
        for _ in range(rows):
            ws.append(sample_row)
        wb.save(tmp_out)

        c1 = time.process_time()
        t1 = time.perf_counter()
        cpu_util = monitor.stop()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)
        wb.close()

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)
        cpu_utils.append(cpu_util)

        if os.path.exists(tmp_out):
            os.remove(tmp_out)

    return {
        "operation": f"Write-Only Streaming ({rows:,} rows)",
        "dataset": f"{rows:,} rows x {cols} cols",
        "total_units": rows,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "cpu_util_pct": round(statistics.median(cpu_utils), 1) if cpu_utils else 100.0,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_roundtrip(lib_name, rows=5000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    load_workbook = mod.load_workbook
    tmp_out = f"/tmp/bench_rt_{lib_name}.xlsx"
    sample_row = ["User_X", 100, 3.1415, True, datetime(2023, 1, 1), "Item_A", 50, 0.05, False, "Active"]

    times = []
    cpu_times = []
    cpu_utils = []
    peak_rss = 0
    proc = psutil.Process()
    monitor = CPUMonitor(interval=0.01)

    for _ in range(3):
        gc.collect()
        monitor.start()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = Workbook()
        ws = wb.active
        for _ in range(rows):
            ws.append(sample_row)
        wb.save(tmp_out)
        wb.close()

        wb_in = load_workbook(tmp_out)
        count = sum(len(r) for r in wb_in.active.iter_rows(values_only=True))
        wb_in.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        cpu_util = monitor.stop()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)
        cpu_utils.append(cpu_util)

        if os.path.exists(tmp_out):
            os.remove(tmp_out)

    return {
        "operation": f"Roundtrip Write+Read ({rows:,} rows)",
        "dataset": f"{rows:,} rows x {cols} cols",
        "total_units": rows,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "cpu_util_pct": round(statistics.median(cpu_utils), 1) if cpu_utils else 100.0,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def compute_metrics(bench_dict):
    times = bench_dict["times"]
    cpu_times = bench_dict["cpu_times"]
    total_units = bench_dict["total_units"]

    median_t = statistics.median(times)
    mean_t = statistics.mean(times)
    min_t = min(times)
    max_t = max(times)
    stdev_t = statistics.stdev(times) if len(times) > 1 else 0.0

    median_cpu = statistics.median(cpu_times)
    throughput = total_units / median_t if median_t > 0 else 0.0

    return {
        "operation": bench_dict["operation"],
        "dataset": bench_dict["dataset"],
        "unit_name": bench_dict["unit_name"],
        "median_sec": median_t,
        "mean_sec": mean_t,
        "min_sec": min_t,
        "max_sec": max_t,
        "stdev_sec": stdev_t,
        "median_cpu_sec": median_cpu,
        "cpu_util_pct": bench_dict.get("cpu_util_pct", 100.0),
        "peak_rss_mb": bench_dict["peak_rss_mb"],
        "throughput": throughput,
    }


def generate_phase2_plots(p2_data):
    os.makedirs("benchmarks/reports", exist_ok=True)

    ops = [item["name"] for item in p2_data]
    times = [item["time_s"] for item in p2_data]
    colors = ["#e74c3c", "#f39c12", "#2ecc71", "#1abc9c", "#3498db"]

    # 1. Reading Time
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(ops, times, color=colors[:len(ops)], width=0.5)
    ax.set_ylabel("Execution Time (seconds) - Lower is better")
    ax.set_title("XLSX Read Time (10,000 rows x 10 columns)")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.4f}s",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase2_reading_time.png", dpi=150)
    plt.close()

    # 2. Reading Speedup
    openpyxl_std_time = next(item["time_s"] for item in p2_data if "openpyxl (Standard)" in item["name"])
    speedup_items = []
    for item in p2_data:
        speedup = openpyxl_std_time / item["time_s"] if item["time_s"] > 0 else 1.0
        speedup_items.append((item["name"], speedup))

    fig, ax = plt.subplots(figsize=(10, 5))
    s_names = [x[0] for x in speedup_items]
    s_vals = [x[1] for x in speedup_items]
    bars = ax.bar(s_names, s_vals, color="#2980b9", width=0.5)
    ax.axhline(1.0, color="gray", linestyle="--", linewidth=1, label="Baseline (openpyxl std = 1.0x)")
    ax.set_ylabel("Speedup Multiplier vs openpyxl (Higher is better)")
    ax.set_title("XLSX Reading Speedup Multiplier vs openpyxl Standard")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.1f}x",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase2_speedup.png", dpi=150)
    plt.close()

    # 3. Peak RSS Memory
    mems = [item["rss_mb"] for item in p2_data]
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(ops, mems, color=["#e67e22", "#d35400", "#27ae60", "#16a085", "#2980b9"][:len(ops)], width=0.5)
    ax.set_ylabel("Peak RSS Memory (MB) - Lower is better")
    ax.set_title("Peak Memory Usage During XLSX Read (10,000 rows)")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.1f} MB",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase2_memory.png", dpi=150)
    plt.close()


def generate_phase3_plots(p3_data):
    os.makedirs("benchmarks/reports", exist_ok=True)

    ops = [item["name"] for item in p3_data]
    times = [item["time_s"] for item in p3_data]
    colors = ["#e74c3c", "#2ecc71", "#e67e22", "#27ae60", "#c0392b", "#16a085", "#8e44ad", "#2980b9"]

    # 1. Writing Execution Time
    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(ops, times, color=colors[:len(ops)], width=0.55)
    ax.set_ylabel("Execution Time (seconds) - Lower is better")
    ax.set_title("XLSX Writing Execution Time: RPyXL vs openpyxl")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.3f}s",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=20, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase3_writing_time.png", dpi=150)
    plt.close()

    # 2. Writing Speedup Multipliers
    fig, ax = plt.subplots(figsize=(10, 5))
    speedup_ops = []
    speedups = []
    for pair in [("Standard Save (10k)", "openpyxl (Std Save 10k)", "RPyXL (Std Save 10k)"),
                 ("Write-Only Streaming (10k)", "openpyxl (Write-Only 10k)", "RPyXL (Write-Only 10k)"),
                 ("Write-Only Large (50k)", "openpyxl (Write-Only 50k)", "RPyXL (Write-Only 50k)"),
                 ("Roundtrip (5k)", "openpyxl (Roundtrip 5k)", "RPyXL (Roundtrip 5k)")]:
        label, op_name, px_name = pair
        op_item = next((x for x in p3_data if x["name"] == op_name), None)
        px_item = next((x for x in p3_data if x["name"] == px_name), None)
        if op_item and px_item and px_item["time_s"] > 0:
            sp = op_item["time_s"] / px_item["time_s"]
            speedup_ops.append(label)
            speedups.append(sp)

    bars = ax.bar(speedup_ops, speedups, color="#27ae60", width=0.45)
    ax.axhline(1.0, color="gray", linestyle="--", linewidth=1, label="Baseline (openpyxl = 1.0x)")
    ax.set_ylabel("Speedup Multiplier vs openpyxl (Higher is better)")
    ax.set_title("RPyXL Writing Speedup over openpyxl")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.1f}x",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=10, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase3_speedup.png", dpi=150)
    plt.close()

    # 3. Peak RSS Memory
    mems = [item["rss_mb"] for item in p3_data]
    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(ops, mems, color=colors[:len(ops)], width=0.55)
    ax.set_ylabel("Peak RSS Memory (MB) - Lower is better")
    ax.set_title("Memory Footprint During XLSX Writing (Standard & Streaming)")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.1f} MB",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=20, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase3_memory.png", dpi=150)
    plt.close()

    # 4. CPU Time & Utilization
    cpu_secs = [item["cpu_s"] for item in p3_data]
    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(ops, cpu_secs, color="#8e44ad", width=0.55)
    ax.set_ylabel("Total CPU Time (User + System Seconds) - Lower is better")
    ax.set_title("CPU Utilization & Total CPU Consumption During Writing")
    for bar, item in zip(bars, p3_data):
        h = bar.get_height()
        pct = item.get("cpu_pct", 100.0)
        ax.annotate(f"{h:.3f}s\n({pct:.0f}% CPU)",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=8, fontweight="bold")
    plt.xticks(rotation=20, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase3_cpu_utilization.png", dpi=150)
    plt.close()


def run_all():
    os.makedirs("benchmarks/reports", exist_ok=True)
    os.makedirs("docs/reports", exist_ok=True)
    bench_file = ensure_benchmark_file(rows=10000, cols=10)

    print("================================================================================")
    print("RPyXL Comprehensive Benchmark Suite (Phase 1, Phase 2, and Phase 3)")
    print("================================================================================")

    # -------------------------------------------------------------------------
    # Part 1: Phase 1 In-Memory Operations
    # -------------------------------------------------------------------------
    print("\n>>> Running Phase 1 In-Memory Operations...")
    p1_benchmarks = [
        ("Workbook Creation", lambda lib: bench_create_workbook(lib, iterations=100)),
        ("Cell Access (Write)", lambda lib: bench_cell_access_write(lib, rows=100, cols=100)),
        ("Row Appending", lambda lib: bench_append_rows(lib, rows=5000, cols=10)),
        ("Row Iteration (Values)", lambda lib: bench_iteration_values(lib, rows=5000, cols=10)),
    ]

    p1_results = []
    for name, fn in p1_benchmarks:
        print(f"  Measuring {name}...")
        rpyxl_m = compute_metrics(fn("rpyxl"))
        openpyxl_m = compute_metrics(fn("openpyxl"))

        speedup = openpyxl_m["median_sec"] / rpyxl_m["median_sec"] if rpyxl_m["median_sec"] > 0 else 1.0
        mem_diff_pct = (
            (1.0 - (rpyxl_m["peak_rss_mb"] / openpyxl_m["peak_rss_mb"])) * 100
            if openpyxl_m["peak_rss_mb"] > 0
            else 0.0
        )

        p1_results.append({
            "Operation": rpyxl_m["operation"],
            "Dataset": rpyxl_m["dataset"],
            "openpyxl_time_s": round(openpyxl_m["median_sec"], 4),
            "rpyxl_time_s": round(rpyxl_m["median_sec"], 4),
            "speedup": round(speedup, 2),
            "openpyxl_rss_mb": openpyxl_m["peak_rss_mb"],
            "rpyxl_rss_mb": rpyxl_m["peak_rss_mb"],
            "mem_reduction_pct": round(mem_diff_pct, 1),
            "rpyxl_throughput": f"{round(rpyxl_m['throughput'], 1):,} {rpyxl_m['unit_name']}/s",
            "openpyxl_throughput": f"{round(openpyxl_m['throughput'], 1):,} {openpyxl_m['unit_name']}/s",
        })

    # -------------------------------------------------------------------------
    # Part 2: Phase 2 XLSX Reader Operations
    # -------------------------------------------------------------------------
    print("\n>>> Running Phase 2 XLSX Reader Benchmarks (10,000 rows x 10 cols)...")
    m_op_std = compute_metrics(bench_load_full_workbook("openpyxl", bench_file, read_only=False))
    m_op_ro = compute_metrics(bench_load_full_workbook("openpyxl", bench_file, read_only=True))
    m_px_std = compute_metrics(bench_load_full_workbook("rpyxl", bench_file, read_only=False))
    m_px_ro = compute_metrics(bench_load_full_workbook("rpyxl", bench_file, read_only=True))
    m_cal = compute_metrics(bench_calamine_full_read(bench_file))

    p2_chart_data = [
        {"name": "openpyxl (Standard)", "time_s": round(m_op_std["median_sec"], 4), "rss_mb": m_op_std["peak_rss_mb"]},
        {"name": "openpyxl (read_only)", "time_s": round(m_op_ro["median_sec"], 4), "rss_mb": m_op_ro["peak_rss_mb"]},
        {"name": "RPyXL (Standard)", "time_s": round(m_px_std["median_sec"], 4), "rss_mb": m_px_std["peak_rss_mb"]},
        {"name": "RPyXL (read_only)", "time_s": round(m_px_ro["median_sec"], 4), "rss_mb": m_px_ro["peak_rss_mb"]},
        {"name": "calamine (C-Rust)", "time_s": round(m_cal["median_sec"], 4), "rss_mb": m_cal["peak_rss_mb"]},
    ]

    ttfr_op_std = compute_metrics(bench_time_to_first_row("openpyxl", bench_file, read_only=False))
    ttfr_op_ro = compute_metrics(bench_time_to_first_row("openpyxl", bench_file, read_only=True))
    ttfr_px_std = compute_metrics(bench_time_to_first_row("rpyxl", bench_file, read_only=False))
    ttfr_px_ro = compute_metrics(bench_time_to_first_row("rpyxl", bench_file, read_only=True))

    p2_results = [
        {
            "Operation": "Full XLSX Read (Standard)",
            "openpyxl": f"{m_op_std['median_sec']:.4f}s",
            "rpyxl": f"{m_px_std['median_sec']:.4f}s",
            "speedup": f"{m_op_std['median_sec'] / m_px_std['median_sec']:.2f}x",
            "openpyxl_rss": f"{m_op_std['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_px_std['peak_rss_mb']:.1f} MB",
            "throughput": f"{m_px_std['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Full XLSX Read (read_only=True)",
            "openpyxl": f"{m_op_ro['median_sec']:.4f}s",
            "rpyxl": f"{m_px_ro['median_sec']:.4f}s",
            "speedup": f"{m_op_ro['median_sec'] / m_px_ro['median_sec']:.2f}x",
            "openpyxl_rss": f"{m_op_ro['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_px_ro['peak_rss_mb']:.1f} MB",
            "throughput": f"{m_px_ro['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Time to First Row (read_only=True)",
            "openpyxl": f"{ttfr_op_ro['median_sec']*1000:.2f} ms",
            "rpyxl": f"{ttfr_px_ro['median_sec']*1000:.2f} ms",
            "speedup": f"{ttfr_op_ro['median_sec'] / ttfr_px_ro['median_sec']:.2f}x",
            "openpyxl_rss": f"{ttfr_op_ro['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{ttfr_px_ro['peak_rss_mb']:.1f} MB",
            "throughput": "Instantaneous",
        },
        {
            "Operation": "Time to First Row (Standard vs RO)",
            "openpyxl": f"{ttfr_op_std['median_sec']*1000:.2f} ms",
            "rpyxl": f"{ttfr_px_ro['median_sec']*1000:.2f} ms",
            "speedup": f"{ttfr_op_std['median_sec'] / ttfr_px_ro['median_sec']:.1f}x",
            "openpyxl_rss": f"{ttfr_op_std['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{ttfr_px_ro['peak_rss_mb']:.1f} MB",
            "throughput": "Instantaneous",
        },
        {
            "Operation": "Calamine Full Read Baseline",
            "openpyxl": f"{m_op_std['median_sec']:.4f}s (openpyxl std)",
            "rpyxl": f"{m_cal['median_sec']:.4f}s (calamine)",
            "speedup": f"{m_op_std['median_sec'] / m_cal['median_sec']:.2f}x",
            "openpyxl_rss": f"{m_op_std['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_cal['peak_rss_mb']:.1f} MB",
            "throughput": f"{m_cal['throughput']:,.0f} rows/s",
        },
    ]
    generate_phase2_plots(p2_chart_data)

    # -------------------------------------------------------------------------
    # Part 3: Phase 3 XLSX Writer Operations
    # -------------------------------------------------------------------------
    print("\n>>> Running Phase 3 High-Performance XLSX Writer Benchmarks...")

    print("  Measuring Standard Save (10,000 rows x 10 cols)...")
    m_w_std_op = compute_metrics(bench_standard_write("openpyxl", rows=10000, cols=10))
    m_w_std_px = compute_metrics(bench_standard_write("rpyxl", rows=10000, cols=10))

    print("  Measuring Write-Only Streaming (10,000 rows x 10 cols)...")
    m_wo_10k_op = compute_metrics(bench_write_only_streaming("openpyxl", rows=10000, cols=10))
    m_wo_10k_px = compute_metrics(bench_write_only_streaming("rpyxl", rows=10000, cols=10))

    print("  Measuring Write-Only Streaming Large Dataset (50,000 rows x 10 cols)...")
    m_wo_50k_op = compute_metrics(bench_write_only_streaming("openpyxl", rows=50000, cols=10))
    m_wo_50k_px = compute_metrics(bench_write_only_streaming("rpyxl", rows=50000, cols=10))

    print("  Measuring End-to-End Roundtrip (Write + Read 5,000 rows)...")
    m_rt_op = compute_metrics(bench_roundtrip("openpyxl", rows=5000, cols=10))
    m_rt_px = compute_metrics(bench_roundtrip("rpyxl", rows=5000, cols=10))

    p3_chart_data = [
        {"name": "openpyxl (Std Save 10k)", "time_s": round(m_w_std_op["median_sec"], 4), "cpu_s": round(m_w_std_op["median_cpu_sec"], 4), "cpu_pct": m_w_std_op["cpu_util_pct"], "rss_mb": m_w_std_op["peak_rss_mb"]},
        {"name": "RPyXL (Std Save 10k)", "time_s": round(m_w_std_px["median_sec"], 4), "cpu_s": round(m_w_std_px["median_cpu_sec"], 4), "cpu_pct": m_w_std_px["cpu_util_pct"], "rss_mb": m_w_std_px["peak_rss_mb"]},
        {"name": "openpyxl (Write-Only 10k)", "time_s": round(m_wo_10k_op["median_sec"], 4), "cpu_s": round(m_wo_10k_op["median_cpu_sec"], 4), "cpu_pct": m_wo_10k_op["cpu_util_pct"], "rss_mb": m_wo_10k_op["peak_rss_mb"]},
        {"name": "RPyXL (Write-Only 10k)", "time_s": round(m_wo_10k_px["median_sec"], 4), "cpu_s": round(m_wo_10k_px["median_cpu_sec"], 4), "cpu_pct": m_wo_10k_px["cpu_util_pct"], "rss_mb": m_wo_10k_px["peak_rss_mb"]},
        {"name": "openpyxl (Write-Only 50k)", "time_s": round(m_wo_50k_op["median_sec"], 4), "cpu_s": round(m_wo_50k_op["median_cpu_sec"], 4), "cpu_pct": m_wo_50k_op["cpu_util_pct"], "rss_mb": m_wo_50k_op["peak_rss_mb"]},
        {"name": "RPyXL (Write-Only 50k)", "time_s": round(m_wo_50k_px["median_sec"], 4), "cpu_s": round(m_wo_50k_px["median_cpu_sec"], 4), "cpu_pct": m_wo_50k_px["cpu_util_pct"], "rss_mb": m_wo_50k_px["peak_rss_mb"]},
        {"name": "openpyxl (Roundtrip 5k)", "time_s": round(m_rt_op["median_sec"], 4), "cpu_s": round(m_rt_op["median_cpu_sec"], 4), "cpu_pct": m_rt_op["cpu_util_pct"], "rss_mb": m_rt_op["peak_rss_mb"]},
        {"name": "RPyXL (Roundtrip 5k)", "time_s": round(m_rt_px["median_sec"], 4), "cpu_s": round(m_rt_px["median_cpu_sec"], 4), "cpu_pct": m_rt_px["cpu_util_pct"], "rss_mb": m_rt_px["peak_rss_mb"]},
    ]

    p3_results = [
        {
            "Operation": "Standard Save (10,000 rows x 10 cols)",
            "openpyxl_time": f"{m_w_std_op['median_sec']:.4f}s",
            "rpyxl_time": f"{m_w_std_px['median_sec']:.4f}s",
            "speedup": f"{m_w_std_op['median_sec'] / m_w_std_px['median_sec']:.2f}x",
            "openpyxl_cpu": f"{m_w_std_op['median_cpu_sec']:.3f}s ({m_w_std_op['cpu_util_pct']:.0f}%)",
            "rpyxl_cpu": f"{m_w_std_px['median_cpu_sec']:.3f}s ({m_w_std_px['cpu_util_pct']:.0f}%)",
            "openpyxl_rss": f"{m_w_std_op['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_w_std_px['peak_rss_mb']:.1f} MB",
            "rpyxl_throughput": f"{m_w_std_px['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Write-Only Streaming (10,000 rows x 10 cols)",
            "openpyxl_time": f"{m_wo_10k_op['median_sec']:.4f}s",
            "rpyxl_time": f"{m_wo_10k_px['median_sec']:.4f}s",
            "speedup": f"{m_wo_10k_op['median_sec'] / m_wo_10k_px['median_sec']:.2f}x",
            "openpyxl_cpu": f"{m_wo_10k_op['median_cpu_sec']:.3f}s ({m_wo_10k_op['cpu_util_pct']:.0f}%)",
            "rpyxl_cpu": f"{m_wo_10k_px['median_cpu_sec']:.3f}s ({m_wo_10k_px['cpu_util_pct']:.0f}%)",
            "openpyxl_rss": f"{m_wo_10k_op['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_wo_10k_px['peak_rss_mb']:.1f} MB",
            "rpyxl_throughput": f"{m_wo_10k_px['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Write-Only Streaming (50,000 rows x 10 cols)",
            "openpyxl_time": f"{m_wo_50k_op['median_sec']:.4f}s",
            "rpyxl_time": f"{m_wo_50k_px['median_sec']:.4f}s",
            "speedup": f"{m_wo_50k_op['median_sec'] / m_wo_50k_px['median_sec']:.2f}x",
            "openpyxl_cpu": f"{m_wo_50k_op['median_cpu_sec']:.3f}s ({m_wo_50k_op['cpu_util_pct']:.0f}%)",
            "rpyxl_cpu": f"{m_wo_50k_px['median_cpu_sec']:.3f}s ({m_wo_50k_px['cpu_util_pct']:.0f}%)",
            "openpyxl_rss": f"{m_wo_50k_op['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_wo_50k_px['peak_rss_mb']:.1f} MB",
            "rpyxl_throughput": f"{m_wo_50k_px['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Roundtrip Write + Read (5,000 rows x 10 cols)",
            "openpyxl_time": f"{m_rt_op['median_sec']:.4f}s",
            "rpyxl_time": f"{m_rt_px['median_sec']:.4f}s",
            "speedup": f"{m_rt_op['median_sec'] / m_rt_px['median_sec']:.2f}x",
            "openpyxl_cpu": f"{m_rt_op['median_cpu_sec']:.3f}s ({m_rt_op['cpu_util_pct']:.0f}%)",
            "rpyxl_cpu": f"{m_rt_px['median_cpu_sec']:.3f}s ({m_rt_px['cpu_util_pct']:.0f}%)",
            "openpyxl_rss": f"{m_rt_op['peak_rss_mb']:.1f} MB",
            "rpyxl_rss": f"{m_rt_px['peak_rss_mb']:.1f} MB",
            "rpyxl_throughput": f"{m_rt_px['throughput']:,.0f} rows/s",
        },
    ]

    generate_phase3_plots(p3_chart_data)

    # -------------------------------------------------------------------------
    # Save Reports
    # -------------------------------------------------------------------------
    md_file = "benchmarks/reports/comparison.md"
    env = get_environment_info()
    with open(md_file, "w") as f:
        f.write("# RPyXL Comprehensive Benchmark Report (Phases 1, 2, and 3)\n\n")
        f.write(f"Generated at: {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')}\n\n")
        f.write(f"**Environment**: {env['os']} ({env['machine']}), {env['cpu_count_logical']} Logical Cores, {env['total_ram_mb']} MB RAM\n\n")

        f.write("## 1. Phase 3: High-Performance XLSX Writer & CPU Utilization Benchmarks\n\n")
        f.write("| Operation | openpyxl (Wall) | RPyXL (Wall) | Speedup | openpyxl CPU Time | RPyXL CPU Time | openpyxl RSS | RPyXL RSS | RPyXL Throughput |\n")
        f.write("|---|---|---|---|---|---|---|---|---|\n")
        for r in p3_results:
            f.write(f"| {r['Operation']} | {r['openpyxl_time']} | {r['rpyxl_time']} | **{r['speedup']}** | {r['openpyxl_cpu']} | {r['rpyxl_cpu']} | {r['openpyxl_rss']} | {r['rpyxl_rss']} | {r['rpyxl_throughput']} |\n")

        f.write("\n### Phase 3 Performance Graphs\n\n")
        f.write("#### 1. Writing Execution Time (Wall-Clock)\n")
        f.write("![Writing Time](phase3_writing_time.png)\n\n")
        f.write("#### 2. Speedup Factor over openpyxl\n")
        f.write("![Writing Speedup](phase3_speedup.png)\n\n")
        f.write("#### 3. Peak RSS Memory Utilization\n")
        f.write("![Writing Memory](phase3_memory.png)\n\n")
        f.write("#### 4. CPU Consumption and Utilization Percentage\n")
        f.write("![Writing CPU](phase3_cpu_utilization.png)\n\n")

        f.write("## 2. Phase 2: High-Performance XLSX Reader Benchmarks\n\n")
        f.write("| Operation | openpyxl | RPyXL | Speedup | openpyxl RSS | RPyXL RSS | RPyXL Throughput |\n")
        f.write("|---|---|---|---|---|---|---|\n")
        for r in p2_results:
            f.write(f"| {r['Operation']} | {r['openpyxl']} | {r['rpyxl']} | **{r['speedup']}** | {r['openpyxl_rss']} | {r['rpyxl_rss']} | {r['throughput']} |\n")

        f.write("\n### Phase 2 Performance Graphs\n\n")
        f.write("#### 1. Total Reading Execution Time (10,000 rows)\n")
        f.write("![Reading Time](phase2_reading_time.png)\n\n")
        f.write("#### 2. Speedup Factor over openpyxl\n")
        f.write("![Reading Speedup](phase2_speedup.png)\n\n")
        f.write("#### 3. Peak RSS Memory Utilization\n")
        f.write("![Reading Memory](phase2_memory.png)\n\n")

        f.write("## 3. Phase 1: In-Memory Operations Benchmarks\n\n")
        f.write("| Operation | Dataset | openpyxl (s) | RPyXL (s) | Speedup | openpyxl RSS | RPyXL RSS | Mem Reduction | RPyXL Throughput |\n")
        f.write("|---|---|---|---|---|---|---|---|---|\n")
        for r in p1_results:
            f.write(f"| {r['Operation']} | {r['Dataset']} | {r['openpyxl_time_s']}s | {r['rpyxl_time_s']}s | **{r['speedup']}x** | {r['openpyxl_rss_mb']} MB | {r['rpyxl_rss_mb']} MB | {r['mem_reduction_pct']}% | {r['rpyxl_throughput']} |\n")

    results_json = {
        "environment": env,
        "phase1": p1_results,
        "phase2": p2_results,
        "phase3": p3_results,
        "phase2_chart_data": p2_chart_data,
        "phase3_chart_data": p3_chart_data,
    }
    with open("benchmarks/reports/phase3_results.json", "w") as f:
        json.dump(results_json, f, indent=2)

    print("\n================================================================================")
    print("Benchmark completed successfully!")
    print(f"Report: {md_file}")
    print("JSON: benchmarks/reports/phase3_results.json")
    print("Plots: benchmarks/reports/phase3_*.png")
    print("================================================================================")


if __name__ == "__main__":
    run_all()
