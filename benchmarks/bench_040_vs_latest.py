import json
import os
import subprocess
import sys
import tempfile
import time

PYTHON_040 = "/private/tmp/venv_040/bin/python3"
PYTHON_LATEST = sys.executable

BENCHMARK_CODE = '''
import gc
import json
import os
import sys
import time
import tempfile
import rpyxl

WARMUP = 1
REPS = 3

def get_median(fn):
    times = []
    for _ in range(REPS):
        gc.collect()
        t0 = time.perf_counter()
        fn()
        t1 = time.perf_counter()
        times.append(t1 - t0)
    times.sort()
    return times[len(times) // 2]

results = {}

# 1. Workbook Creation
def bench_create():
    for _ in range(200):
        wb = rpyxl.Workbook()
        wb.close()
results["workbook_creation"] = get_median(bench_create)

# 2. Cell Access Write (100x100)
def bench_cell_write():
    wb = rpyxl.Workbook()
    ws = wb.active
    for r in range(1, 101):
        for c in range(1, 101):
            ws.cell(row=r, column=c, value=r * c)
    wb.close()
results["cell_write"] = get_median(bench_cell_write)

# 3. Row Appending (5000 rows x 10 cols)
sample_row = [100.5, "Alpha", True, 42, None, "Beta", 99.9, False, 12, "Gamma"]
def bench_append():
    wb = rpyxl.Workbook()
    ws = wb.active
    for _ in range(5000):
        ws.append(sample_row)
    wb.close()
results["append_rows"] = get_median(bench_append)

# 4. In-Memory Row Iteration (5000 rows x 10 cols)
wb_iter = rpyxl.Workbook()
ws_iter = wb_iter.active
for _ in range(5000):
    ws_iter.append(sample_row)

def bench_iter():
    cnt = 0
    for row in ws_iter.iter_rows(values_only=True):
        cnt += len(row)
results["row_iteration"] = get_median(bench_iter)
wb_iter.close()

# 5. Standard Save (10,000 rows x 10 cols)
wb_save = rpyxl.Workbook()
ws_save = wb_save.active
for _ in range(10000):
    ws_save.append(sample_row)

with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
    save_path = f.name

def bench_save():
    wb_save.save(save_path)
results["standard_save_10k"] = get_median(bench_save)
wb_save.close()

# 6. Write-Only Streaming (10,000 rows x 10 cols)
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
    wo_path_10k = f.name

def bench_write_only_10k():
    wb = rpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()
    for _ in range(10000):
        ws.append(sample_row)
    wb.save(wo_path_10k)
    wb.close()
results["write_only_10k"] = get_median(bench_write_only_10k)

# 7. Write-Only Streaming (50,000 rows x 10 cols)
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
    wo_path_50k = f.name

def bench_write_only_50k():
    wb = rpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()
    for _ in range(50000):
        ws.append(sample_row)
    wb.save(wo_path_50k)
    wb.close()
results["write_only_50k"] = get_median(bench_write_only_50k)

# 8. Full XLSX Read (Standard & read_only) using the 10k file
def bench_read_standard():
    wb = rpyxl.load_workbook(save_path, read_only=False)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
results["read_standard_10k"] = get_median(bench_read_standard)

def bench_read_readonly():
    wb = rpyxl.load_workbook(save_path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
results["read_readonly_10k"] = get_median(bench_read_readonly)

# 9. Time to First Row (read_only=True)
def bench_ttfr():
    wb = rpyxl.load_workbook(save_path, read_only=True)
    ws = wb.active
    first_row = next(ws.iter_rows(values_only=True))
    wb.close()
results["ttfr_readonly"] = get_median(bench_ttfr)

# 10. Roundtrip Write + Read (5000 rows x 10 cols)
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
    rt_path = f.name

def bench_roundtrip():
    wb = rpyxl.Workbook()
    ws = wb.active
    for _ in range(5000):
        ws.append(sample_row)
    wb.save(rt_path)
    wb.close()
    wb2 = rpyxl.load_workbook(rt_path, read_only=False)
    ws2 = wb2.active
    rows = list(ws2.iter_rows(values_only=True))
    wb2.close()
results["roundtrip_5k"] = get_median(bench_roundtrip)

# Cleanup
for p in [save_path, wo_path_10k, wo_path_50k, rt_path]:
    try:
        os.remove(p)
    except OSError:
        pass

print(json.dumps(results))
'''

def run_suite(python_bin, label):
    print(f"Running benchmarks for {label} ({python_bin})...")
    res = subprocess.run([python_bin, "-c", BENCHMARK_CODE], capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error in {label}:", res.stderr)
        sys.exit(1)
    return json.loads(res.stdout.strip().splitlines()[-1])

def main():
    res_040 = run_suite(PYTHON_040, "rpyxl 0.4.0")
    res_latest = run_suite(PYTHON_LATEST, "rpyxl latest (post-0.4.0)")

    labels = {
        "workbook_creation": ("Workbook Creation", "200 workbooks"),
        "cell_write": ("Cell Access Write", "10,000 cells (100x100)"),
        "append_rows": ("Row Appending", "5,000 rows x 10 cols"),
        "row_iteration": ("In-Memory Row Iteration", "5,000 rows x 10 cols"),
        "standard_save_10k": ("Standard Save", "10,000 rows x 10 cols"),
        "write_only_10k": ("Write-Only Streaming", "10,000 rows x 10 cols"),
        "write_only_50k": ("Write-Only Streaming (Large)", "50,000 rows x 10 cols"),
        "read_standard_10k": ("Full Read (Standard)", "10,000 rows x 10 cols"),
        "read_readonly_10k": ("Full Read (read_only=True)", "10,000 rows x 10 cols"),
        "ttfr_readonly": ("Time to First Row (read_only)", "10,000 rows x 10 cols"),
        "roundtrip_5k": ("Roundtrip Write + Read", "5,000 rows x 10 cols"),
    }

    print("\n" + "="*80)
    print("RPyXL 0.4.0 vs Latest Changes Benchmark Comparison")
    print("="*80)
    print(f"{'Benchmark Operation':<32} | {'Dataset':<26} | {'v0.4.0':<10} | {'Latest':<10} | {'Speedup':<10}")
    print("-" * 98)

    for key, (op_name, dataset) in labels.items():
        t040 = res_040.get(key, 0.0)
        tlatest = res_latest.get(key, 0.0)
        speedup = t040 / tlatest if tlatest > 0 else 1.0

        if "ttfr" in key:
            t040_str = f"{t040 * 1000:.2f} ms"
            tlatest_str = f"{tlatest * 1000:.2f} ms"
        else:
            t040_str = f"{t040:.4f} s"
            tlatest_str = f"{tlatest:.4f} s"

        speedup_str = f"{speedup:.2f}x"
        print(f"{op_name:<32} | {dataset:<26} | {t040_str:<10} | {tlatest_str:<10} | {speedup_str:<10}")

    print("="*98)

if __name__ == "__main__":
    main()
