"""
Write random people rows with rpyxl and openpyxl and compare timings.

Install both first:   pip install rpyxl openpyxl
Run:                  python benchmark_rpyxl_vs_openpyxl.py

Each scenario is run REPEATS times per library. The two libraries are
interleaved within each repeat (rpyxl, openpyxl, rpyxl, openpyxl, ...) so
background load (OneDrive, antivirus, etc.) hits both about equally.
Results are reported as the MEDIAN, with min/max of the total shown too.

Environment variables:
  N_ROWS=50000     number of data rows (default 50,000 -> ~1.6 MB files;
                   500000 -> ~16 MB files, and much slower: lower REPEATS)
  REPEATS=5        runs per scenario per library (default 5)
  KEEP_FILES=1     keep the last generated .xlsx files instead of deleting them
"""

import importlib
import os
import random
import statistics
import time

N_ROWS = int(os.environ.get("N_ROWS", 50_000))
REPEATS = max(1, int(os.environ.get("REPEATS", 5)))
KEEP_FILES = os.environ.get("KEEP_FILES") == "1"
HEADERS = ["Name", "Age", "City", "Job Title", "Email"]
LIB_NAMES = ("rpyxl", "openpyxl")

first = [
    "Aarav",
    "Priya",
    "Liam",
    "Sofia",
    "Noah",
    "Meera",
    "Ethan",
    "Isha",
    "Lucas",
    "Zara",
    "Rohan",
    "Emma",
    "Kabir",
    "Olivia",
    "Arjun",
    "Mia",
    "Dev",
    "Ava",
    "Neil",
    "Anika",
]
last = [
    "Sharma",
    "Patel",
    "Smith",
    "Garcia",
    "Khan",
    "Iyer",
    "Brown",
    "Nair",
    "Wilson",
    "Mehta",
    "Kapoor",
    "Lee",
    "Rao",
    "Johnson",
    "Gupta",
    "Martin",
    "Joshi",
    "Clark",
    "Das",
    "Reddy",
]
cities = [
    "Mumbai",
    "Pune",
    "Delhi",
    "Bengaluru",
    "Chennai",
    "London",
    "Toronto",
    "Singapore",
    "Sydney",
    "Dubai",
]
jobs = [
    "Data Analyst",
    "Software Engineer",
    "Product Manager",
    "Designer",
    "Accountant",
    "Marketing Lead",
    "HR Specialist",
    "Sales Executive",
    "Teacher",
    "Nurse",
]

# Every output file path used by this run, so leftovers can be deleted at the end.
generated = set()


def remove_file(p):
    try:
        os.remove(p)
    except FileNotFoundError:
        pass
    except PermissionError:
        print(f"Could not delete {p} (locked by OneDrive/antivirus?)")


def cleanup():
    """Delete all generated .xlsx files (skipped when KEEP_FILES=1)."""
    if KEEP_FILES:
        print("KEEP_FILES=1 -> generated files kept:")
        for p in sorted(generated):
            if os.path.exists(p):
                print(f"  {p}")
        return
    for p in sorted(generated):
        remove_file(p)
    print("Generated files deleted.")


def make_rows(n):
    """Generate the data ONCE so both libraries write identical rows
    and data generation is not part of the timings."""
    random.seed(42)
    rows = []
    for i in range(n):
        f, l = random.choice(first), random.choice(last)
        rows.append(
            [
                f"{f} {l}",
                random.randint(22, 58),
                random.choice(cities),
                random.choice(jobs),
                f"{f.lower()}.{l.lower()}{i}@example.com",
            ]
        )
    return rows


def load_lib(name):
    try:
        return importlib.import_module(name)
    except ImportError:
        return None


def lib_version(name, lib):
    """Best-effort version lookup (some builds don't expose __version__)."""
    if lib is None:
        return "not installed"
    v = getattr(lib, "__version__", None)
    if v:
        return v
    try:
        from importlib.metadata import version

        return version(name)
    except Exception:
        return "unknown"


# ---------- Scenario 1: plain write (normal mode) ----------
def plain_write(lib, rows, path):
    t0 = time.perf_counter()
    wb = lib.Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    t1 = time.perf_counter()
    wb.save(path)
    t2 = time.perf_counter()
    return t1 - t0, t2 - t1


# ---------- Scenario 2: write_only (streaming) mode ----------
def write_only(lib, rows, path):
    t0 = time.perf_counter()
    wb = lib.Workbook(write_only=True)
    ws = wb.create_sheet("People")
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    t1 = time.perf_counter()
    wb.save(path)
    t2 = time.perf_counter()
    return t1 - t0, t2 - t1


# ---------- Scenario 3: styled write (normal mode) ----------
def styled_write(lib, rows, path):
    styles = importlib.import_module(lib.__name__ + ".styles")
    Font, PatternFill, Alignment = styles.Font, styles.PatternFill, styles.Alignment

    head_font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    head_fill = PatternFill("solid", fgColor="1F3A5F")
    stripe_fill = PatternFill("solid", fgColor="EAF1FB")
    body_font = Font(name="Arial", size=11)
    center = Alignment(horizontal="center")

    t0 = time.perf_counter()
    wb = lib.Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for c in ws[1]:
        c.font, c.fill = head_font, head_fill
    for i, r in enumerate(rows, start=2):
        stripe = i % 2 == 0
        for j, value in enumerate(r, start=1):
            c = ws.cell(row=i, column=j, value=value)
            c.font = body_font
            if stripe:
                c.fill = stripe_fill
            if j == 2:
                c.alignment = center
    ws.freeze_panes = "A2"
    t1 = time.perf_counter()
    wb.save(path)
    t2 = time.perf_counter()
    return t1 - t0, t2 - t1


SCENARIOS = [
    ("Plain write", plain_write),
    ("write_only (streaming)", write_only),
    ("Styled write", styled_write),
]


def summarize(runs):
    """runs: list of (build, save, size_mb). Returns dict of medians + min/max total."""
    builds = [r[0] for r in runs]
    saves = [r[1] for r in runs]
    totals = [r[0] + r[1] for r in runs]
    return {
        "build": statistics.median(builds),
        "save": statistics.median(saves),
        "total": statistics.median(totals),
        "tmin": min(totals),
        "tmax": max(totals),
        "size": runs[-1][2],
        "n": len(runs),
    }


def main():
    print(f"Generating {N_ROWS:,} random rows ...")
    t = time.perf_counter()
    rows = make_rows(N_ROWS)
    print(f"Data generated in {time.perf_counter() - t:.2f}s (not included in timings)")

    libs = {name: load_lib(name) for name in LIB_NAMES}

    print("\nLibrary versions:")
    for name, lib in libs.items():
        print(f"  {name}: {lib_version(name, lib)}")
    print(f"\nRows: {N_ROWS:,} | Repeats per scenario: {REPEATS} (median reported)\n")

    summary = {}

    try:
        for scenario, fn in SCENARIOS:
            print(f"Running: {scenario} ...", flush=True)
            raw = {name: [] for name in libs}
            failed = {}

            for rep in range(1, REPEATS + 1):
                # Interleave libraries within each repeat so noise hits both.
                for name, lib in libs.items():
                    if lib is None or name in failed:
                        continue
                    path = f"bench_{name}_{scenario.split()[0].lower()}.xlsx"
                    generated.add(path)
                    try:
                        build, save = fn(lib, rows, path)
                        size_mb = os.path.getsize(path) / 1e6
                        raw[name].append((build, save, size_mb))
                        print(
                            f"  run {rep}/{REPEATS}  {name:<9} "
                            f"build {build:6.2f}s  save {save:6.2f}s  total {build + save:6.2f}s",
                            flush=True,
                        )
                    except Exception as e:  # unsupported feature, etc.
                        failed[name] = f"failed: {type(e).__name__}: {e}"
                        print(
                            f"  run {rep}/{REPEATS}  {name:<9} {failed[name]}",
                            flush=True,
                        )
                    finally:
                        # Delete between runs so large files don't pile up
                        # (the last run's files stay if KEEP_FILES=1).
                        if not KEEP_FILES or rep < REPEATS:
                            remove_file(path)

            for name, lib in libs.items():
                if lib is None:
                    summary[(scenario, name)] = "not installed"
                elif name in failed:
                    summary[(scenario, name)] = failed[name]
                else:
                    summary[(scenario, name)] = summarize(raw[name])
            print()

        # ---------- Report ----------
        print(f"MEDIAN of {REPEATS} runs per scenario (rows: {N_ROWS:,})")
        print(
            f"{'Scenario':<26}{'Library':<10}{'Build (s)':>10}{'Save (s)':>10}"
            f"{'Total (s)':>11}{'Min-Max total':>16}{'Size (MB)':>11}"
        )
        print("-" * 94)
        for scenario, _ in SCENARIOS:
            for name in libs:
                r = summary[(scenario, name)]
                if isinstance(r, str):
                    print(f"{scenario:<26}{name:<10}{r}")
                else:
                    rng = f"{r['tmin']:.2f}-{r['tmax']:.2f}"
                    print(
                        f"{scenario:<26}{name:<10}{r['build']:>10.2f}{r['save']:>10.2f}"
                        f"{r['total']:>11.2f}{rng:>16}{r['size']:>11.2f}"
                    )
            a, b_ = summary[(scenario, "rpyxl")], summary[(scenario, "openpyxl")]
            if not isinstance(a, str) and not isinstance(b_, str):
                ratio = b_["total"] / a["total"]
                print(f"{'':<26}-> openpyxl total / rpyxl total = {ratio:.2f}x (median)")
            print()
    finally:
        # Runs even if a scenario crashes or the run is interrupted.
        cleanup()


if __name__ == "__main__":
    main()
