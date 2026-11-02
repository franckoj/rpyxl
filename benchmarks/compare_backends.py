"""Fresh-process comparisons of the public native/rust_xlsxwriter/openpyxl writers.

Example (intentionally substantial; do not run alongside correctness tests)::

    python benchmarks/compare_backends.py --backends native rust_xlsxwriter openpyxl

Each case uses at least seven rotated, paired repetitions. The default target is
20,000 x 20, including a styled case; a separate 1,000-row plain append control
can additionally compare native/openpyxl write-only modes (--include-write-only).
Candidate write-only is NOT implemented and is never benchmarked. No Calamine
experiment is included here. No packages are installed by this driver.

Imports/provenance and complete semantic validation are outside the measured
create + generate/append/style + save + final-close interval. Save includes all
backend staging/fixup work. RSS is sampled, not an exact allocation measurement.
ZIP methods are observable, but compression levels are not; these are library-
default comparisons, NOT strict compression-matched like-for-like results.
Every invocation creates a new report directory; old results are never loaded.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import importlib.metadata
import json
import math
import os
import platform
import random
import statistics
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
from collections import Counter
from datetime import datetime, timedelta, timezone
from itertools import zip_longest
from pathlib import Path
from zipfile import ZipFile

BACKENDS = ("native", "rust_xlsxwriter", "openpyxl")
SCENARIOS = ("plain", "styled", "formulas", "dates")
DATE_FORMAT = "yyyy-mm-dd hh:mm:ss.000"
SEED = 20261102
COMPRESSION_POLICY = (
    "Library defaults, not strict like-for-like: native zip 2 vs candidate zip 9 "
    "and Python zipfile may use different deflaters/default levels. ZIP files do "
    "not record the compression level; no common public level control is used."
)


def _json_write(path, value):
    Path(path).write_text(
        json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def _sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _distribution_versions():
    versions = []
    for distribution in importlib.metadata.distributions():
        # Wheel metadata directory names encode normalized name/version. Avoid
        # hydrating every METADATA file on cloud-backed environments. Retain a
        # public metadata fallback for non-filesystem/legacy distributions.
        metadata_path = getattr(distribution, "_path", None)
        leaf = getattr(metadata_path, "name", "")
        if leaf.endswith(".dist-info") and "-" in leaf[:-10]:
            name, version = leaf[:-10].rsplit("-", 1)
        else:
            name = distribution.metadata.get("Name", "unknown")
            version = distribution.version
        versions.append((name, version, str(metadata_path)))
    return sorted(versions)


def _provenance():
    import openpyxl
    import psutil
    import rpyxl
    import rpyxl._core
    import rpyxl.workbook

    root = Path(__file__).resolve().parents[1]
    core = Path(rpyxl._core.__file__).resolve()
    workbook = Path(rpyxl.workbook.__file__).resolve()
    lock = root / "Cargo.lock"
    return {
        "python_executable": sys.executable,
        "python_version": sys.version,
        # platform's Windows WMI probes can block for minutes; the Win32 API
        # and process environment provide provenance without launching probes.
        "platform": str(sys.getwindowsversion())
        if os.name == "nt"
        else platform.platform(),
        "machine": (
            os.environ.get("PROCESSOR_ARCHITEW6432")
            or os.environ.get("PROCESSOR_ARCHITECTURE")
        )
        if os.name == "nt"
        else platform.machine(),
        "processor": os.environ.get("PROCESSOR_IDENTIFIER")
        if os.name == "nt"
        else platform.processor(),
        "logical_cpus": os.cpu_count(),
        "physical_memory_bytes": psutil.virtual_memory().total,
        "rpyxl_package_path": str(Path(rpyxl.__file__).resolve()),
        "rpyxl_version": rpyxl.__version__,
        "rpyxl_core_path": str(core),
        "rpyxl_core_sha256": _sha256(core),
        "rpyxl_workbook_path": str(workbook),
        "rpyxl_workbook_sha256": _sha256(workbook),
        "openpyxl_package_path": str(Path(openpyxl.__file__).resolve()),
        "openpyxl_version": openpyxl.__version__,
        "psutil_version": psutil.__version__,
        "all_distribution_versions": _distribution_versions(),
        "version_inventory_source": "Wheel metadata directory names; METADATA fallback for legacy installs",
        "driver_path": str(Path(__file__).resolve()),
        "driver_sha256": _sha256(__file__),
        "checkout_cargo_lock_sha256": _sha256(lock) if lock.exists() else None,
        "cargo_lock_note": "Checkout lockfile is not proof of installed wheel dependencies",
        "environment": {
            key: os.environ.get(key)
            for key in (
                "PYTHONPATH",
                "PYTHONHASHSEED",
                "RAYON_NUM_THREADS",
                "OMP_NUM_THREADS",
            )
        },
    }


def _rows(scenario, rows, cols, seed):
    """Regenerated for validation; never retain a second complete input matrix."""
    rng = random.Random(seed)
    epoch = datetime(2024, 1, 1, 12, 0, 0, 123000)
    for row in range(1, rows + 1):
        values = []
        for col in range(1, cols + 1):
            number = rng.randrange(1, 1_000_000)
            if scenario == "styled":
                value = number / 8
            elif scenario == "dates":
                value = epoch + timedelta(days=row % 366, seconds=col * 61)
            elif scenario == "formulas":
                value = f"=A{row}+{col}" if col % 2 == 0 else number
            elif col % 4 == 0:
                value = bool(number % 2)
            elif col % 4 == 1:
                value = number
            elif col % 4 == 2:
                value = number / 8
            else:
                value = f"日本語 café {row}:{col}:{number}\nline"
            values.append(value)
        yield values


def _style_models(backend):
    if backend == "openpyxl":
        from openpyxl import styles
    else:
        from rpyxl import styles

    # Explicit identical 8-digit FF-ARGB colors, including openpyxl. No six-hex
    # shortcuts (which rpyxl pads with 00), themes, tints, or inherited font color.
    return {
        "font": styles.Font(name="Arial", size=12, bold=True, color="FF123456"),
        "fill": styles.PatternFill("solid", fgColor="FFABCDEF"),
        "border": styles.Border(
            left=styles.Side("thin", "FF102030"),
            bottom=styles.Side("double", "FF405060"),
        ),
        "alignment": styles.Alignment(
            horizontal="center", vertical="top", wrap_text=True
        ),
        "protection": styles.Protection(locked=False, hidden=True),
        "number_format": '0.0000 "units"',
    }


def _style_signature(cell):
    def side(value):
        if value is None or value.style is None:
            return None
        color = value.color
        return [
            value.style,
            color.type if color is not None else None,
            color.rgb if color is not None and color.type == "rgb" else None,
        ]

    font = cell.font
    return {
        "font": [
            font.name,
            font.sz,
            bool(font.bold),
            bool(font.italic),
            font.color.type if font.color is not None else None,
            font.color.rgb
            if font.color is not None and font.color.type == "rgb"
            else None,
        ],
        "fill": [cell.fill.patternType, cell.fill.fgColor.type, cell.fill.fgColor.rgb],
        "border": [
            side(getattr(cell.border, name))
            for name in ("left", "right", "top", "bottom")
        ],
        "alignment": [
            cell.alignment.horizontal,
            cell.alignment.vertical,
            bool(cell.alignment.wrap_text),
            cell.alignment.textRotation,
            bool(cell.alignment.shrink_to_fit),
            cell.alignment.indent,
        ],
        "protection": [bool(cell.protection.locked), bool(cell.protection.hidden)],
        "number_format": cell.number_format,
    }


EXPECTED_STYLE = {
    "font": ["Arial", 12, True, False, "rgb", "FF123456"],
    "fill": ["solid", "rgb", "FFABCDEF"],
    "border": [["thin", "rgb", "FF102030"], None, None, ["double", "rgb", "FF405060"]],
    "alignment": ["center", "top", True, 0, False, 0],
    "protection": [False, True],
    "number_format": '0.0000 "units"',
}


def _canonical(value):
    if isinstance(value, bool):
        return ["bool", value]
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise ValueError("Nonfinite output")
        return ["number", format(value, ".17g")]
    if isinstance(value, datetime):
        return ["datetime", value.isoformat(timespec="milliseconds")]
    if isinstance(value, str):
        return ["formula" if value.startswith("=") else "text", value]
    if value is None:
        return ["empty", None]
    raise ValueError(f"Unexpected output value type {type(value)!r}")


def _validate(path, scenario, rows, cols, seed):
    """Full independent read, not file-size/sampling or ZIP-byte equality."""
    import openpyxl

    digest = hashlib.sha256()
    loaded = openpyxl.load_workbook(path, read_only=True, data_only=False)
    cells_checked = 0
    try:
        if loaded.sheetnames != ["Data"] or loaded.active.title != "Data":
            raise ValueError("Sheet order/name/active selection changed")
        if loaded.epoch.year != 1899:
            raise ValueError("Output is not a 1900-epoch workbook")
        marker = object()
        for row_num, (actual, expected) in enumerate(
            zip_longest(
                loaded.active.iter_rows(),
                _rows(scenario, rows, cols, seed),
                fillvalue=marker,
            ),
            1,
        ):
            if actual is marker or expected is marker or len(actual) != cols:
                raise ValueError(f"Output dimensions changed at row {row_num}")
            for cell, value in zip(actual, expected):
                actual_value = _canonical(cell.value)
                if actual_value != _canonical(value):
                    raise ValueError(
                        f"Value mismatch at {cell.coordinate}: {cell.value!r} != {value!r}"
                    )
                expected_type = (
                    "b"
                    if isinstance(value, bool)
                    else "d"
                    if isinstance(value, datetime)
                    else "f"
                    if isinstance(value, str) and value.startswith("=")
                    else "s"
                    if isinstance(value, str)
                    else "n"
                )
                if cell.data_type != expected_type:
                    raise ValueError(f"Cell type mismatch at {cell.coordinate}")
                signature = None
                if scenario == "styled":
                    signature = _style_signature(cell)
                    if signature != EXPECTED_STYLE:
                        raise ValueError(
                            f"Style mismatch at {cell.coordinate}: {signature!r}"
                        )
                elif scenario == "dates":
                    signature = cell.number_format
                    if signature != DATE_FORMAT:
                        raise ValueError(
                            f"Temporal format mismatch at {cell.coordinate}"
                        )
                elif cell.has_style:
                    raise ValueError(f"Unexpected explicit style at {cell.coordinate}")
                digest.update(
                    json.dumps(
                        [actual_value, signature],
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ).encode("utf-8")
                )
                digest.update(b"\n")
                cells_checked += 1
    finally:
        loaded.close()
    formula_caches_checked = 0
    if scenario == "formulas":
        cached = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            for row in cached.active.iter_rows():
                for cell in row:
                    if cell.column % 2 == 0:
                        if cell.value is not None:
                            raise ValueError(
                                f"Uncached formula became {cell.value!r} at {cell.coordinate}"
                            )
                        formula_caches_checked += 1
        finally:
            cached.close()
        if formula_caches_checked != rows * (cols // 2):
            raise ValueError("Formula cache validation did not cover every formula")
    if cells_checked != rows * cols:
        raise ValueError("Incomplete cell validation")
    with ZipFile(path) as archive:
        bad_member = archive.testzip()
        if bad_member is not None:
            raise ValueError(f"Invalid ZIP member: {bad_member}")
        members = archive.infolist()
        compression = {
            "policy": COMPRESSION_POLICY,
            "levels": "unknown; not stored in ZIP and not matched",
            "method_member_counts": dict(
                Counter(str(member.compress_type) for member in members)
            ),
            "method_names": {"0": "stored", "8": "deflate"},
            "uncompressed_bytes": sum(member.file_size for member in members),
            "compressed_member_bytes": sum(member.compress_size for member in members),
        }
    return {
        "passed": True,
        "reader": "openpyxl read_only; full values/types/authored styles + formula caches + ZIP CRC",
        "cells_checked": cells_checked,
        "formula_caches_checked": formula_caches_checked,
        "semantic_sha256": digest.hexdigest(),
        "compression": compression,
    }


class _RssSampler:
    def __init__(self):
        import psutil

        self.process = psutil.Process()
        self.baseline = self.process.memory_info().rss
        self.peak = self.baseline
        self.stop = threading.Event()
        self.thread = threading.Thread(target=self._sample, daemon=True)

    def _sample(self):
        while not self.stop.wait(0.01):
            self.peak = max(self.peak, self.process.memory_info().rss)

    def current(self):
        current = self.process.memory_info().rss
        self.peak = max(self.peak, current)
        return current

    def finish(self):
        self.stop.set()
        self.thread.join(timeout=2)
        final = self.current()
        info = self.process.memory_info()
        return {
            "baseline_bytes": self.baseline,
            "absolute_sampled_peak_bytes": self.peak,
            "incremental_sampled_peak_bytes": max(0, self.peak - self.baseline),
            "final_bytes": final,
            "windows_lifetime_peak_working_set_bytes": getattr(info, "peak_wset", None),
            "sampling_interval_seconds": 0.01,
            "note": "Sampled RSS may miss short peaks; lifetime OS peak includes imports, not validation",
        }


def _measure(args):
    import openpyxl
    import rpyxl

    if args.write_only and args.backend == "rust_xlsxwriter":
        raise NotImplementedError("Candidate write-only is not implemented")
    if args.write_only and args.scenario != "plain":
        raise NotImplementedError(
            "Write-only control is plain append only; do not imply style support"
        )
    if args.backend == "rust_xlsxwriter" and not callable(
        getattr(rpyxl.NativeWorkbook, "_save_rust_xlsxwriter", None)
    ):
        raise RuntimeError(
            "Rebuild/install the wheel: real _save_rust_xlsxwriter ABI required"
        )
    gc.collect()
    sampler = _RssSampler()
    sampler.thread.start()
    stages = {}
    book = None
    closed = False

    def stage(name, operation):
        wall, cpu = time.perf_counter(), time.process_time()
        result = operation()
        stages[name] = {
            "wall_seconds": time.perf_counter() - wall,
            "cpu_seconds": time.process_time() - cpu,
            "rss_at_end_bytes": sampler.current(),
        }
        return result

    def create():
        nonlocal book
        if args.backend == "openpyxl":
            book = openpyxl.Workbook(write_only=args.write_only)
        elif args.backend == "native":
            # Works with the baseline PyPI 0.2.0 wheel as well as the new API.
            book = rpyxl.Workbook(write_only=args.write_only)
        else:
            book = rpyxl.Workbook(backend="rust_xlsxwriter")
            if book.backend != "rust_xlsxwriter" or not callable(
                getattr(book._native, "_save_rust_xlsxwriter", None)
            ):
                raise RuntimeError(
                    "Rebuild/install the wheel: real _save_rust_xlsxwriter ABI required"
                )
        sheet = book.create_sheet("Data") if args.write_only else book.active
        if not args.write_only:
            sheet.title = "Data"
        styles = _style_models(args.backend) if args.scenario == "styled" else None
        return sheet, styles

    def append():
        for row_num, values in enumerate(
            _rows(args.scenario, args.rows, args.cols, args.seed), 1
        ):
            sheet.append(values)
            if styles is not None:
                for col in range(1, args.cols + 1):
                    cell = sheet.cell(row_num, col)
                    for name, value in styles.items():
                        setattr(cell, name, value)
            elif args.scenario == "dates":
                for col in range(1, args.cols + 1):
                    sheet.cell(row_num, col).number_format = DATE_FORMAT

    wall, cpu = time.perf_counter(), time.process_time()
    try:
        sheet, styles = stage("create", create)
        stage("generate_append_style", append)
        stage("save", lambda: book.save(args.workbook))
        stage("final_close", book.close)
        closed = True
        total = {
            "wall_seconds": time.perf_counter() - wall,
            "cpu_seconds": time.process_time() - cpu,
        }
    finally:
        rss = sampler.finish()
        if book is not None and not closed:
            book.close()
    return {
        "stages": stages,
        "end_to_end": total,
        "rss": rss,
        "output_bytes": Path(args.workbook).stat().st_size,
    }


def _child(args):
    result = {
        "status": "error",
        "backend": args.backend,
        "write_only": args.write_only,
        "scenario": args.scenario,
        "rows": args.rows,
        "cols": args.cols,
        "seed": args.seed,
        "pid": os.getpid(),
        "compression_policy": COMPRESSION_POLICY,
    }
    try:
        result["provenance"] = _provenance()
        print(
            f"Measuring {args.backend}, {args.scenario}, {args.rows}x{args.cols}, "
            f"write_only={args.write_only}",
            flush=True,
        )
        result.update(_measure(args))
        print(
            "Timing complete; validating every cell outside measured interval",
            flush=True,
        )
        started = time.perf_counter()
        result["validation"] = _validate(
            args.workbook, args.scenario, args.rows, args.cols, args.seed
        )
        result["validation_wall_seconds"] = time.perf_counter() - started
        result["output_sha256"] = _sha256(args.workbook)
        result["status"] = "ok"
    except Exception as exc:  # noqa: BLE001 - child boundary retains failed-run provenance
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["traceback"] = traceback.format_exc()
        traceback.print_exc()
    finally:
        _json_write(args.result, result)
    return 0 if result["status"] == "ok" else 1


def _orders(variants, repetitions):
    for repetition in range(repetitions):
        offset = repetition % len(variants)
        order = variants[offset:] + variants[:offset]
        if (repetition // len(variants)) % 2:
            order = list(reversed(order))
        yield repetition, order


def _summary(records, expected_count):
    good = [
        record
        for record in records
        if record.get("status") == "ok" and record.get("validation", {}).get("passed")
    ]
    groups = {}
    for record in records:
        groups.setdefault(record["case"], []).append(record)
    summaries = {}
    for case, runs in groups.items():
        valid = [run for run in runs if run in good]
        digests = {run["validation"]["semantic_sha256"] for run in valid}
        provenance = {
            (
                run["provenance"]["rpyxl_core_sha256"],
                run["provenance"]["rpyxl_workbook_sha256"],
                run["provenance"]["openpyxl_version"],
                run["provenance"]["driver_sha256"],
            )
            for run in valid
        }
        complete = len(valid) == len(runs) and all(
            run["status"] == "ok" for run in runs
        )
        comparable = complete and len(digests) == 1 and len(provenance) == 1
        variants = {}
        for run in valid:
            variants.setdefault(run["variant"], []).append(run)
        stats = {}
        for variant, samples in variants.items():

            def distribution(values):
                return {
                    "median": statistics.median(values),
                    "min": min(values),
                    "max": max(values),
                }

            stats[variant] = {
                "runs": len(samples),
                "end_to_end_wall_seconds": distribution(
                    [r["end_to_end"]["wall_seconds"] for r in samples]
                ),
                "end_to_end_cpu_seconds": distribution(
                    [r["end_to_end"]["cpu_seconds"] for r in samples]
                ),
                "save_wall_seconds": distribution(
                    [r["stages"]["save"]["wall_seconds"] for r in samples]
                ),
                "absolute_rss_peak_bytes": distribution(
                    [r["rss"]["absolute_sampled_peak_bytes"] for r in samples]
                ),
                "incremental_rss_peak_bytes": distribution(
                    [r["rss"]["incremental_sampled_peak_bytes"] for r in samples]
                ),
                "output_bytes": distribution([r["output_bytes"] for r in samples]),
            }
        summaries[case] = {
            "semantic_comparable": comparable,
            "compression_matched_like_for_like": False,
            "reason": "All outputs fully validated and semantic digests agree"
            if comparable
            else "Missing/failed validation, differing semantics, or differing child provenance",
            "variants": stats,
        }
    return {
        "scope": "This invocation only; no historical aggregation",
        "complete": len(records) == expected_count and len(good) == expected_count,
        "compression_policy": COMPRESSION_POLICY,
        "cases": summaries,
    }


def _parent(args):
    root = Path(args.output_dir).resolve()
    session = root / (
        datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        + "-"
        + uuid.uuid4().hex[:8]
    )
    session.mkdir(parents=True, exist_ok=False)
    variants = [(backend, False) for backend in args.backends]
    cases = [(scenario, args.rows, variants) for scenario in args.scenarios]
    controls = list(variants)
    if args.include_write_only:
        controls += [("native", True), ("openpyxl", True)]
    cases.append(("append_control", args.control_rows, controls))
    expected_count = sum(len(selected) * args.repetitions for _, _, selected in cases)
    manifest = {
        "argv": sys.argv,
        "configuration": vars(args),
        "expected_runs": expected_count,
        "started_utc": datetime.now(timezone.utc).isoformat(),
        "compression_policy": COMPRESSION_POLICY,
        "timing": "Create + generate/append/style + save + final close; imports and validation excluded",
        "baseline": "Native can use unmodified PyPI rpyxl 0.2.0; installed child paths/hash are authoritative",
        "order": "Fresh subprocess per run; same seed across backends, rotated and reverse-cycle paired orders",
        "warmups": 0,
        "historical_aggregation": False,
        "candidate_write_only": "Not implemented; never selected",
        "write_only_scope": "Optional smaller plain append control only",
    }
    _json_write(session / "manifest.json", manifest)
    print(
        f"Reports: {session}\nPlanned child runs: {expected_count}\n{COMPRESSION_POLICY}",
        flush=True,
    )
    records = []
    started = time.monotonic()
    stop = False
    with (session / "runs.jsonl").open("w", encoding="utf-8") as raw:
        for case, rows, selected in cases:
            if stop:
                break
            for repetition, order in _orders(selected, args.repetitions):
                if stop:
                    break
                for order_index, (backend, write_only) in enumerate(order):
                    remaining = args.parent_timeout - (time.monotonic() - started)
                    if remaining <= 0:
                        print(
                            "Parent timeout reached; stopping without declaring incomplete cases comparable",
                            flush=True,
                        )
                        stop = True
                        break
                    variant = f"write_only_{backend}" if write_only else backend
                    run_id = (
                        f"{case}-r{repetition + 1:02d}-o{order_index + 1}-{variant}"
                    )
                    result_path = session / f"{run_id}.json"
                    log_path = session / f"{run_id}.log"
                    context = {
                        "case": case,
                        "variant": variant,
                        "repetition": repetition + 1,
                        "order_index": order_index,
                        "run_id": run_id,
                        "log_path": str(log_path),
                        "result_path": str(result_path),
                    }
                    print(f"[{len(records) + 1}/{expected_count}] {run_id}", flush=True)
                    # Workbooks are temporary by default; raw metrics/validation/provenance
                    # and child stdout/stderr survive even failures and timeouts.
                    with tempfile.TemporaryDirectory(
                        prefix="rpyxl-backends-"
                    ) as temporary:
                        workbook = (
                            session / f"{run_id}.xlsx"
                            if args.keep_workbooks
                            else Path(temporary) / "output.xlsx"
                        )
                        command = [
                            sys.executable,
                            str(Path(__file__).resolve()),
                            "--child",
                            "--backend",
                            backend,
                            "--scenario",
                            "plain" if case == "append_control" else case,
                            "--rows",
                            str(rows),
                            "--cols",
                            str(args.cols),
                            "--seed",
                            str(args.seed),
                            "--workbook",
                            str(workbook),
                            "--result",
                            str(result_path),
                        ]
                        if write_only:
                            command.append("--write-only")
                        env = dict(os.environ, PYTHONHASHSEED="0")
                        with log_path.open("w", encoding="utf-8") as log:
                            log.write(
                                json.dumps({"command": command, **context}) + "\n"
                            )
                            log.flush()
                            try:
                                process = subprocess.run(
                                    command,
                                    stdout=log,
                                    stderr=subprocess.STDOUT,
                                    timeout=min(args.timeout, remaining),
                                    env=env,
                                    check=False,
                                )
                                if result_path.exists():
                                    record = json.loads(
                                        result_path.read_text(encoding="utf-8")
                                    )
                                else:
                                    record = {
                                        "status": "error",
                                        "error": "Child produced no result",
                                    }
                                record["returncode"] = process.returncode
                                if process.returncode != 0:
                                    record["status"] = "error"
                            except subprocess.TimeoutExpired:
                                record = {
                                    "status": "timeout",
                                    "error": "Child exceeded bounded runtime",
                                }
                                log.write(
                                    "\nChild timed out; killed and waited by subprocess.run\n"
                                )
                    record.update(context)
                    _json_write(result_path, record)
                    records.append(record)
                    raw.write(json.dumps(record, sort_keys=True) + "\n")
                    raw.flush()
                    print(f"  {record['status']}; log: {log_path.name}", flush=True)
    summary = _summary(records, expected_count)
    # Also gate each case on the intended run count if the parent deadline interrupted it.
    for case, _, selected in cases:
        expected_case = len(selected) * args.repetitions
        entry = summary["cases"].get(case)
        if (
            entry is not None
            and sum(r["case"] == case for r in records) != expected_case
        ):
            entry["semantic_comparable"] = False
            entry["reason"] = "Parent deadline interrupted the planned repetitions"
    _json_write(session / "summary.json", summary)
    print(f"Summary: {session / 'summary.json'}", flush=True)
    for case, values in summary["cases"].items():
        print(
            f"  {case}: semantic_comparable={values['semantic_comparable']} (library-default compression)",
            flush=True,
        )
    return (
        0
        if summary["complete"]
        and all(case["semantic_comparable"] for case in summary["cases"].values())
        else 1
    )


def _parser():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--backends", choices=BACKENDS, nargs="+", default=list(BACKENDS)
    )
    parser.add_argument(
        "--scenarios", choices=SCENARIOS, nargs="+", default=list(SCENARIOS)
    )
    parser.add_argument("--rows", type=int, default=20_000)
    parser.add_argument("--cols", type=int, default=20)
    parser.add_argument("--control-rows", type=int, default=1_000)
    parser.add_argument(
        "--repetitions", type=int, default=7, help="At least seven paired repetitions"
    )
    parser.add_argument("--seed", type=int, default=SEED)
    parser.add_argument(
        "--include-write-only",
        action="store_true",
        help="Add native + openpyxl write-only to the smaller plain append control only",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=300,
        help="Per-child timeout, including validation (seconds)",
    )
    parser.add_argument(
        "--parent-timeout",
        type=float,
        default=7_200,
        help="Overall child scheduling deadline (seconds)",
    )
    parser.add_argument(
        "--output-dir",
        default=str(Path(__file__).resolve().parent / "reports" / "backends"),
    )
    parser.add_argument("--keep-workbooks", action="store_true")
    parser.add_argument("--child", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--backend", choices=BACKENDS, help=argparse.SUPPRESS)
    parser.add_argument("--scenario", choices=SCENARIOS, help=argparse.SUPPRESS)
    parser.add_argument("--write-only", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--workbook", help=argparse.SUPPRESS)
    parser.add_argument("--result", help=argparse.SUPPRESS)
    return parser


def main():
    parser = _parser()
    args = parser.parse_args()
    if not 1 <= args.rows <= 1_048_576 or not 1 <= args.control_rows <= 1_048_576:
        parser.error("Rows and control rows must be within 1..1048576")
    if not 1 <= args.cols <= 16_384:
        parser.error("Columns must be within 1..16384")
    if args.repetitions < 7:
        parser.error(
            "At least seven repetitions are required; no underpowered comparisons"
        )
    if args.timeout <= 0 or args.parent_timeout <= 0:
        parser.error("Timeouts must be positive")
    if len(set(args.backends)) != len(args.backends) or len(set(args.scenarios)) != len(
        args.scenarios
    ):
        parser.error("Backend/scenario selections must not contain duplicates")
    if args.child:
        if not all((args.backend, args.scenario, args.workbook, args.result)):
            parser.error(
                "Internal child invocation requires backend/scenario/workbook/result"
            )
        return _child(args)
    if len(args.backends) < 2:
        parser.error("Select at least two backends to compare")
    return _parent(args)


if __name__ == "__main__":
    raise SystemExit(main())
