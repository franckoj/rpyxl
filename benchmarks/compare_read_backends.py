"""Isolated cached-value reader comparison of one existing XLSX; no generation.

Run only after other benchmarks/tests finish. Supply the already-built executable
and an existing scalar fixture via --executable and --input. Each invocation uses
at least seven fresh processes per backend, sequentially in alternating order,
and writes a new session below ignored dist/reader-results (no historical merge).

Native loads a full Workbook; Calamine streams worksheet cells without a Range
or RPyXL copy. These are different storage/API contracts, using library defaults,
NOT a like-for-like parser/API or security comparison. No production routing is
changed. RSS is externally sampled with psutil, not an exact allocation peak;
CPU is the last available external sample, not guaranteed final process CPU.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import platform
import re
import statistics
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

BACKENDS = ("native", "calamine")
STORAGE = {
    "native": "full_native_to_workbook_then_borrowed_cell_iteration",
    "calamine": "xlsx_cell_reader_stream_no_range_no_rpyxl_copy",
}
TIMINGS = (
    "prevalidate_open_zip",
    "backend_open_load",
    "consume",
    "backend_total_with_drop",
    "overall_total",
)
POLICY = {
    "native": "Full native Workbook, then borrowed cell iteration",
    "calamine": "Worksheet cell stream; no Range and no RPyXL copy",
    "same_api_or_storage_contract": False,
    "library_defaults": True,
    "production_routing_changed": False,
    "mode": "Cached scalar values only; no formula evaluation or generation",
    "security": "Shared ZIP prevalidation; native rechecks; XML security differs",
    "timing": "Use backend_total_with_drop for combined open/load + consume + drop; phase boundaries differ",
    "wall": "External child lifetime also includes startup and JSON output",
    "rss": "Absolute sampled child RSS, not parent RSS or exact allocation high-water mark",
    "cpu": "Last sampled user+system seconds; lower bound, final sample may precede exit",
    "cache": "No cache eviction or discarded warmups; hashing warms file cache; order rotates",
    "semantics": "All runs must agree on ordered sheet names, nonempty bounds/counts, checksum_cells and scalar_checksum; zero excluded cells",
    "types": "Raw type histograms retained, not compared; executable normalizes exactly integral float/int checksums; empty storage may differ",
    "checksum": "Executable's noncryptographic scalar checksum is a comparison gate, not a fidelity proof",
}


def _json_write(path, value):
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def _identity(path):
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": digest.hexdigest(),
    }


def _orders(repetitions):
    for repetition in range(repetitions):
        offset = repetition % len(BACKENDS)
        yield repetition + 1, BACKENDS[offset:] + BACKENDS[:offset]


def _count(value, label):
    if type(value) is not int or value < 0:
        raise ValueError(f"{label} must be a nonnegative integer")
    return value


def _signature(report, backend):
    """Validate this scalar-only report; deliberately do not compare type maps."""
    if not isinstance(report, dict):
        raise TypeError("Report must be a JSON object")
    if report.get("schema") != "backend_read_v1" or report.get("backend") != backend:
        raise ValueError("Unexpected report schema/backend")
    if report.get("mode") != "cached_values_only_no_formula_evaluation":
        raise ValueError("Unexpected cached-value mode")
    if report.get("threads") != 1 or report.get("storage") != STORAGE[backend]:
        raise ValueError("Unexpected threads/storage contract")
    timing = report.get("timing_ns")
    if not isinstance(timing, dict):
        raise TypeError("Missing timing_ns object")
    for key in TIMINGS:
        _count(timing.get(key), f"timing_ns.{key}")
    sheets = report.get("sheets")
    if not isinstance(sheets, list) or not sheets:
        raise ValueError("Missing worksheet summaries")
    signature = []
    names = set()
    total_visited = total_nonempty = 0
    for sheet in sheets:
        if not isinstance(sheet, dict):
            raise TypeError("Sheet summary must be a JSON object")
        name = sheet.get("name")
        if not isinstance(name, str) or not name or name in names:
            raise ValueError("Missing/duplicate worksheet name")
        names.add(name)
        visited, nonempty, checksum_cells, excluded = (
            _count(sheet.get(key), f"{name}.{key}")
            for key in ("cells_visited", "nonempty", "checksum_cells", "excluded_cells")
        )
        if excluded != 0:
            raise ValueError(f"{name}: excluded cells are unsupported, not a pass")
        if nonempty > visited or checksum_cells != nonempty:
            raise ValueError(f"{name}: inconsistent nonempty/checksum counts")
        types = sheet.get("types")
        if not isinstance(types, dict) or any(
            key not in {"empty", "bool", "string", "int", "float"} for key in types
        ):
            raise ValueError(f"{name}: missing or unsupported scalar type histogram")
        if (
            sum(_count(value, f"{name}.types.{key}") for key, value in types.items())
            != visited
        ):
            raise ValueError(f"{name}: type histogram does not sum to cells_visited")
        bounds = sheet.get("nonempty_bounds_1based")
        if nonempty == 0:
            if bounds is not None:
                raise ValueError(f"{name}: empty sheet has nonempty bounds")
        elif (
            not isinstance(bounds, list)
            or len(bounds) != 4
            or any(type(value) is not int or value < 1 for value in bounds)
            or not (bounds[0] <= bounds[2] <= 1_048_576)
            or not (bounds[1] <= bounds[3] <= 16_384)
        ):
            raise ValueError(f"{name}: invalid one-based nonempty bounds")
        checksum = sheet.get("scalar_checksum")
        if not isinstance(checksum, str) or not re.fullmatch(r"[0-9a-f]{16}", checksum):
            raise ValueError(f"{name}: invalid scalar_checksum")
        signature.append(
            {
                "name": name,
                "nonempty": nonempty,
                "nonempty_bounds_1based": bounds,
                "checksum_cells": checksum_cells,
                "scalar_checksum": checksum,
            }
        )
        total_visited += visited
        total_nonempty += nonempty
    if (
        _count(report.get("cells_visited"), "cells_visited") != total_visited
        or _count(report.get("nonempty"), "nonempty") != total_nonempty
        or total_nonempty == 0
    ):
        raise ValueError("Inconsistent workbook totals or no nonempty scalar cells")
    return signature


def _run(command, backend, session, run_id, timeout, interval, psutil):
    stdout_path = session / f"{run_id}.stdout.json"
    stderr_path = session / f"{run_id}.stderr.log"
    record = {
        "run_id": run_id,
        "backend": backend,
        "command": command,
        "timeout_seconds": timeout,
        "stdout_path": str(stdout_path),
        "stderr_path": str(stderr_path),
        "status": "error",
        "rss_sample_count": 0,
        "sampled_peak_rss_bytes": None,
        "cpu_sample_count": 0,
        "last_sampled_cpu_seconds": None,
        "last_cpu_sample_elapsed_seconds": None,
    }
    process = None
    started = time.monotonic()
    try:
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            process = subprocess.Popen(
                command,
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                cwd=str(Path(__file__).resolve().parents[1]),
                shell=False,
            )
            record["pid"] = process.pid
            monitored = psutil.Process(process.pid)
            deadline = started + timeout
            while True:
                try:
                    rss = monitored.memory_info().rss
                    record["rss_sample_count"] += 1
                    record["sampled_peak_rss_bytes"] = max(
                        record["sampled_peak_rss_bytes"] or 0, rss
                    )
                except psutil.Error as exc:
                    record["rss_sampling_last_error"] = str(exc)
                try:
                    cpu = monitored.cpu_times()
                    record["last_sampled_cpu_seconds"] = cpu.user + cpu.system
                    record["cpu_sample_count"] += 1
                    record["last_cpu_sample_elapsed_seconds"] = (
                        time.monotonic() - started
                    )
                except psutil.Error as exc:
                    record["cpu_sampling_last_error"] = str(exc)
                if process.poll() is not None:
                    record["status"] = "ok" if process.returncode == 0 else "error"
                    if process.returncode != 0:
                        record["error"] = f"Child exited with code {process.returncode}"
                    break
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    record["status"] = "timeout"
                    record["error"] = "Child exceeded bounded runtime"
                    break
                time.sleep(min(interval, remaining))
    except (OSError, psutil.Error, subprocess.SubprocessError) as exc:
        record["error"] = f"{type(exc).__name__}: {exc}"
    finally:
        # Also reap a child on timeout, monitoring failure, or Ctrl-C. No next child
        # starts until this one has exited; failure to reap aborts scheduling.
        if process is not None:
            if process.poll() is None:
                process.kill()
            process.wait(timeout=5)
            record["returncode"] = process.returncode
        record["process_wall_seconds"] = time.monotonic() - started
    if record["status"] == "ok":
        try:
            record["report"] = json.loads(stdout_path.read_bytes())
            record["semantic_signature"] = _signature(record["report"], backend)
            if not record["rss_sample_count"] or not record["sampled_peak_rss_bytes"]:
                raise ValueError(
                    "No positive external RSS sample; measurement incomplete"
                )
        except (ValueError, TypeError, OSError) as exc:
            record["status"] = "error"
            record["error"] = f"Invalid/incomplete measurement: {exc}"
    return record


def _distribution(values):
    values = [value for value in values if value is not None]
    if not values:
        return None
    return {
        "samples": len(values),
        "median": statistics.median(values),
        "min": min(values),
        "max": max(values),
    }


def _summary(records, repetitions, identities_match, scheduling_error=None):
    expected = repetitions * len(BACKENDS)
    good = [record for record in records if record["status"] == "ok"]
    counts = {
        backend: sum(r["backend"] == backend for r in good) for backend in BACKENDS
    }
    complete = (
        len(records) == len(good) == expected
        and all(count == repetitions for count in counts.values())
        and scheduling_error is None
    )
    reference = good[0] if good else None
    mismatches = []
    if reference:
        for record in good[1:]:
            if record["semantic_signature"] != reference["semantic_signature"]:
                mismatches.append(
                    {
                        "run_id": record["run_id"],
                        "reference_run_id": reference["run_id"],
                    }
                )
    comparable = complete and identities_match and not mismatches
    stats = {}
    # Keep incomplete or semantically mismatched measurements in raw files, but
    # never publish them as a successful performance comparison.
    if comparable:
        for backend in BACKENDS:
            runs = [record for record in good if record["backend"] == backend]
            stats[backend] = {
                key: _distribution([record[key] for record in runs])
                for key in (
                    "process_wall_seconds",
                    "sampled_peak_rss_bytes",
                    "last_sampled_cpu_seconds",
                )
            }
            stats[backend]["timing_ns"] = {
                key: _distribution(
                    [record["report"]["timing_ns"][key] for record in runs]
                )
                for key in TIMINGS
            }
    return {
        "schema": "isolated_reader_comparison_v1",
        "scope": "This invocation only; no historical aggregation",
        "policy": POLICY,
        "expected_runs": expected,
        "recorded_runs": len(records),
        "valid_runs_per_backend": counts,
        "complete": complete,
        "artifact_hashes_unchanged": identities_match,
        "semantic_comparable": comparable,
        "passed": comparable,
        "scheduling_error": scheduling_error,
        "mismatches": mismatches,
        "failures": [
            {"run_id": r["run_id"], "status": r["status"], "error": r.get("error")}
            for r in records
            if r["status"] != "ok"
        ],
        "reference_sheets": reference["semantic_signature"] if reference else None,
        "statistics": stats,
    }


def _parser():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--executable", required=True, type=Path)
    parser.add_argument(
        "--input", required=True, type=Path, help="Existing scalar XLSX"
    )
    parser.add_argument(
        "--repetitions", type=int, default=7, help="At least seven per backend"
    )
    parser.add_argument("--timeout", type=float, default=120, help="Per-child seconds")
    parser.add_argument(
        "--parent-timeout", type=float, default=3600, help="Scheduling seconds"
    )
    parser.add_argument(
        "--sample-interval", type=float, default=0.005, help="RSS/CPU sample seconds"
    )
    parser.add_argument(
        "--output-dir", type=Path, default=root / "dist" / "reader-results"
    )
    return parser


def main(argv=None):
    parser = _parser()
    args = parser.parse_args(argv)
    if args.repetitions < 7:
        parser.error("At least seven fresh processes per backend are required")
    for key in ("timeout", "parent_timeout", "sample_interval"):
        value = getattr(args, key)
        if not math.isfinite(value) or value <= 0:
            parser.error(f"--{key.replace('_', '-')} must be finite and positive")
    args.executable = args.executable.resolve()
    args.input = args.input.resolve()
    for path in (args.executable, args.input):
        if not path.is_file():
            parser.error(f"Not an existing file: {path}")
    args.output_dir = args.output_dir.resolve()
    dist = Path(__file__).resolve().parents[1] / "dist"
    if (
        dist.resolve() not in args.output_dir.parents
        and args.output_dir != dist.resolve()
    ):
        parser.error("--output-dir must be inside the project's ignored dist directory")
    try:
        import psutil
    except ImportError:
        parser.error(
            "psutil is required in the driver interpreter; no packages are installed automatically"
        )

    session = args.output_dir / (
        datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        + "-"
        + uuid.uuid4().hex[:8]
    )
    session.mkdir(parents=True, exist_ok=False)
    artifacts = {"executable": args.executable, "input": args.input}
    before = {key: _identity(path) for key, path in artifacts.items()}
    schedule = [
        {
            "run_id": f"read-r{repetition:02d}-o{order_index}-{backend}",
            "repetition": repetition,
            "order_index": order_index,
            "backend": backend,
        }
        for repetition, order in _orders(args.repetitions)
        for order_index, backend in enumerate(order, 1)
    ]
    manifest = {
        "schema": "isolated_reader_comparison_v1",
        "started_utc": datetime.now(timezone.utc).isoformat(),
        "configuration": {
            key: str(value) if isinstance(value, Path) else value
            for key, value in vars(args).items()
        },
        "artifacts_before": before,
        "driver": _identity(Path(__file__).resolve()),
        "policy": POLICY,
        "schedule": schedule,
        "warmups": 0,
        "parallel_children": 1,
        "environment": {
            key: os.environ.get(key)
            for key in ("RAYON_NUM_THREADS", "OMP_NUM_THREADS", "PROCESSOR_IDENTIFIER")
        },
        "host": {
            "python_executable": sys.executable,
            "python_version": sys.version,
            "platform": str(sys.getwindowsversion())
            if os.name == "nt"
            else platform.platform(),
            "logical_cpus": os.cpu_count(),
            "physical_memory_bytes": psutil.virtual_memory().total,
            "psutil_version": psutil.__version__,
        },
    }
    _json_write(session / "manifest.json", manifest)
    print(f"Reports: {session}\nSequential child runs: {len(schedule)}", flush=True)
    records = []
    started = time.monotonic()
    scheduling_error = None
    try:
        with (session / "runs.jsonl").open("w", encoding="utf-8") as raw:
            for context in schedule:
                remaining = args.parent_timeout - (time.monotonic() - started)
                if remaining <= 0:
                    scheduling_error = "Parent scheduling deadline reached"
                    break
                print(
                    f"[{len(records) + 1}/{len(schedule)}] {context['run_id']}",
                    flush=True,
                )
                command = [str(args.executable), context["backend"], str(args.input)]
                record = _run(
                    command,
                    context["backend"],
                    session,
                    context["run_id"],
                    min(args.timeout, remaining),
                    args.sample_interval,
                    psutil,
                )
                record.update(context)
                _json_write(session / f"{context['run_id']}.json", record)
                records.append(record)
                raw.write(json.dumps(record, sort_keys=True) + "\n")
                raw.flush()
                print(
                    f"  {record['status']}: {record.get('error', 'validated')}",
                    flush=True,
                )
    except (KeyboardInterrupt, OSError, subprocess.SubprocessError) as exc:
        scheduling_error = f"{type(exc).__name__}: {exc}"
    try:
        after = {key: _identity(path) for key, path in artifacts.items()}
        identities_match = before == after
    except OSError as exc:
        after = {"error": str(exc)}
        identities_match = False
    manifest["artifacts_after"] = after
    manifest["finished_utc"] = datetime.now(timezone.utc).isoformat()
    _json_write(session / "manifest.json", manifest)
    summary = _summary(records, args.repetitions, identities_match, scheduling_error)
    _json_write(session / "summary.json", summary)
    print(
        f"Summary: {session / 'summary.json'}\nPassed: {summary['passed']}", flush=True
    )
    return 0 if summary["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
