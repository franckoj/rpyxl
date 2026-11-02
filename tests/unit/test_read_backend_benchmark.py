"""Pure unit tests for the checkout's isolated reader benchmark driver."""

import importlib.util
from pathlib import Path

import pytest


@pytest.fixture(scope="module")
def driver():
    root = Path(__file__).resolve().parents[2]
    path = root / "benchmarks" / "compare_read_backends.py"
    spec = importlib.util.spec_from_file_location("read_backend_benchmark_driver", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(autouse=True)
def forbid_processes(driver, monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail("Unit tests must not launch benchmark processes")

    monkeypatch.setattr(driver.subprocess, "Popen", forbidden)
    monkeypatch.setattr(driver, "_run", forbidden)


def _report(driver, backend):
    types = (
        {"bool": 1, "string": 1, "int": 2, "float": 1}
        if backend == "native"
        else {"bool": 1, "string": 1, "float": 3}
    )
    return {
        "schema": "backend_read_v1",
        "backend": backend,
        "mode": "cached_values_only_no_formula_evaluation",
        "threads": 1,
        "storage": driver.STORAGE[backend],
        "timing_ns": dict.fromkeys(driver.TIMINGS, 100),
        "cells_visited": 5,
        "nonempty": 5,
        "sheets": [
            {
                "name": "Scalars Ω",
                "cells_visited": 5,
                "nonempty": 5,
                "nonempty_bounds_1based": [1, 1, 1, 5],
                "types": types,
                "checksum_cells": 5,
                "excluded_cells": 0,
                "scalar_checksum": "0123456789abcdef",
            },
            {
                "name": "Empty",
                "cells_visited": 0,
                "nonempty": 0,
                "nonempty_bounds_1based": None,
                "types": {},
                "checksum_cells": 0,
                "excluded_cells": 0,
                "scalar_checksum": "0000000000000000",
            },
        ],
    }


@pytest.fixture
def records(driver):
    runs = []
    for repetition, order in driver._orders(7):
        for order_index, backend in enumerate(order, 1):
            report = _report(driver, backend)
            runs.append(
                {
                    "run_id": f"read-r{repetition:02d}-o{order_index}-{backend}",
                    "backend": backend,
                    "status": "ok",
                    "report": report,
                    "semantic_signature": driver._signature(report, backend),
                    "process_wall_seconds": 0.1,
                    "sampled_peak_rss_bytes": 4096,
                    "last_sampled_cpu_seconds": 0.05,
                }
            )
    return runs


def _assert_not_comparable(summary):
    assert summary["passed"] is False
    assert summary["semantic_comparable"] is False
    assert summary["statistics"] == {}


def test_integral_float_type_differences_do_not_change_signature(driver):
    native = _report(driver, "native")
    calamine = _report(driver, "calamine")

    assert native["sheets"][0]["types"] != calamine["sheets"][0]["types"]
    assert driver._signature(native, "native") == driver._signature(
        calamine, "calamine"
    )
    assert native["sheets"][0]["types"]["int"] == 2
    assert calamine["sheets"][0]["types"]["float"] == 3


def test_empty_storage_differences_do_not_change_signature(driver):
    native = _report(driver, "native")
    calamine = _report(driver, "calamine")
    calamine["cells_visited"] += 2
    calamine["sheets"][0]["cells_visited"] += 2
    calamine["sheets"][0]["types"]["empty"] = 2

    assert driver._signature(native, "native") == driver._signature(
        calamine, "calamine"
    )


@pytest.mark.parametrize("backend", ["native", "calamine"])
def test_excluded_cells_are_rejected(driver, backend):
    report = _report(driver, backend)
    report["sheets"][0]["excluded_cells"] = 1

    with pytest.raises(ValueError, match="excluded cells are unsupported"):
        driver._signature(report, backend)


@pytest.mark.parametrize("backend", ["native", "calamine"])
@pytest.mark.parametrize(
    "kind", ["date", "datetime", "duration", "error", "unexpected_formula"]
)
def test_unsupported_types_are_rejected_even_without_exclusions(driver, backend, kind):
    report = _report(driver, backend)
    report["sheets"][0]["types"] = {kind: 5}

    with pytest.raises(ValueError, match="unsupported scalar type histogram"):
        driver._signature(report, backend)


@pytest.mark.parametrize(
    "field,value,message",
    [
        ("nonempty", 6, "inconsistent nonempty/checksum counts"),
        ("checksum_cells", 4, "inconsistent nonempty/checksum counts"),
        ("cells_visited", True, "must be a nonnegative integer"),
        ("nonempty_bounds_1based", [0, 1, 1, 5], "invalid one-based nonempty bounds"),
        ("nonempty_bounds_1based", [2, 1, 1, 5], "invalid one-based nonempty bounds"),
        ("scalar_checksum", "not-a-checksum", "invalid scalar_checksum"),
        ("types", {"int": 4}, "type histogram does not sum"),
    ],
)
def test_inconsistent_sheet_summaries_are_rejected(driver, field, value, message):
    report = _report(driver, "native")
    report["sheets"][0][field] = value

    with pytest.raises(ValueError, match=message):
        driver._signature(report, "native")


def test_duplicate_sheet_names_are_rejected(driver):
    report = _report(driver, "native")
    report["sheets"][1]["name"] = report["sheets"][0]["name"]

    with pytest.raises(ValueError, match="duplicate worksheet name"):
        driver._signature(report, "native")


def test_complete_matching_summary_publishes_statistics(driver, records):
    summary = driver._summary(records, 7, True)

    assert summary["complete"] is True
    assert summary["passed"] is True
    assert summary["semantic_comparable"] is True
    assert summary["expected_runs"] == summary["recorded_runs"] == 14
    assert summary["valid_runs_per_backend"] == {"native": 7, "calamine": 7}
    assert summary["mismatches"] == summary["failures"] == []
    for backend in ("native", "calamine"):
        stats = summary["statistics"][backend]
        assert stats["sampled_peak_rss_bytes"] == {
            "samples": 7,
            "median": 4096,
            "min": 4096,
            "max": 4096,
        }
        assert stats["timing_ns"]["backend_total_with_drop"]["median"] == 100


@pytest.mark.parametrize(
    "field,value",
    [
        ("name", "Different sheet"),
        ("nonempty_bounds_1based", [2, 1, 2, 5]),
        ("nonempty", 4),
        ("checksum_cells", 4),
        ("scalar_checksum", "fedcba9876543210"),
    ],
)
def test_semantic_mismatch_blocks_comparability_and_statistics(
    driver, records, field, value
):
    records[-1]["semantic_signature"][0][field] = value

    summary = driver._summary(records, 7, True)

    assert summary["complete"] is True
    _assert_not_comparable(summary)
    assert summary["mismatches"] == [
        {"run_id": records[-1]["run_id"], "reference_run_id": records[0]["run_id"]}
    ]


def test_sheet_order_mismatch_blocks_comparability(driver, records):
    records[-1]["semantic_signature"].reverse()

    summary = driver._summary(records, 7, True)

    _assert_not_comparable(summary)
    assert summary["mismatches"]


@pytest.mark.parametrize("empty", [False, True])
def test_incomplete_runs_block_comparability(driver, records, empty):
    records = [] if empty else records[:-1]

    summary = driver._summary(records, 7, True)

    assert summary["complete"] is False
    _assert_not_comparable(summary)


@pytest.mark.parametrize("status", ["error", "timeout"])
def test_failed_runs_block_comparability(driver, records, status):
    records[-1].update(status=status, error="Synthetic child failure")

    summary = driver._summary(records, 7, True)

    assert summary["complete"] is False
    _assert_not_comparable(summary)
    assert summary["failures"] == [
        {
            "run_id": records[-1]["run_id"],
            "status": status,
            "error": "Synthetic child failure",
        }
    ]


def test_seven_valid_runs_per_backend_are_required(driver, records):
    records[-1]["backend"] = "native"

    summary = driver._summary(records, 7, True)

    assert summary["recorded_runs"] == 14
    assert summary["valid_runs_per_backend"] == {"native": 8, "calamine": 6}
    assert summary["complete"] is False
    _assert_not_comparable(summary)


def test_changed_artifact_hashes_block_comparability(driver, records):
    summary = driver._summary(records, 7, False)

    assert summary["complete"] is True
    assert summary["artifact_hashes_unchanged"] is False
    _assert_not_comparable(summary)


def test_scheduling_error_blocks_comparability(driver, records):
    summary = driver._summary(
        records, 7, True, scheduling_error="Parent deadline reached"
    )

    assert summary["complete"] is False
    assert summary["scheduling_error"] == "Parent deadline reached"
    _assert_not_comparable(summary)


def test_cpu_samples_are_optional(driver, records):
    for record in records:
        record["last_sampled_cpu_seconds"] = None

    summary = driver._summary(records, 7, True)

    assert summary["passed"] is True
    for backend in ("native", "calamine"):
        assert summary["statistics"][backend]["last_sampled_cpu_seconds"] is None


def test_orders_rotate_and_include_each_backend_once_per_repetition(driver):
    orders = list(driver._orders(7))

    assert [repetition for repetition, _ in orders] == list(range(1, 8))
    assert [order for _, order in orders] == [
        ("native", "calamine"),
        ("calamine", "native"),
        ("native", "calamine"),
        ("calamine", "native"),
        ("native", "calamine"),
        ("calamine", "native"),
        ("native", "calamine"),
    ]


def _cli_arguments(driver):
    path = str(Path(driver.__file__).resolve())
    return ["--executable", path, "--input", path]


def test_cli_defaults_to_seven_repetitions(driver):
    args = driver._parser().parse_args(_cli_arguments(driver))

    assert args.repetitions == 7


@pytest.mark.parametrize("repetitions", [-1, 0, 6])
def test_cli_rejects_fewer_than_seven_repetitions(driver, capsys, repetitions):
    with pytest.raises(SystemExit) as exc:
        driver.main([*_cli_arguments(driver), f"--repetitions={repetitions}"])

    assert exc.value.code == 2
    assert "At least seven fresh processes per backend" in capsys.readouterr().err


@pytest.mark.parametrize(
    "option", ["--timeout", "--parent-timeout", "--sample-interval"]
)
@pytest.mark.parametrize("value", ["0", "-1", "nan", "inf", "-inf"])
def test_cli_requires_finite_positive_time_limits(driver, capsys, option, value):
    with pytest.raises(SystemExit) as exc:
        driver.main([*_cli_arguments(driver), f"{option}={value}"])

    assert exc.value.code == 2
    assert f"{option} must be finite and positive" in capsys.readouterr().err
