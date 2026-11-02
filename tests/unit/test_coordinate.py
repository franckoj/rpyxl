"""Regression tests for the public coordinate utilities."""

import pytest

from rpyxl.utils import coordinate_to_tuple, range_boundaries


@pytest.mark.parametrize(
    "coordinate, expected",
    [
        ("A1", (1, 1)),
        ("AA100", (100, 27)),
        ("xfd1048576", (1048576, 16384)),
        ("$B$10", (10, 2)),
        ("$B10", (10, 2)),
        ("B$10", (10, 2)),
        ("  c3  ", (3, 3)),
        ("A001", (1, 1)),
    ],
)
def test_coordinate_to_tuple_valid(coordinate, expected):
    assert coordinate_to_tuple(coordinate) == expected


@pytest.mark.parametrize(
    "coordinate",
    [
        "A1B2", "1A", "A1B", "1A2", "A1$2", "A$B1", "$$A1", "A$$1",
        "A1$", "A\u0661", "\u212a1", "\u017f1", "A 1", "A1:B2", "", "A", "1",
    ],
)
def test_coordinate_to_tuple_rejects_malformed_references(coordinate):
    with pytest.raises(ValueError, match="Invalid coordinate format"):
        coordinate_to_tuple(coordinate)


@pytest.mark.parametrize("coordinate", ["A0", "A1048577", "XFE1"])
def test_coordinate_to_tuple_rejects_out_of_bounds(coordinate):
    with pytest.raises(ValueError, match="out of bounds"):
        coordinate_to_tuple(coordinate)


@pytest.mark.parametrize("range_string", ["A1B2:C3", "A1:1B2", "A1:A\u0661"])
def test_range_boundaries_rejects_malformed_endpoints(range_string):
    with pytest.raises(ValueError, match="Invalid coordinate format"):
        range_boundaries(range_string)


def test_range_boundaries_preserves_absolute_references_and_reversed_bounds():
    assert range_boundaries("$C$5:$A$1") == (1, 1, 3, 5)
