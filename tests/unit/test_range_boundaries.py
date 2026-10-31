"""Regression tests for range separator validation."""

import pytest

from rpyxl.utils import range_boundaries


@pytest.mark.parametrize(
    "range_string",
    ["A1:B2:C3", "A1:B2:", "A1:B2:garbage", "A1:B2::", "A1::B2"],
)
def test_range_boundaries_rejects_extra_separators(range_string):
    with pytest.raises(ValueError):
        range_boundaries(range_string)


@pytest.mark.parametrize("range_string", ["", ":", "A1:", ":B2"])
def test_range_boundaries_rejects_missing_endpoints(range_string):
    with pytest.raises(ValueError):
        range_boundaries(range_string)


@pytest.mark.parametrize(
    "range_string, expected",
    [
        ("A1", (1, 1, 1, 1)),
        ("A1:B2", (1, 1, 2, 2)),
        ("B2:A1", (1, 1, 2, 2)),
        (" $A$1:$B$2 ", (1, 1, 2, 2)),
    ],
)
def test_range_boundaries_preserves_valid_ranges(range_string, expected):
    assert range_boundaries(range_string) == expected
