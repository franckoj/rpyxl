"""Unit tests for RPyXL Worksheet operations."""

import pytest
from rpyxl import Workbook


def test_worksheet_title_and_renaming():
    wb = Workbook()
    ws = wb.active
    assert ws.title == "Sheet"

    ws.title = "SalesData"
    assert ws.title == "SalesData"
    assert wb.sheetnames == ["SalesData"]

    # Invalid title characters
    for invalid_char in [":", "\\", "/", "?", "*", "[", "]"]:
        with pytest.raises(ValueError):
            ws.title = f"Bad{invalid_char}Name"

    # Title exceeding 31 chars
    with pytest.raises(ValueError):
        ws.title = "A" * 32

    # Title starting or ending with apostrophe
    with pytest.raises(ValueError):
        ws.title = "'Quoted'"


def test_worksheet_dimensions_and_bounds():
    wb = Workbook()
    ws = wb.active

    assert ws.min_row == 1
    assert ws.max_row == 1
    assert ws.min_column == 1
    assert ws.max_column == 1
    assert ws.dimensions == "A1"

    ws["B2"] = 10
    assert ws.min_row == 2
    assert ws.max_row == 2
    assert ws.min_column == 2
    assert ws.max_column == 2
    assert ws.dimensions == "B2"

    ws["D5"] = 20
    assert ws.min_row == 2
    assert ws.max_row == 5
    assert ws.min_column == 2
    assert ws.max_column == 4
    assert ws.dimensions == "B2:D5"


def test_worksheet_append():
    wb = Workbook()
    ws = wb.active

    row1 = ["INV001", 1500.50, True]
    row2 = ["INV002", 2300.00, False]

    ws.append(row1)
    ws.append(row2)

    assert ws.max_row == 2
    assert ws.max_column == 3
    assert ws["A1"].value == "INV001"
    assert ws["B1"].value == 1500.50
    assert ws["C1"].value is True
    assert ws["A2"].value == "INV002"
    assert ws["B2"].value == 2300.00
    assert ws["C2"].value is False


def test_iter_rows_values_only():
    wb = Workbook()
    ws = wb.active

    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    ws.append([7, 8, 9])

    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) == 3
    assert rows[0] == (1, 2, 3)
    assert rows[1] == (4, 5, 6)
    assert rows[2] == (7, 8, 9)

    # Sub-range
    sub_rows = list(ws.iter_rows(min_row=1, max_row=2, min_col=2, max_col=3, values_only=True))
    assert len(sub_rows) == 2
    assert sub_rows[0] == (2, 3)
    assert sub_rows[1] == (5, 6)


def test_iter_rows_cells():
    wb = Workbook()
    ws = wb.active

    ws.append(["A", "B"])
    ws.append(["C", "D"])

    rows = list(ws.iter_rows(values_only=False))
    assert len(rows) == 2
    assert len(rows[0]) == 2
    assert rows[0][0].value == "A"
    assert rows[0][1].value == "B"
    assert rows[1][0].value == "C"
    assert rows[1][1].value == "D"


def test_iter_cols():
    wb = Workbook()
    ws = wb.active

    ws.append([1, 2])
    ws.append([3, 4])

    cols_values = list(ws.iter_cols(values_only=True))
    assert len(cols_values) == 2
    assert cols_values[0] == (1, 3)
    assert cols_values[1] == (2, 4)


def test_properties_values_rows_columns():
    wb = Workbook()
    ws = wb.active
    ws.append([10, 20])
    ws.append([30, 40])

    assert list(ws.values) == [(10, 20), (30, 40)]
    assert len(list(ws.rows)) == 2
    assert len(list(ws.columns)) == 2


def test_range_slice():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["B1"] = 2
    ws["A2"] = 3
    ws["B2"] = 4

    slice_res = ws["A1:B2"]
    assert len(slice_res) == 2
    assert len(slice_res[0]) == 2
    assert slice_res[0][0].value == 1
    assert slice_res[0][1].value == 2
    assert slice_res[1][0].value == 3
    assert slice_res[1][1].value == 4
