"""Unit tests for RPyXL Cell operations."""

import pytest
from rpyxl import Workbook


def test_cell_coordinates_and_properties():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Test"
    cell = ws["A1"]

    assert cell.row == 1
    assert cell.column == 1
    assert cell.coordinate == "A1"
    assert cell.column_letter == "A"
    assert cell.data_type == "s"
    assert repr(cell) == "<Cell 'Sheet'.A1>"


def test_cell_values_various_types():
    wb = Workbook()
    ws = wb.active

    # String
    ws["A1"] = "Hello World"
    assert ws["A1"].value == "Hello World"
    assert ws["A1"].data_type == "s"

    # Integer
    ws["A2"] = 42
    assert ws["A2"].value == 42
    assert isinstance(ws["A2"].value, int)
    assert ws["A2"].data_type == "n"

    # Float
    ws["A3"] = 3.14159
    assert ws["A3"].value == pytest.approx(3.14159)
    assert isinstance(ws["A3"].value, float)
    assert ws["A3"].data_type == "n"

    # Boolean
    ws["A4"] = True
    assert ws["A4"].value is True
    assert isinstance(ws["A4"].value, bool)
    assert ws["A4"].data_type == "b"

    ws["A5"] = False
    assert ws["A5"].value is False
    assert isinstance(ws["A5"].value, bool)
    assert ws["A5"].data_type == "b"

    # None
    ws["A6"] = None
    assert ws["A6"].value is None

    # Formula
    ws["A7"] = "=SUM(A2:A3)"
    assert ws["A7"].value == "=SUM(A2:A3)"
    assert ws["A7"].data_type == "f"


def test_cell_mutation_and_reference_integrity():
    wb = Workbook()
    ws = wb.active

    ws["B2"] = 100
    cell = ws["B2"]
    assert cell.value == 100

    # Modify through cell object
    cell.value = 200
    assert ws["B2"].value == 200

    # Modify through worksheet
    ws["B2"] = 300
    assert cell.value == 300


def test_cell_bounds_and_invalid_coordinates():
    wb = Workbook()
    ws = wb.active

    # Valid limits
    ws.cell(row=1, column=1)
    ws.cell(row=1048576, column=16384)

    # Out of bounds row
    with pytest.raises(ValueError):
        ws.cell(row=0, column=1)

    with pytest.raises(ValueError):
        ws.cell(row=1048577, column=1)

    # Out of bounds col
    with pytest.raises(ValueError):
        ws.cell(row=1, column=0)

    with pytest.raises(ValueError):
        ws.cell(row=1, column=16385)

    # Invalid string coordinate
    with pytest.raises(ValueError):
        _ = ws["INVALID"]

    with pytest.raises(ValueError):
        _ = ws[""]
