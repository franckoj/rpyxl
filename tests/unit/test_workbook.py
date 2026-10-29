"""Unit tests for RPyXL Workbook operations."""

import pytest
from rpyxl import Workbook


def test_workbook_initialization():
    wb = Workbook()
    assert len(wb) == 1
    assert wb.sheetnames == ["Sheet"]
    assert wb.active.title == "Sheet"


def test_create_sheets():
    wb = Workbook()

    # Append custom named sheet
    ws2 = wb.create_sheet(title="Invoices")
    assert ws2.title == "Invoices"
    assert wb.sheetnames == ["Sheet", "Invoices"]

    # Insert at beginning
    ws0 = wb.create_sheet(title="Cover", index=0)
    assert ws0.title == "Cover"
    assert wb.sheetnames == ["Cover", "Sheet", "Invoices"]

    # Default name generation
    ws_auto = wb.create_sheet()
    assert ws_auto.title == "Sheet1"
    assert wb.sheetnames == ["Cover", "Sheet", "Invoices", "Sheet1"]


def test_duplicate_sheet_rejected():
    wb = Workbook()
    with pytest.raises(ValueError):
        wb.create_sheet(title="Sheet")

    with pytest.raises(ValueError):
        wb.create_sheet(title="sheet")  # Case-insensitive collision check


def test_remove_sheets():
    wb = Workbook()
    ws2 = wb.create_sheet(title="Sheet2")
    assert len(wb) == 2

    # Remove by worksheet object
    wb.remove(ws2)
    assert wb.sheetnames == ["Sheet"]

    # Cannot remove last sheet
    with pytest.raises(ValueError):
        wb.remove(wb.active)


def test_del_sheet_by_name():
    wb = Workbook()
    wb.create_sheet(title="Sheet2")
    assert "Sheet2" in wb

    del wb["Sheet2"]
    assert "Sheet2" not in wb
    assert wb.sheetnames == ["Sheet"]


def test_workbook_indexing_and_contains():
    wb = Workbook()
    wb.create_sheet(title="Data")

    assert "Data" in wb
    assert "Unknown" not in wb

    ws = wb["Data"]
    assert ws.title == "Data"

    with pytest.raises(KeyError):
        _ = wb["NonExistent"]


def test_workbook_iteration():
    wb = Workbook()
    wb.create_sheet(title="SheetA")
    wb.create_sheet(title="SheetB")

    names = [ws.title for ws in wb]
    assert names == ["Sheet", "SheetA", "SheetB"]


def test_active_sheet_assignment():
    wb = Workbook()
    ws2 = wb.create_sheet(title="Second")

    # Set by Worksheet object
    wb.active = ws2
    assert wb.active.title == "Second"

    # Set by integer index
    wb.active = 0
    assert wb.active.title == "Sheet"

    # Set by title string
    wb.active = "Second"
    assert wb.active.title == "Second"

    # Out of range
    with pytest.raises(IndexError):
        wb.active = 99
