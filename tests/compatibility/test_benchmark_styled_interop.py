import io

import openpyxl
import pytest
import rpyxl
from rpyxl.styles import Alignment, Font, PatternFill


def test_styled_write_roundtrip_matches_openpyxl():
    """Write styled sheet with benchmark's 5 combinations + freeze_panes, verify openpyxl reads identically."""
    wb = rpyxl.Workbook()
    ws = wb.active
    ws.title = "People"

    header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    body_font = Font(name="Arial", size=11)
    stripe_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    center = Alignment(horizontal="center")

    headers = ["Name", "Age", "City", "Job Title", "Email"]
    ws.append(headers)

    for j in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=j)
        c.font = header_font
        c.fill = header_fill
        if j == 2:
            c.alignment = center

    data = [
        ["Sofia Sharma", 39, "Bengaluru", "Software Engineer", "sofia.sharma@example.com"],
        ["Noah Garcia", 56, "Pune", "Nurse", "noah.garcia@example.com"],
        ["Olivia Patel", 23, "Pune", "Designer", "olivia.patel@example.com"],
        ["Lucas Smith", 45, "Delhi", "Product Manager", "lucas.smith@example.com"],
    ]

    for i, row in enumerate(data, start=2):
        stripe = i % 2 == 0
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=val)
            c.font = body_font
            if stripe:
                c.fill = stripe_fill
            if j == 2:
                c.alignment = center

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Read back with openpyxl
    loaded_wb = openpyxl.load_workbook(buf)
    loaded_ws = loaded_wb["People"]

    assert loaded_ws.freeze_panes == "A2"
    assert loaded_ws.sheet_view.pane is not None
    assert loaded_ws.sheet_view.pane.state == "frozen"
    assert loaded_ws.sheet_view.pane.ySplit == 1

    # Check header row (row 1)
    for j in range(1, len(headers) + 1):
        cell = loaded_ws.cell(row=1, column=j)
        assert cell.font.name == "Arial"
        assert cell.font.size == 12
        assert cell.font.bold is True
        assert cell.fill.fill_type == "solid"
        if j == 2:
            assert cell.alignment.horizontal == "center"

    # Check data rows
    for i, row in enumerate(data, start=2):
        stripe = i % 2 == 0
        for j, val in enumerate(row, start=1):
            cell = loaded_ws.cell(row=i, column=j)
            assert cell.value == val
            assert cell.font.name == "Arial"
            assert cell.font.size == 11
            if stripe:
                assert cell.fill.fill_type == "solid"
            if j == 2:
                assert cell.alignment.horizontal == "center"


@pytest.mark.parametrize(
    "coord,expected_x,expected_y",
    [
        ("A2", 0, 1),
        ("B1", 1, 0),
        ("B2", 1, 1),
        ("C3", 2, 2),
    ],
)
def test_freeze_panes_coordinates_roundtrip(coord, expected_x, expected_y):
    wb = rpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = "H2"
    ws["A2"] = "D1"
    ws["B2"] = "D2"
    ws.freeze_panes = coord

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    loaded = openpyxl.load_workbook(buf)
    loaded_ws = loaded.active

    assert loaded_ws.freeze_panes == coord
    pane = loaded_ws.sheet_view.pane
    assert pane is not None
    assert pane.state == "frozen"
    if expected_x > 0:
        assert pane.xSplit == expected_x
    if expected_y > 0:
        assert pane.ySplit == expected_y


def test_freeze_panes_none_or_a1():
    for coord in (None, "A1"):
        wb = rpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Value"
        ws.freeze_panes = coord

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        loaded = openpyxl.load_workbook(buf)
        loaded_ws = loaded.active
        assert loaded_ws.freeze_panes is None


def test_style_mutation_via_view():
    wb = rpyxl.Workbook()
    ws = wb.active
    c = ws.cell(row=1, column=1, value="Hello")
    c.font = Font(name="Arial", size=11, bold=False)
    assert c.font.bold is False

    # Mutating the view reflects on the cell
    c.font.bold = True
    assert c.font.bold is True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    loaded = openpyxl.load_workbook(buf).active
    assert loaded["A1"].font.bold is True
