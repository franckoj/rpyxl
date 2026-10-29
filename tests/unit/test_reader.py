import os
import datetime
import pytest
import rpyxl
from rpyxl import load_workbook

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "fixtures")

def test_load_workbook_missing_file():
    with pytest.raises(FileNotFoundError):
        load_workbook("non_existent_file_12345.xlsx")

def test_load_workbook_corrupted_file():
    corrupted_path = os.path.join(FIXTURES_DIR, "corrupted.xlsx")
    with pytest.raises((IOError, ValueError, Exception)):
        load_workbook(corrupted_path)

def test_load_workbook_standard_sample_types():
    path = os.path.join(FIXTURES_DIR, "sample_types.xlsx")
    wb = load_workbook(path, read_only=False)

    assert wb.read_only is False
    assert "Sheet1" in wb.sheetnames
    ws = wb["Sheet1"]

    # Check headers
    assert ws["A1"].value == "Name"
    assert ws["B1"].value == "Age"
    assert ws["C1"].value == "Score"
    assert ws["D1"].value == "Active"
    assert ws["E1"].value == "Created"

    # Check row 2 (Alice)
    assert ws["A2"].value == "Alice"
    assert ws["B2"].value == 30
    assert ws["C2"].value == 95.5
    assert ws["D2"].value is True
    assert isinstance(ws["E2"].value, datetime.datetime)
    assert ws["E2"].value == datetime.datetime(2023, 1, 15, 10, 30, 0)

    # Check formula in data_only=False
    assert ws["B5"].value.startswith("=")
    assert "SUM" in ws["B5"].value

    wb.close()

def test_load_workbook_data_only():
    path = os.path.join(FIXTURES_DIR, "sample_types.xlsx")
    wb_form = load_workbook(path, data_only=False)
    assert wb_form["Sheet1"]["B5"].value.startswith("=")
    wb_form.close()

    wb_data = load_workbook(path, data_only=True)
    assert wb_data.data_only is True
    # If not evaluated by Excel before saving, cached value may be None
    val = wb_data["Sheet1"]["B5"].value
    assert val is None or isinstance(val, (int, float))
    wb_data.close()

def test_load_workbook_multilingual_and_escaping():
    path = os.path.join(FIXTURES_DIR, "multilingual.xlsx")
    wb = load_workbook(path)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    assert rows[0] == ("Language", "Sample Text")
    assert rows[1][0] == "English"
    assert "Hello World & Friends <tag> 'quoted' \"double\"" in rows[1][1]
    assert rows[2] == ("Chinese", "你好世界，很高兴认识你")
    assert rows[3] == ("Japanese", "こんにちは世界、よろしく")
    assert rows[4] == ("Arabic", "مرحبا بالعالم")
    assert rows[5] == ("Russian", "Привет мир, добро пожаловать")
    assert "Grüße" in rows[6][1]
    assert "🚀 🦀 🐍 ⚡ 📊" in rows[7][1]
    wb.close()

def test_load_workbook_multi_sheet():
    path = os.path.join(FIXTURES_DIR, "multi_sheet.xlsx")
    wb = load_workbook(path)

    assert wb.sheetnames == ["Summary", "Data 2023", "Data 2024"]
    assert len(wb) == 3

    # Check active sheet was preserved (Data 2023)
    assert wb.active.title == "Data 2023"

    ws_2024 = wb["Data 2024"]
    assert ws_2024["A2"].value == 2024
    assert ws_2024["B2"].value == 750000
    wb.close()

def test_load_workbook_sparse():
    path = os.path.join(FIXTURES_DIR, "sparse.xlsx")
    wb = load_workbook(path)
    ws = wb.active

    assert ws["A1"].value == 100
    assert ws["B1"].value is None
    assert ws["C1"].value == 300
    assert ws["A5"].value == 500
    assert ws["D10"].value == "Far Away"
    wb.close()

def test_load_workbook_context_manager():
    path = os.path.join(FIXTURES_DIR, "sample_types.xlsx")
    with load_workbook(path) as wb:
        assert len(wb.sheetnames) >= 1

def test_streaming_read_only_mode():
    path = os.path.join(FIXTURES_DIR, "sample_types.xlsx")
    wb = load_workbook(path, read_only=True)

    assert wb.read_only is True
    assert "Sheet1" in wb.sheetnames
    ws = wb["Sheet1"]

    assert ws.title == "Sheet1"
    assert ws.min_row == 1
    assert ws.max_row >= 5
    assert ws.min_column == 1
    assert ws.max_column >= 5

    # Test single cell access
    cell_a2 = ws.cell(2, 1)
    assert cell_a2.value == "Alice"
    assert cell_a2.coordinate == "A2"
    assert cell_a2.row == 2
    assert cell_a2.column == 1
    assert cell_a2.data_type == "s"

    cell_bracket = ws["A2"]
    assert cell_bracket.value == "Alice"

    # Test values_only streaming iteration
    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) >= 5
    assert rows[0][0] == "Name"
    assert rows[1][0] == "Alice"
    assert rows[1][1] == 30
    assert rows[1][2] == 95.5
    assert rows[1][3] is True

    # Test cell objects iteration
    cell_rows = list(ws.iter_rows(values_only=False))
    assert len(cell_rows) >= 5
    assert cell_rows[0][0].value == "Name"
    assert cell_rows[1][0].coordinate == "A2"

    # Test rows and values generator properties
    gen_values = list(ws.values)
    assert gen_values[1][0] == "Alice"

    gen_rows = list(ws.rows)
    assert gen_rows[1][0].value == "Alice"

    # Verify read-only mutation protection
    with pytest.raises(TypeError):
        ws.append(["Invalid"])

    with pytest.raises(TypeError):
        ws["A1"] = "Cannot Change"

    with pytest.raises((TypeError, AttributeError)):
        cell_a2.value = "New"


    wb.close()

def test_streaming_sparse_sheet():
    path = os.path.join(FIXTURES_DIR, "sparse.xlsx")
    with load_workbook(path, read_only=True) as wb:
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        assert len(rows) == 10
        # Row 1: A1=100, B1=None, C1=300, D1=None
        assert rows[0][0] == 100
        assert rows[0][1] is None
        assert rows[0][2] == 300
        # Row 5: A5=500
        assert rows[4][0] == 500
        # Row 10: D10="Far Away"
        assert rows[9][3] == "Far Away"
