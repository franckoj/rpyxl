"""Unit tests for RPyXL High-Performance Writer (Standard & Write-Only Modes)."""

import datetime
from pathlib import Path
import pytest
import zipfile

from rpyxl import Workbook, load_workbook


class TestStandardWriter:
    def test_save_creates_valid_zip(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Hello"
        file_path = tmp_path / "valid.xlsx"
        wb.save(file_path)

        assert file_path.exists()
        assert zipfile.is_zipfile(file_path)

        with zipfile.ZipFile(file_path, "r") as z:
            names = z.namelist()
            assert "[Content_Types].xml" in names
            assert "_rels/.rels" in names
            assert "xl/workbook.xml" in names
            assert "xl/styles.xml" in names
            assert "xl/worksheets/sheet1.xml" in names

    def test_save_and_reload_primitive_types(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Text Value"
        ws["A2"] = 42
        ws["A3"] = 3.14159
        ws["A4"] = True
        ws["A5"] = False

        file_path = tmp_path / "types.xlsx"
        wb.save(str(file_path))

        reloaded = load_workbook(file_path)
        r_ws = reloaded.active
        assert r_ws["A1"].value == "Text Value"
        assert r_ws["A2"].value == 42
        assert abs(r_ws["A3"].value - 3.14159) < 1e-5
        assert r_ws["A4"].value is True
        assert r_ws["A5"].value is False

    def test_save_and_reload_dates_and_times(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        dt = datetime.datetime(2023, 10, 15, 14, 30, 0)
        d = datetime.date(2024, 1, 1)

        ws["A1"] = dt
        ws["A2"] = d

        file_path = tmp_path / "dates.xlsx"
        wb.save(file_path)

        reloaded = load_workbook(file_path)
        r_ws = reloaded.active
        assert isinstance(r_ws["A1"].value, (datetime.datetime, datetime.date))
        assert isinstance(r_ws["A2"].value, (datetime.datetime, datetime.date))

    def test_save_and_reload_formula(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 10
        ws["A2"] = 20
        ws["A3"] = "=SUM(A1:A2)"

        file_path = tmp_path / "formula.xlsx"
        wb.save(file_path)

        reloaded = load_workbook(file_path, data_only=False)
        r_ws = reloaded.active
        assert r_ws["A3"].value == "=SUM(A1:A2)"

    def test_save_multiple_sheets(self, tmp_path):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "FirstSheet"
        ws1["A1"] = "Data1"

        ws2 = wb.create_sheet("SecondSheet")
        ws2["B2"] = "Data2"

        file_path = tmp_path / "multi_sheet.xlsx"
        wb.save(file_path)

        reloaded = load_workbook(file_path)
        assert reloaded.sheetnames == ["FirstSheet", "SecondSheet"]
        assert reloaded["FirstSheet"]["A1"].value == "Data1"
        assert reloaded["SecondSheet"]["B2"].value == "Data2"

    def test_save_special_xml_characters(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "AT&T <telecom> \"corp\" 'quoted'"
        ws["A2"] = "Multilingual: 日本語, 中文, العربية, 🚀"

        file_path = tmp_path / "special_chars.xlsx"
        wb.save(file_path)

        reloaded = load_workbook(file_path)
        assert reloaded.active["A1"].value == "AT&T <telecom> \"corp\" 'quoted'"
        assert reloaded.active["A2"].value == "Multilingual: 日本語, 中文, العربية, 🚀"


class TestWriteOnlyWriter:
    def test_write_only_initial_state(self):
        wb = Workbook(write_only=True)
        assert wb.write_only is True
        assert wb.read_only is False
        assert wb.sheetnames == []

    def test_create_sheet_in_write_only(self):
        wb = Workbook(write_only=True)
        ws1 = wb.create_sheet()
        assert ws1.title == "Sheet1"
        assert wb.sheetnames == ["Sheet1"]

        ws2 = wb.create_sheet("CustomData")
        assert ws2.title == "CustomData"
        assert wb.sheetnames == ["Sheet1", "CustomData"]

    def test_write_only_unsupported_operations(self):
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Test")

        # Cell indexing should raise TypeError
        with pytest.raises(TypeError, match="not subscriptable"):
            _ = ws["A1"]

        # cell() method should raise AttributeError
        with pytest.raises(AttributeError, match="no attribute 'cell'"):
            ws.cell(1, 1)

        # iter_rows() should raise AttributeError
        with pytest.raises(AttributeError, match="no attribute 'iter_rows'"):
            ws.iter_rows()

    def test_write_only_append_and_save(self, tmp_path):
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Metrics")

        ws.append(["Metric", "Score", "Active"])
        ws.append(["CPU", 4.5, True])
        ws.append(["RAM", 12.8, False])

        file_path = tmp_path / "streaming.xlsx"
        wb.save(file_path)
        assert file_path.exists()

        reloaded = load_workbook(file_path)
        assert reloaded.sheetnames == ["Metrics"]
        rows = list(reloaded["Metrics"].iter_rows(values_only=True))
        assert rows[0] == ("Metric", "Score", "Active")
        assert rows[1] == ("CPU", 4.5, True)
        assert rows[2] == ("RAM", 12.8, False)

    def test_write_only_context_manager(self, tmp_path):
        file_path = tmp_path / "ctx_mgr.xlsx"
        with Workbook(write_only=True) as wb:
            ws = wb.create_sheet("AutoClose")
            ws.append([10, 20, 30])
            wb.save(file_path)

        assert file_path.exists()
        reloaded = load_workbook(file_path)
        assert list(reloaded["AutoClose"].iter_rows(values_only=True))[0] == (10, 20, 30)

    def test_write_only_empty_save_error(self, tmp_path):
        wb = Workbook(write_only=True)
        file_path = tmp_path / "empty_wo.xlsx"
        with pytest.raises(ValueError, match="Cannot save empty write-only workbook"):
            wb.save(file_path)

    def test_isinstance_check(self):
        wb_std = Workbook()
        wb_wo = Workbook(write_only=True)
        assert isinstance(wb_std, Workbook)
        assert isinstance(wb_wo, Workbook)
