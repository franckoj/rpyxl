"""Real-native contracts for conservative structural package preservation."""

from io import BytesIO
from zipfile import ZipFile

import openpyxl
import pytest
from rpyxl import NativeWorkbook, _metadata, load_workbook
from rpyxl.comments import Comment
from rpyxl.styles import Font


def _source():
    book = openpyxl.Workbook()
    sheet = book.active
    for row in range(1, 13):
        sheet.append([row, f" row {row} ", row / 2, row % 2 == 0])
    sheet["A3"].font = openpyxl.styles.Font(bold=True, color="FF112233")
    book.properties.title = "Preserve original properties"
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def _load(data):
    assert hasattr(NativeWorkbook, "_enable_plain_structure"), (
        "Rebuild the native wheel"
    )
    return load_workbook(BytesIO(data))


def _saved(book):
    output = BytesIO()
    book.save(output)
    return output.getvalue()


def _forbid(*args, **kwargs):
    raise AssertionError("Proven plain structural edits must stay native")


def _edit(sheet, operation):
    if operation == "insert_rows":
        sheet.insert_rows(2, 3)
    elif operation == "delete_rows":
        sheet.delete_rows(2, 3)
    elif operation == "insert_cols":
        sheet.insert_cols(2, 2)
    elif operation == "delete_cols":
        sheet.delete_cols(2, 2)
    else:
        sheet.move_range("A2:C5", rows=2, cols=1)


def _contents(sheet):
    return tuple(tuple(row) for row in sheet.iter_rows(values_only=True))


@pytest.mark.parametrize(
    "operation",
    ["insert_rows", "delete_rows", "insert_cols", "delete_cols", "move_range"],
)
@pytest.mark.parametrize("source_kind", ["path", "stream"])
def test_plain_operations_avoid_python_metadata_and_preserve_output(
    operation, source_kind, tmp_path, monkeypatch
):
    data = _source()
    source = BytesIO(data)
    if source_kind == "path":
        source = tmp_path / "input.xlsx"
        source.write_bytes(data)
    book = _load(data) if source_kind == "stream" else load_workbook(source)
    assert hasattr(book._native, "_enable_plain_structure"), "Rebuild the native wheel"
    control = openpyxl.load_workbook(BytesIO(data))
    try:
        monkeypatch.setattr(_metadata, "load_metadata", _forbid)
        monkeypatch.setattr(_metadata, "apply_metadata", _forbid)
        _edit(book.active, operation)
        _edit(control.active, operation)
        assert book._book is None
        assert book._native_structure_only
        assert book._native._can_preserve_package()
        expected = _contents(control.active)
        for _ in range(2):
            saved = _saved(book)
            result = openpyxl.load_workbook(BytesIO(saved))
            assert _contents(result.active) == expected
            assert result.properties.title == control.properties.title
            for row in control.active:
                for cell in row:
                    other = result.active.cell(cell.row, cell.column)
                    assert other.font.bold == cell.font.bold
                    assert other.number_format == cell.number_format
            result.close()
            with ZipFile(BytesIO(data)) as before, ZipFile(BytesIO(saved)) as after:
                assert before.namelist() == after.namelist()
                for name in before.namelist():
                    if name not in ("xl/styles.xml", "xl/worksheets/sheet1.xml"):
                        assert before.read(name) == after.read(name), name
    finally:
        book.close()
        control.close()


def test_repeated_structural_edits_then_rich_metadata_use_fallback(
    tmp_path, monkeypatch
):
    data = _source()
    book = _load(data)
    control = openpyxl.load_workbook(BytesIO(data))
    calls = []
    original = _metadata.apply_metadata

    def overlay(*args, **kwargs):
        calls.append(True)
        return original(*args, **kwargs)

    monkeypatch.setattr(_metadata, "apply_metadata", overlay)
    try:
        for operation in ("insert_rows", "delete_rows", "insert_cols", "delete_cols"):
            _edit(book.active, operation)
            _edit(control.active, operation)
        assert book._book is None
        path = tmp_path / "same.xlsx"
        book.save(path)
        book.save(path)
        assert not calls
        book.active["A1"].comment = Comment("After edits", "Reviewer")
        book.active["A1"].font = Font(italic=True)
        result = openpyxl.load_workbook(BytesIO(_saved(book)))
        assert len(calls) == 1
        assert _contents(result.active) == _contents(control.active)
        assert result.active["A1"].comment.text == "After edits"
        assert result.active["A1"].font.italic
        assert result.properties.title == "Preserve original properties"
        result.close()
    finally:
        book.close()
        control.close()


@pytest.mark.parametrize(
    "feature", ["dimension", "comment", "freeze", "formula", "merge"]
)
def test_rich_sources_do_not_opt_in(feature):
    control = openpyxl.load_workbook(BytesIO(_source()))
    if feature == "dimension":
        control.active.row_dimensions[3].height = 32
    elif feature == "comment":
        control.active["A3"].comment = openpyxl.comments.Comment("Keep", "Reviewer")
    elif feature == "freeze":
        control.active.freeze_panes = "B2"
    elif feature == "formula":
        control.active["A3"] = "=1+1"
    else:
        control.active.merge_cells("A3:B3")
    output = BytesIO()
    control.save(output)
    control.close()
    book = _load(output.getvalue())
    try:
        assert not book._try_native_structure()
        assert not book._native_structure_only
        if feature in ("freeze", "merge"):
            with pytest.raises(NotImplementedError):
                book.active.insert_rows(2)
        else:
            book.active.insert_rows(2)
            result = openpyxl.load_workbook(BytesIO(_saved(book)))
            if feature == "dimension":
                assert result.active.row_dimensions[4].height == 32
            elif feature == "comment":
                assert result.active["A4"].comment.text == "Keep"
            else:
                assert result.active["A4"].value == "=1+1"
            result.close()
        assert not book._native_structure_only
    finally:
        book.close()


def test_failed_bounds_and_zero_edits_do_not_change_cell_values():
    book = _load(_source())
    try:
        original = _contents(book.active)
        book.active.insert_rows(2, 0)
        with pytest.raises(ValueError):
            book.active.insert_rows(2, 1048576)
        assert _contents(book.active) == original
        result = openpyxl.load_workbook(BytesIO(_saved(book)))
        assert _contents(result.active) == original
        result.close()
    finally:
        book.close()
