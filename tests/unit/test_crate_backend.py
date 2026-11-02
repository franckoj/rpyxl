"""Public opt-in backend contracts; candidate tests require the rebuilt native ABI.

Do not skip or replace the candidate writer when testing an older installed wheel.
The runtime fixture deliberately fails with a rebuild instruction in that case.
See crates/rpyxl-ooxml/RUST_XLSXWRITER.md for the strict supported subset.
"""

from datetime import date, datetime, time
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import openpyxl
import pytest
from rpyxl import NativeWorkbook, Workbook, load_workbook
from rpyxl.comments import Comment
from rpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side

CANDIDATE = "rust_xlsxwriter"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@pytest.fixture
def candidate():
    assert callable(getattr(NativeWorkbook, "_save_rust_xlsxwriter", None)), (
        "Candidate runtime tests require the real _save_rust_xlsxwriter ABI; "
        "rebuild and install the wheel, do not skip or mock the writer"
    )
    book = Workbook(backend=CANDIDATE)
    try:
        assert book.backend == CANDIDATE
        yield book
    finally:
        book.close()


def _save_bytes(book):
    with BytesIO() as output:
        book.save(output)
        assert not output.closed
        return output.getvalue()


def _assert_unchanged(book, tmp_path, error):
    """Check validation before touching existing or absent paths and caller streams."""
    sentinel = b"existing destination must survive" * 200
    path = tmp_path / "existing.xlsx"
    path.write_bytes(sentinel)
    with pytest.raises(error):
        book.save(path)
    assert path.read_bytes() == sentinel
    absent = tmp_path / "absent.xlsx"
    with pytest.raises(error):
        book.save(absent)
    assert not absent.exists()
    with BytesIO(sentinel) as stream:
        stream.seek(13)
        with pytest.raises(error):
            book.save(stream)
        assert not stream.closed
        assert stream.tell() == 13
        assert stream.getvalue() == sentinel


def test_default_and_explicit_native_never_call_candidate(monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail("The native/default save path called the candidate writer")

    monkeypatch.setattr(
        NativeWorkbook, "_save_rust_xlsxwriter", forbidden, raising=False
    )
    for kwargs in ({}, {"backend": "native"}, {"write_only": True}):
        book = Workbook(**kwargs)
        try:
            assert book.backend == "native"
            sheet = book.create_sheet() if book.write_only else book.active
            sheet.append(["native", 7])
            payload = _save_bytes(book)
            loaded = load_workbook(BytesIO(payload))
            try:
                assert loaded.backend == "native"
                assert list(loaded.active.values) == [("native", 7)]
            finally:
                loaded.close()
        finally:
            book.close()


@pytest.mark.parametrize("backend", ["", "xlsxwriter", "RUST_XLSXWRITER", None, 1])
def test_unknown_backend_is_rejected(backend):
    with pytest.raises(ValueError):
        Workbook(backend=backend)


def test_backend_selection_is_keyword_only():
    with pytest.raises(TypeError):
        Workbook(False, CANDIDATE)


def test_candidate_write_only_is_explicitly_unimplemented():
    with pytest.raises(NotImplementedError):
        Workbook(write_only=True, backend=CANDIDATE)


def test_candidate_rejects_even_new_native_injection():
    native = NativeWorkbook()
    try:
        with pytest.raises(ValueError):
            Workbook(backend=CANDIDATE, _native=native)
        # Failed construction must not take ownership of the supplied workbook.
        native.active.append(["still usable"])
        assert native.active["A1"].value == "still usable"
    finally:
        native.close()


@pytest.mark.parametrize("read_only", [False, True])
@pytest.mark.parametrize("data_only", [False, True])
def test_loaded_workbooks_stay_native_and_reject_candidate_injection(
    tmp_path, monkeypatch, read_only, data_only
):
    def forbidden(*args, **kwargs):
        pytest.fail("A loaded/native save called the candidate writer")

    monkeypatch.setattr(
        NativeWorkbook, "_save_rust_xlsxwriter", forbidden, raising=False
    )
    original = Workbook()
    try:
        original.active["A1"] = "source"
        original.active["A1"].font = Font(bold=True)
        payload = _save_bytes(original)
    finally:
        original.close()
    loaded = load_workbook(BytesIO(payload), read_only=read_only, data_only=data_only)
    try:
        assert loaded.backend == "native"
        with pytest.raises(ValueError):
            Workbook(backend=CANDIDATE, _native=loaded._native)
        assert loaded.active["A1"].value == "source"
        if not read_only:
            loaded.save(tmp_path / "native-copy.xlsx")
            result = openpyxl.load_workbook(tmp_path / "native-copy.xlsx")
            try:
                assert result.active["A1"].font.bold
            finally:
                result.close()
    finally:
        loaded.close()


def test_candidate_rejects_source_only_injection_without_taking_ownership():
    original = Workbook()
    try:
        original.active["A1"] = "source"
        payload = _save_bytes(original)
    finally:
        original.close()
    with BytesIO(payload) as source:
        source.seek(7)
        with pytest.raises(ValueError):
            Workbook(backend=CANDIDATE, _source=source)
        assert not source.closed
        assert source.tell() == 7


def test_missing_candidate_abi_requires_rebuild(candidate, monkeypatch, tmp_path):
    # This is the only candidate test that removes the ABI, to test its error path.
    monkeypatch.delattr(type(candidate._native), "_save_rust_xlsxwriter")
    with pytest.raises(RuntimeError, match="(?i)rebuild|wheel"):
        candidate.save(tmp_path / "missing-abi.xlsx")
    assert not (tmp_path / "missing-abi.xlsx").exists()


@pytest.mark.parametrize("destination", ["path", "bytesio"])
def test_candidate_values_dates_unicode_and_carriage_returns(
    candidate, tmp_path, destination
):
    values = [
        None,
        True,
        False,
        0,
        -123,
        2**53,
        -(2**53),
        2**54,
        1.25,
        -0.125,
        "日本語 café 🦀",
        " leading and trailing ",
        "line\nfeed",
        "carriage\rreturn",
        "CRLF\r\nend\ttab",
        "https://example.com/not-a-hyperlink",
        date(2024, 2, 29),
        datetime(2024, 2, 29, 12, 34, 56, 123000),
    ]
    for row, value in enumerate(values, 1):
        candidate.active.cell(row, 1, value)
    candidate.active["B1"] = "=literal, not a formula"
    candidate.active["B1"].data_type = "s"
    second = candidate.create_sheet("Δεύτερο")
    second["Z9"] = "sparse"
    candidate.active = second
    target = tmp_path / "values.xlsx" if destination == "path" else BytesIO()
    try:
        candidate.save(target)
        payload = target.read_bytes() if destination == "path" else target.getvalue()
        readers = [
            openpyxl.load_workbook(BytesIO(payload)),
            load_workbook(BytesIO(payload)),
            load_workbook(BytesIO(payload), read_only=True),
        ]
        try:
            for loaded in readers:
                assert loaded.sheetnames == ["Sheet", "Δεύτερο"]
                assert loaded.active.title == "Δεύτερο"
                assert loaded["Δεύτερο"]["Z9"].value == "sparse"
                for row, value in enumerate(values, 1):
                    actual = loaded["Sheet"].cell(row, 1).value
                    if (
                        isinstance(value, date)
                        and not isinstance(value, datetime)
                        and isinstance(actual, datetime)
                    ):
                        actual = actual.date()
                    assert actual == value, (row, actual, value)
                    if isinstance(value, bool):
                        assert type(actual) is bool
                assert loaded["Sheet"]["B1"].value == "=literal, not a formula"
                assert loaded["Sheet"]["B1"].data_type != "f"
        finally:
            for loaded in readers:
                loaded.close()
        with ZipFile(BytesIO(payload)) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
            props = root.find("s:workbookPr", NS)
            assert props is None or props.get("date1904", "0") in ("0", "false")
    finally:
        if isinstance(target, BytesIO):
            assert not target.closed
            target.close()


@pytest.mark.parametrize("default_getters", [False, True])
def test_empty_candidate_and_default_metadata_getters_are_supported(
    candidate, default_getters
):
    if default_getters:
        cell = candidate.active["A1"]
        for component in ("font", "fill", "border", "alignment", "protection"):
            getattr(cell, component)
        _ = candidate.properties, candidate.defined_names
        _ = candidate.active.column_dimensions["A"], candidate.active.row_dimensions[1]
    payload = _save_bytes(candidate)
    loaded = openpyxl.load_workbook(BytesIO(payload))
    try:
        assert loaded.sheetnames == ["Sheet"]
        assert loaded.active["A1"].value is None
    finally:
        loaded.close()


def test_standalone_time_uses_supported_temporal_format(candidate):
    candidate.active["A1"] = time(12, 34, 56, 123000)
    payload = _save_bytes(candidate)
    loaded = openpyxl.load_workbook(BytesIO(payload))
    try:
        assert loaded.active["A1"].value == time(12, 34, 56, 123000)
    finally:
        loaded.close()


def test_opaque_argb_styles_and_styled_blank(candidate):
    cell = candidate.active["A1"]
    cell.value = 12.5
    cell.font = Font(name="Arial", size=12, bold=True, italic=True, color="FF123456")
    cell.fill = PatternFill("solid", fgColor="FFABCDEF")
    cell.border = Border(
        left=Side("thin", "FF102030"), bottom=Side("double", "FF405060")
    )
    cell.alignment = Alignment(
        horizontal="center", vertical="top", wrap_text=True, text_rotation=30
    )
    cell.protection = Protection(locked=False, hidden=True)
    cell.number_format = '0.0000 "units"'
    blank = candidate.active["Z20"]
    blank.font = Font(name="Arial", size=12, bold=True, color="FF998877")
    blank.number_format = "0.00"
    payload = _save_bytes(candidate)
    for reader in (openpyxl.load_workbook, load_workbook):
        loaded = reader(BytesIO(payload))
        try:
            actual = loaded.active["A1"]
            assert actual.value == 12.5
            assert actual.font.name == "Arial"
            assert actual.font.sz == 12
            assert actual.font.bold and actual.font.italic
            assert actual.font.color.rgb == "FF123456"
            assert actual.fill.patternType == "solid"
            assert actual.fill.fgColor.rgb == "FFABCDEF"
            assert actual.border.left.style == "thin"
            assert actual.border.left.color.rgb == "FF102030"
            assert actual.border.bottom.style == "double"
            assert actual.border.bottom.color.rgb == "FF405060"
            assert actual.alignment.horizontal == "center"
            assert actual.alignment.vertical == "top"
            assert actual.alignment.wrap_text
            assert actual.alignment.textRotation == 30
            assert not actual.protection.locked and actual.protection.hidden
            assert actual.number_format == '0.0000 "units"'
            assert loaded.active["Z20"].value is None
            assert loaded.active["Z20"].font.bold
            assert loaded.active["Z20"].font.color.rgb == "FF998877"
            assert loaded.active["Z20"].number_format == "0.00"
        finally:
            loaded.close()


def test_repeated_save_source_mutation_and_nested_copy_on_write(candidate):
    source = Font(name="Arial", size=12, color=Color(rgb="FF123456"))
    border = Border(left=Side("thin", "FF112233"))
    for coordinate in ("A1", "B1"):
        candidate.active[coordinate] = 1
        candidate.active[coordinate].font = source
        candidate.active[coordinate].border = border
    first = _save_bytes(candidate)
    source.color.rgb = "FF000000"
    border.left.color.rgb = "FF000000"
    candidate.active["A1"].font.color.rgb = "FFABCDEF"
    candidate.active["A1"].border.left.color.rgb = "FF445566"
    candidate.active["C1"].font = source
    second = _save_bytes(candidate)
    for payload, expected in ((first, "FF123456"), (second, "FFABCDEF")):
        loaded = openpyxl.load_workbook(BytesIO(payload))
        try:
            assert loaded.active["A1"].font.color.rgb == expected
            assert loaded.active["B1"].font.color.rgb == "FF123456"
            assert loaded.active["B1"].border.left.color.rgb == "FF112233"
            if payload is second:
                assert loaded.active["A1"].border.left.color.rgb == "FF445566"
                assert loaded.active["C1"].font.color.rgb == "FF000000"
        finally:
            loaded.close()
    candidate.active["A1"].font = None
    candidate.active["A1"].border = None
    loaded = openpyxl.load_workbook(BytesIO(_save_bytes(candidate)))
    try:
        assert not loaded.active["A1"].font.bold
        assert loaded.active["A1"].font.color != loaded.active["B1"].font.color
        assert loaded.active["A1"].border.left.style is None
        assert loaded.active["B1"].border.left.style == "thin"
    finally:
        loaded.close()


@pytest.mark.parametrize("destination", ["path", "bytesio"])
def test_repeated_save_replaces_truncates_and_retains_caller_ownership(
    candidate, tmp_path, destination
):
    stale = candidate.create_sheet("Stale")
    for row in range(64):
        stale.append([row, "obsolete contents" * 8])
    target = tmp_path / "replace.xlsx" if destination == "path" else BytesIO()
    sentinel = b"old trailing bytes" * 8192
    if destination == "path":
        target.write_bytes(sentinel)
    else:
        target.write(sentinel)
        target.seek(17)
    try:
        candidate.save(target)
        first = target.read_bytes() if destination == "path" else target.getvalue()
        candidate.remove(stale)
        candidate.active["A1"] = "shorter"
        if destination == "bytesio":
            target.seek(19)
        candidate.save(target)
        last = target.read_bytes() if destination == "path" else target.getvalue()
        assert len(last) < len(first) < len(sentinel)
        assert last.startswith(b"PK\x03\x04")
        assert last[-22:-18] == b"PK\x05\x06"  # No stale bytes after the ZIP directory.
        with ZipFile(BytesIO(last)) as archive:
            assert archive.testzip() is None
        loaded = load_workbook(BytesIO(last))
        try:
            assert loaded.sheetnames == ["Sheet"]
            assert loaded.active["A1"].value == "shorter"
            assert loaded.backend == "native"  # Saving never opts the reader in.
            with pytest.raises(ValueError):
                Workbook(backend=CANDIDATE, _native=loaded._native)
        finally:
            loaded.close()
        candidate.close()
        if destination == "bytesio":
            assert not target.closed
            target.seek(0)
            assert target.read(2) == b"PK"
    finally:
        if isinstance(target, BytesIO):
            target.close()


@pytest.mark.parametrize(
    "formula",
    ["=A1+2", "=SUM(A1:A2)", '=IF(A1>0,"yes","no")', "=ROUND(1.25,1)", "=NOW()"],
)
def test_formulas_are_not_calculated_and_missing_cache_remains_none(candidate, formula):
    candidate.active["A1"] = 1
    candidate.active["A2"] = 2
    candidate.active["B1"] = formula
    payload = _save_bytes(candidate)
    for reader in (openpyxl.load_workbook, load_workbook):
        for data_only in (False, True):
            loaded = reader(BytesIO(payload), data_only=data_only)
            try:
                assert loaded.active["B1"].value == (None if data_only else formula)
            finally:
                loaded.close()
    loaded = load_workbook(BytesIO(payload), read_only=True, data_only=True)
    try:
        assert next(iter(loaded.active.values))[1] is None
    finally:
        loaded.close()


@pytest.mark.parametrize(
    "value",
    [
        "",
        "literal _x000D_ token",
        "_x0041_",
        "bad\x01text",
        2**53 + 1,
        -(2**53 + 1),
        2**60,
        2**63 - 1,
        -(2**63),
        float("nan"),
        float("inf"),
        -float("inf"),
    ],
)
def test_unsupported_values_fail_without_touching_destination(
    candidate, tmp_path, value
):
    candidate.active["A1"] = value
    _assert_unchanged(candidate, tmp_path, ValueError)


@pytest.mark.parametrize(
    "formula",
    [
        "=XLOOKUP(A1,A1:A2,B1:B2)",
        "=FILTER(A1:A2,A1:A2>0)",
        "=UNKNOWN(1)",
        "=SUM(Table1[Amount])",
        "=[other.xlsx]Sheet1!A1",
        "=A1#",
        "={1,2}",
        "=@SUM(A1:A2)",
        "=1+\n2",
    ],
)
def test_unsupported_formula_subset_fails_safely(candidate, tmp_path, formula):
    candidate.active["A1"] = formula
    _assert_unchanged(candidate, tmp_path, ValueError)


@pytest.mark.parametrize("component", ["font", "fill", "border"])
@pytest.mark.parametrize(
    "color",
    [
        Color(rgb="123456"),
        Color(rgb="00123456"),
        Color(rgb="80123456"),
        Color(theme=1),
        Color(rgb="FF123456", tint=0.25),
        Color(indexed=3),
        Color(auto=True),
    ],
)
def test_non_opaque_rgb_colors_are_rejected_by_native_validator(
    candidate, tmp_path, component, color
):
    if component == "font":
        value = Font(name="Arial", size=12, color=color)
    elif component == "fill":
        value = PatternFill("solid", fgColor=color)
    else:
        value = Border(left=Side("thin", color))
    setattr(candidate.active["A1"], component, value)
    _assert_unchanged(candidate, tmp_path, ValueError)


@pytest.mark.parametrize(
    "component,value",
    [
        ("font", Font(name=None)),
        ("font", Font(scheme="minor")),
        ("fill", PatternFill("darkGrid", fgColor="FF123456")),
        ("border", Border(diagonal=Side("thin", "FF123456"), diagonalUp=True)),
        ("alignment", Alignment(horizontal="center", indent=1)),
    ],
)
def test_unsupported_modeled_styles_fail_safely(candidate, tmp_path, component, value):
    setattr(candidate.active["A1"], component, value)
    _assert_unchanged(candidate, tmp_path, ValueError)


@pytest.mark.parametrize(
    "component", ["font", "fill", "border", "alignment", "protection"]
)
@pytest.mark.parametrize("kind", ["custom-object", "opaque-model", "unpooled-model"])
def test_non_frozen_style_components_are_not_silently_dropped(
    candidate, tmp_path, component, kind
):
    models = {
        "font": Font,
        "fill": PatternFill,
        "border": Border,
        "alignment": Alignment,
        "protection": Protection,
    }
    value = object() if kind == "custom-object" else models[component]()
    if kind == "opaque-model":
        value.unsupported_extension = "must not disappear"
    cell = candidate.active["A1"]
    if kind == "unpooled-model":
        # Direct sidecar mutation must not bypass the FrozenStyle requirement.
        getattr(cell, component)
        setattr(candidate.active._metadata.cells["A1"], component, value)
    else:
        setattr(cell, component, value)
    _assert_unchanged(candidate, tmp_path, NotImplementedError)


@pytest.mark.parametrize(
    "feature",
    [
        "properties",
        "merge",
        "comment",
        "hyperlink",
        "width",
        "height",
        "freeze",
        "filter",
        "validation",
        "unknown-book",
        "unknown-sheet",
        "unknown-cell",
    ],
)
def test_metadata_overlay_is_rejected_before_destination_changes(
    candidate, tmp_path, feature
):
    sheet = candidate.active
    sheet["A1"] = "kept"
    if feature == "properties":
        candidate.properties.title = "Unsupported metadata"
    elif feature == "merge":
        sheet.merge_cells("A1:B1")
    elif feature == "comment":
        sheet["A1"].comment = Comment("note", "author")
    elif feature == "hyperlink":
        sheet["A1"].hyperlink = "https://example.com"
    elif feature == "width":
        sheet.column_dimensions["A"].width = 22
    elif feature == "height":
        sheet.row_dimensions[1].height = 30
    elif feature == "freeze":
        sheet.freeze_panes = "B2"
    elif feature == "filter":
        sheet.auto_filter.ref = "A1:B2"
    elif feature == "validation":
        sheet._metadata.data_validations.append(object())
    else:
        sheet["A1"].font = Font(name="Arial", size=12)
        if feature == "unknown-book":
            candidate._book.unsupported_extension = "x"
        elif feature == "unknown-sheet":
            sheet._metadata.unsupported_extension = "x"
        else:
            sheet._metadata.cells["A1"].unsupported_extension = "x"
    _assert_unchanged(candidate, tmp_path, NotImplementedError)


def test_invalid_edit_after_success_does_not_replace_last_good_save(
    candidate, tmp_path
):
    candidate.active["A1"] = "valid"
    path = tmp_path / "last-good.xlsx"
    candidate.save(path)
    before = path.read_bytes()
    candidate.active["A1"].font = Font(color="123456")
    with pytest.raises(ValueError):
        candidate.save(path)
    assert path.read_bytes() == before
    candidate.active["A1"].font = Font(color="FF123456")
    candidate.save(path)
    loaded = openpyxl.load_workbook(path)
    try:
        assert loaded.active["A1"].value == "valid"
        assert loaded.active["A1"].font.color.rgb == "FF123456"
    finally:
        loaded.close()
