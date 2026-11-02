"""Save-local eligibility reuse without hiding mutable or lazy metadata."""

from collections import Counter
from dataclasses import dataclass
from io import BytesIO
from unittest.mock import Mock
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import openpyxl
import pytest
from rpyxl import Workbook, _metadata, _native_styles, load_workbook
from rpyxl._models import CellMetadata, Comment, MetadataBook, MetadataSheet
from rpyxl.styles import Font


@pytest.fixture
def metadata_book():
    return MetadataBook(
        sheets={
            "Sheet": MetadataSheet(cells={"A1": CellMetadata(font=Font(bold=True))})
        }
    )


def _save(book):
    output = BytesIO()
    book.save(output)
    output.seek(0)
    return output


def _source(*, unsupported=False):
    source = openpyxl.Workbook()
    source.active["A1"] = "source"
    source.active["A1"].font = openpyxl.styles.Font(bold=True)
    source.active["B1"].comment = openpyxl.comments.Comment("Original", "Author")
    output = BytesIO()
    source.save(output)
    source.close()
    if not unsupported:
        return output
    # Valid OOXML, but numeric escapes disqualify native package preservation.
    rewritten = BytesIO()
    with ZipFile(output) as original, ZipFile(rewritten, "w") as target:
        for name in original.namelist():
            data = original.read(name)
            if name == "[Content_Types].xml":
                data = data.replace(b'Extension="xml"', b'Extension="&#120;ml"')
            target.writestr(name, data)
    return rewritten


def test_style_only_check_does_not_reflect_or_snapshot_each_cell(
    metadata_book, monkeypatch
):
    _native_styles._field_layout.cache_clear()
    reflect = Mock(wraps=_native_styles.fields)
    snapshot = Mock(wraps=_native_styles._snapshot)
    monkeypatch.setattr(_native_styles, "fields", reflect)
    monkeypatch.setattr(_native_styles, "_snapshot", snapshot)

    assert not _native_styles.requires_metadata_overlay(metadata_book)
    reflected_types = Counter(call.args[0] for call in reflect.call_args_list)
    assert CellMetadata not in reflected_types
    assert reflected_types and set(reflected_types.values()) == {1}
    snapshot_count = snapshot.call_count
    reflection_count = reflect.call_count
    snapshot.reset_mock()

    metadata_book.sheets["Sheet"].cells.update(
        {f"A{row}": CellMetadata(font=Font(bold=True)) for row in range(2, 1002)}
    )
    assert not _native_styles.requires_metadata_overlay(metadata_book)
    assert snapshot.call_count == snapshot_count
    assert reflect.call_count == reflection_count


@pytest.mark.parametrize("name", ["hyperlink", "comment"])
@pytest.mark.parametrize("value", ["", False, 0, {}, [], "annotation"])
def test_even_falsey_nonstyle_annotations_require_overlay(metadata_book, name, value):
    cell = metadata_book.sheets["Sheet"].cells["A1"]
    assert not _native_styles.requires_metadata_overlay(metadata_book)
    setattr(cell, name, value)
    assert _native_styles.requires_metadata_overlay(metadata_book)
    setattr(cell, name, None)
    assert not _native_styles.requires_metadata_overlay(metadata_book)


@pytest.mark.parametrize("value", [None, {}, ET.Element("cell")])
def test_unknown_cell_attributes_are_checked_after_direct_mutation(
    metadata_book, value
):
    cell = metadata_book.sheets["Sheet"].cells["A1"]
    assert not _native_styles.requires_metadata_overlay(metadata_book)
    cell.future_feature = value
    assert _native_styles.requires_metadata_overlay(metadata_book)
    del cell.future_feature
    assert not _native_styles.requires_metadata_overlay(metadata_book)

    # An unknown key must not be hidden by an unchanged dictionary length.
    del cell.font
    cell.future_feature = value
    assert _native_styles.requires_metadata_overlay(metadata_book)


class _CellSubclass(CellMetadata):
    pass


@dataclass
class _ExtendedCell(CellMetadata):
    annotation: object = None


@pytest.mark.parametrize("model_type", [_CellSubclass, _ExtendedCell])
def test_cell_subclasses_conservatively_require_overlay(metadata_book, model_type):
    metadata_book.sheets["Sheet"].cells["A1"] = model_type(font=Font(bold=True))
    assert _native_styles.requires_metadata_overlay(metadata_book)


@pytest.mark.parametrize("package_api", [True, False])
@pytest.mark.parametrize("rich", [True, False])
def test_new_book_checks_eligibility_once_per_save(monkeypatch, package_api, rich):
    check = Mock(wraps=_native_styles.requires_metadata_overlay)
    apply = Mock(wraps=_metadata.apply_metadata)
    monkeypatch.setattr(_native_styles, "requires_metadata_overlay", check)
    monkeypatch.setattr(_metadata, "apply_metadata", apply)
    with Workbook() as book:
        assert book._native_package_api
        book._native_package_api = package_api
        book.active["A1"].font = Font(bold=True)
        if rich:
            book.properties.title = "Rich metadata"
        for save_number in range(1, 3):
            saved = openpyxl.load_workbook(_save(book))
            assert check.call_count == save_number
            assert apply.call_count == (save_number if rich else 0)
            assert saved.active["A1"].font.bold
            if rich:
                assert saved.properties.title == "Rich metadata"
            saved.close()


@pytest.mark.parametrize("target", ["book", "sheet", "cell"])
def test_direct_unknown_attribute_mutation_changes_next_save_eligibility(
    monkeypatch, target
):
    check = Mock(wraps=_native_styles.requires_metadata_overlay)
    apply = Mock(wraps=_metadata.apply_metadata)
    monkeypatch.setattr(_native_styles, "requires_metadata_overlay", check)
    monkeypatch.setattr(_metadata, "apply_metadata", apply)
    with Workbook() as book:
        book.active["A1"].font = Font(bold=True)
        sheet = book._book.sheets[book.active.title]
        model = {"book": book._book, "sheet": sheet, "cell": sheet.cells["A1"]}[target]
        _save(book)
        assert check.call_count == 1
        assert apply.call_count == 0

        model.future_feature = {}
        _save(book)
        assert check.call_count == 2
        assert apply.call_count == 1

        del model.future_feature
        _save(book)
        assert check.call_count == 3
        assert apply.call_count == 1


def test_annotation_and_nested_style_mutation_between_saves(monkeypatch):
    check = Mock(wraps=_native_styles.requires_metadata_overlay)
    apply = Mock(wraps=_metadata.apply_metadata)
    monkeypatch.setattr(_native_styles, "requires_metadata_overlay", check)
    monkeypatch.setattr(_metadata, "apply_metadata", apply)
    with Workbook() as book:
        book.active["A1"].font = Font(bold=True, color="FF123456")
        _save(book)
        cell = book._book.sheets[book.active.title].cells["A1"]
        cell.comment = Comment("Direct annotation", "Author")
        book.active["A1"].font.color.rgb = "FFABCDEF"
        saved = openpyxl.load_workbook(_save(book))
        assert saved.active["A1"].comment.text == "Direct annotation"
        assert saved.active["A1"].font.color.rgb == "FFABCDEF"
        assert check.call_count == 2
        assert apply.call_count == 1
        saved.close()

        cell.comment = None
        book.active["A1"].font.italic = True
        saved = openpyxl.load_workbook(_save(book))
        assert saved.active["A1"].comment is None
        assert saved.active["A1"].font.italic
        assert check.call_count == 3
        assert apply.call_count == 1
        saved.close()


def test_loaded_style_only_save_stays_lazy_with_one_check(monkeypatch):
    with load_workbook(_source()) as book:
        assert book._native_package_api
        assert book._native._can_preserve_package()
        book.active["A1"].font.italic = True
        check = Mock(wraps=_native_styles.requires_metadata_overlay)
        monkeypatch.setattr(_native_styles, "requires_metadata_overlay", check)

        def unexpected(*args, **kwargs):
            raise AssertionError(
                "Style-only saves must not promote or overlay metadata"
            )

        monkeypatch.setattr(_metadata, "load_metadata", unexpected)
        monkeypatch.setattr(_metadata, "apply_metadata", unexpected)
        for save_number in range(1, 3):
            saved = openpyxl.load_workbook(_save(book))
            assert check.call_count == save_number
            assert not book._full_metadata_loaded
            assert saved.active["A1"].font.bold
            assert saved.active["A1"].font.italic
            assert saved.active["B1"].comment.text == "Original"
            saved.close()


def test_full_metadata_promotion_rechecks_eligibility(monkeypatch):
    with load_workbook(_source(unsupported=True)) as book:
        assert book._native_package_api
        assert not book._native._can_preserve_package()
        book.active["A1"].font.italic = True
        results = []
        original = _native_styles.requires_metadata_overlay

        def check(metadata):
            result = original(metadata)
            results.append(result)
            return result

        apply = Mock(wraps=_metadata.apply_metadata)
        monkeypatch.setattr(_native_styles, "requires_metadata_overlay", check)
        monkeypatch.setattr(_metadata, "apply_metadata", apply)
        saved = openpyxl.load_workbook(_save(book))
        assert results == [False, True]
        assert book._full_metadata_loaded
        assert apply.call_count == 1
        assert saved.active["A1"].font.italic
        assert saved.active["B1"].comment.text == "Original"
        saved.close()

        results.clear()
        _save(book)
        assert results == [True]  # No second check once already fully promoted.
        assert apply.call_count == 2


def test_native_structure_optin_requires_an_original_package():
    with Workbook() as book:
        assert not book._native_structure_only
        assert not book._try_native_structure()
        assert not book._native_structure_only
        assert not book._structure_changed
        assert book._book is None


def test_old_native_without_structure_api_declines_without_loading_metadata():
    with load_workbook(_source()) as book:
        if hasattr(book._native, "_enable_plain_structure"):
            pytest.skip("This compatibility guard targets older native wheels")
        assert book._native_package_api
        assert book._has_original
        assert not book._try_native_structure()
        assert not book._native_structure_only
        assert not book._structure_changed
        assert book._book is None
        assert not book._full_metadata_loaded


def test_prepare_structure_clears_native_structure_mode_and_loads_metadata():
    with load_workbook(_source()) as book:
        book._native_structure_only = True
        book._prepare_structure()
        assert not book._native_structure_only
        assert not book._structure_changed
        assert book._full_metadata_loaded
        assert (
            book._book.sheets[book.active.title].cells["B1"].comment.text == "Original"
        )
        assert not book._native_style_reader.opaque_cells


@pytest.mark.parametrize("structure_changed", [False, True])
def test_invalid_native_structure_proof_forces_metadata_fallback(
    monkeypatch, structure_changed
):
    with load_workbook(_source(unsupported=True)) as book:
        assert not book._native._can_preserve_package()
        # Exercise save's invalidated-proof state even with an older wheel.
        book._native_structure_only = True
        book._structure_changed = structure_changed
        apply = Mock(wraps=_metadata.apply_metadata)
        monkeypatch.setattr(_metadata, "apply_metadata", apply)
        for save_number in range(1, 3):
            saved = openpyxl.load_workbook(_save(book))
            assert book._full_metadata_loaded
            assert apply.call_count == save_number
            assert saved.active["A1"].value == "source"
            assert saved.active["B1"].comment.text == "Original"
            saved.close()
