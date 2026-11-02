"""Wrapper regressions; deliberately kept independent of openpyxl."""
from copy import copy
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from rpyxl.chart import BarChart, Reference
from rpyxl.comments import Comment
from rpyxl.formatting.rule import CellIsRule
from rpyxl.styles import Font
from rpyxl.workbook.defined_name import DefinedName
from rpyxl.worksheet.datavalidation import DataValidation
from rpyxl.worksheet.table import Table, TableStyleInfo

from rpyxl import Cell, Workbook, Worksheet, load_workbook

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def roundtrip(book):
    output = BytesIO()
    book.save(output)
    output.seek(0)
    return load_workbook(output)


def test_wrapper_identity_and_native_value_path():
    book = Workbook()
    sheet = book.active
    sheet.append([1, '=A1+1'])
    assert isinstance(book, Workbook)
    assert isinstance(sheet, Worksheet)
    assert isinstance(sheet['A1'], Cell)
    assert sheet.parent is book
    assert sheet['A1'].parent is sheet
    assert book[sheet.title] is sheet
    assert list(sheet.values) == [(1, '=A1+1')]
    assert book._book is None
    assert sheet['A1'].value == 1
    assert book._book is None
    assert next(sheet.rows)[0].parent is sheet
    book.active = sheet
    assert book.active is sheet


def test_simple_save_bypasses_overlay(monkeypatch):
    import rpyxl._metadata as metadata

    def unexpected(*args):
        raise AssertionError('Value-only new books must not overlay metadata')

    monkeypatch.setattr(metadata, 'apply_metadata', unexpected)
    book = Workbook()
    book.active['A1'] = 'plain'
    book.save(BytesIO())
    assert book._book is None


def test_native_validation_precedes_metadata_reading(monkeypatch):
    import rpyxl._metadata as metadata
    import rpyxl.reader.excel as reader

    def reject(*args, **kwargs):
        raise ValueError('native rejected input')

    def unexpected(*args, **kwargs):
        raise AssertionError('Metadata parser must not run before native validation')

    monkeypatch.setattr(reader, '_load_workbook', reject)
    monkeypatch.setattr(metadata, 'load_metadata', unexpected)
    with pytest.raises(ValueError, match='native rejected'):
        reader.load_workbook(BytesIO(b'invalid'))


def test_copyable_defaults_and_nested_mutation_persist():
    book = Workbook()
    sheet = book.active
    font = copy(sheet['A1'].font)
    font.bold = True
    sheet['A1'].font = font
    sheet['A1'].alignment.wrap_text = True
    sheet['B1'].font.italic = True
    assert not sheet['A1'].font.italic
    loaded = roundtrip(book).active
    assert loaded['A1'].font.bold
    assert loaded['A1'].alignment.wrap_text
    assert loaded['B1'].font.italic


def test_feature_metadata_roundtrip():
    book = Workbook()
    sheet = book.active
    sheet.append(['Name', 'Amount'])
    sheet.append(['Example', 12.5])
    sheet['B2'].number_format = '0.0000'
    sheet['A2'].comment = Comment('A note', 'Reviewer')
    sheet['A2'].hyperlink = 'https://example.com'
    sheet.column_dimensions['B'].width = 22
    sheet.row_dimensions[2].height = 30
    sheet.freeze_panes = sheet['B2']
    sheet.auto_filter.ref = 'A1:B2'
    sheet.sheet_properties.tabColor = 'FF0000'
    sheet.protection.set_password('secret')
    sheet.page_setup.orientation = 'landscape'
    book.properties.title = 'Metadata test'
    book.defined_names.add(DefinedName('Amounts', attr_text="'Sheet'!$B$2"))
    validation = DataValidation(type='whole', formula1=0, formula2=100, operator='between')
    validation.add('B2:B5')
    sheet.add_data_validation(validation)
    sheet.conditional_formatting.add('B2:B5', CellIsRule(operator='greaterThan', formula=['10'], font=Font(b=True)))
    table = Table(displayName='Entries', ref='A1:B2')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
    sheet.add_table(table)
    chart = BarChart()
    chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=2), titles_from_data=True)
    sheet.add_chart(chart, 'D1')

    loaded = roundtrip(book)
    target = loaded.active
    assert loaded.properties.title == 'Metadata test'
    assert loaded.defined_names['Amounts'].attr_text == "'Sheet'!$B$2"
    assert target['B2'].number_format == '0.0000'
    assert target['A2'].comment.author == 'Reviewer'
    assert target['A2'].hyperlink.target == 'https://example.com'
    assert target.column_dimensions['B'].width == 22
    assert target.row_dimensions[2].height == 30
    assert target.freeze_panes == 'B2'
    assert target.auto_filter.ref == 'A1:B2'
    assert target.sheet_properties.tabColor.rgb == '00FF0000'
    assert target.protection.password
    assert target.page_setup.orientation == 'landscape'
    assert target.data_validations[0].sqref == 'B2:B5'
    assert len(target.conditional_formatting) == 1
    assert target.tables['Entries'].tableStyleInfo.showRowStripes
    assert len(target._charts) == 1


def test_merges_clear_non_anchor_values_and_unmerge_does_not_resurrect():
    book = Workbook()
    sheet = book.active
    sheet.append(['anchor', 'discard'])
    old_cell = sheet['B1']
    sheet.merge_cells('A1:B1')
    assert sheet['A1'].value == 'anchor'
    assert sheet['B1'].value is None
    with pytest.raises(AttributeError):
        sheet['B1'] = 'blocked'
    with pytest.raises(AttributeError):
        old_cell.value = 'also blocked'
    assert list(sheet.values) == [('anchor', None)]
    sheet.unmerge_cells('A1:B1')
    assert sheet['B1'].value is None
    sheet['B1'] = 'new'
    assert roundtrip(book).active['B1'].value == 'new'


def test_editable_load_clears_non_anchor_values_in_merged_source():
    book = Workbook()
    sheet = book.active
    sheet.append(['anchor', 'orphan'])
    # Direct range registration simulates a producer that does not clean values.
    sheet.merged_cells.add('A1:B1')
    loaded = roundtrip(book)
    assert loaded.active['B1'].value is None
    assert list(loaded.active.values) == [('anchor', None)]


def test_structural_edits_move_sparse_styles_comments_and_dimensions():
    book = Workbook()
    sheet = book.active
    sheet['B2'] = 'move'
    sheet['B2'].font = Font(bold=True)
    sheet['B2'].comment = Comment('move too', 'Author')
    sheet['B2'].hyperlink = 'https://example.com'
    sheet.row_dimensions[2].height = 25
    sheet.column_dimensions['B'].width = 19
    sheet.insert_rows(2)
    sheet.insert_cols(2)
    assert sheet['C3'].value == 'move'
    assert sheet['C3'].font.bold
    assert sheet['C3'].comment.text == 'move too'
    assert sheet['C3'].hyperlink.ref == 'C3'
    assert sheet.row_dimensions[3].height == 25
    assert sheet.column_dimensions['C'].width == 19
    sheet.move_range('C3', rows=2, cols=1)
    assert sheet['D5'].font.bold
    assert sheet['D5'].hyperlink.ref == 'D5'
    assert sheet['C3'].value is None
    sheet.delete_rows(1)
    sheet.delete_cols(1)
    loaded = roundtrip(book).active
    assert loaded['C4'].value == 'move'
    assert loaded['C4'].font.bold
    assert loaded['C4'].comment.text == 'move too'


def test_copy_deepcopies_validations_and_styles_and_rename_rekeys_metadata():
    book = Workbook()
    sheet = book.active
    sheet['A1'] = 1
    sheet['A1'].font = Font(bold=True)
    validation = DataValidation(type='whole', formula1=0, sqref='A1:A3')
    sheet.add_data_validation(validation)
    copied = book.copy_worksheet(sheet)
    copied.data_validations[0].formula1 = 5
    copied['A1'].font.italic = True
    assert sheet.data_validations[0].formula1 == 0
    assert not sheet['A1'].font.italic
    copied.title = 'Renamed'
    loaded = roundtrip(book)
    assert loaded['Renamed']['A1'].font.italic
    assert loaded['Renamed'].data_validations[0].formula1 == '5'
    assert copied.parent is book


def test_unsafe_reference_edits_fail_before_native_mutation():
    book = Workbook()
    sheet = book.active
    sheet['A1'] = 10
    book.defined_names.add(DefinedName('Input', attr_text="'Sheet'!$A$1"))
    with pytest.raises(NotImplementedError):
        sheet.title = 'Renamed'
    with pytest.raises(NotImplementedError):
        sheet.insert_rows(1)
    assert sheet.title == 'Sheet'
    assert sheet['A1'].value == 10


def test_read_only_value_iteration_stays_lazy_and_stream_can_close():
    book = Workbook()
    book.active['A1'] = 12
    book.active['A1'].font = Font(bold=True)
    book.active['A1'].number_format = '0.00000'
    source = BytesIO()
    book.save(source)
    source.seek(7)
    loaded = load_workbook(source, read_only=True)
    assert source.tell() == 7
    source.close()
    assert loaded._book is None
    sheet = loaded.active
    assert list(sheet.values) == [(12,)]
    assert loaded._book is None
    cell = sheet['A1']
    assert cell.parent is sheet
    with pytest.warns(RuntimeWarning, match='metadata for all sheets'):
        assert cell.font.bold
    assert cell.number_format == '0.00000'
    with pytest.raises(TypeError):
        cell.value = 1
    loaded.close()
    assert loaded._source.closed


def test_overlay_uses_renamed_title_and_removes_cleared_features():
    book = Workbook()
    sheet = book.active
    sheet['A1'] = 'text'
    sheet['A1'].comment = Comment('remove', 'Author')
    sheet['A1'].hyperlink = 'https://example.com'
    loaded = roundtrip(book)
    target = loaded.active
    target['A1'].comment = None
    target['A1'].hyperlink = None
    target.title = 'After'
    output = BytesIO()
    loaded.save(output)
    with ZipFile(output) as archive:
        root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        assert root.find(S + 'hyperlinks') is None
    reloaded = load_workbook(output)
    assert reloaded.sheetnames == ['After']
    assert reloaded.active['A1'].comment is None
    assert reloaded.active['A1'].hyperlink is None
