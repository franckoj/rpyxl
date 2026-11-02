"""Small writer/reader matrix, not a duplicate of the wrapper metadata suite."""
import warnings
from contextlib import ExitStack
from datetime import date, datetime, time, timedelta
from importlib import import_module
from io import BytesIO

import openpyxl
import pytest

import rpyxl

LIBRARIES = {'openpyxl': openpyxl, 'rpyxl': rpyxl}
# Numeric Excel serials are rounded to milliseconds by openpyxl. Allow one
# millisecond, not an entire second, for readers with different rounding paths.
TEMPORAL_TOLERANCE = timedelta(milliseconds=1)


def _populate(book, library):
    styles = import_module(library + '.styles')
    comments = import_module(library + '.comments')
    tables = import_module(library + '.worksheet.table')
    validations = import_module(library + '.worksheet.datavalidation')
    rules = import_module(library + '.formatting.rule')
    charts = import_module(library + '.chart')
    names = import_module(library + '.workbook.defined_name')
    sheet = book.active
    sheet.title = 'Interop'
    sheet.append(['Name', 'Amount', 'Date', 'Datetime', 'Time'])
    sheet.append(['Example', 12.5, date(2024, 2, 29),
                  datetime(2024, 2, 29, 12, 34, 56, 123456),
                  time(12, 34, 56, 789123)])
    sheet['B2'].number_format = '0.0000'
    sheet['C2'].number_format = 'yyyy-mm-dd'
    sheet['D2'].number_format = 'yyyy-mm-dd hh:mm:ss.000'
    sheet['E2'].number_format = 'hh:mm:ss.000'
    sheet['B2'].font = styles.Font(name='Arial', bold=True, color='FF123456')
    sheet['B2'].fill = styles.PatternFill(fill_type='solid', fgColor='FFABCDEF')
    sheet['B2'].alignment = styles.Alignment(horizontal='right', wrap_text=True)
    sheet['B2'].border = styles.Border(left=styles.Side(style='thin', color='FF112233'))
    sheet['A2'].hyperlink = 'https://example.com/report'
    sheet['A2'].comment = comments.Comment('Reviewed amount', 'Reviewer')
    sheet['A4'] = 'Merged label'
    sheet.merge_cells('A4:B4')
    sheet.column_dimensions['B'].width = 23
    sheet.column_dimensions['C'].hidden = True
    sheet.row_dimensions[2].height = 28
    sheet.row_dimensions[3].hidden = True
    sheet.freeze_panes = 'C2'
    sheet.auto_filter.ref = 'A1:E2'
    book.properties.title = 'Interoperability control'
    book.properties.creator = 'Test author'
    book.properties.description = 'Small cross-library fixture'
    table = tables.Table(displayName='Entries', ref='A1:B2')
    table.tableStyleInfo = tables.TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
    sheet.add_table(table)
    validation = validations.DataValidation(type='whole', operator='between', formula1=0, formula2=100)
    validation.add('B2:B3')
    sheet.add_data_validation(validation)
    sheet.conditional_formatting.add(
        'B2:B3', rules.CellIsRule(operator='greaterThan', formula=['10'], font=styles.Font(bold=True)))
    chart = charts.BarChart()
    chart.add_data(charts.Reference(sheet, min_col=2, min_row=1, max_row=2), titles_from_data=True)
    sheet.add_chart(chart, 'G2')
    book.defined_names.add(names.DefinedName('Amounts', attr_text="'Interop'!$B$2"))


def _save(library, target):
    book = LIBRARIES[library].Workbook()
    try:
        _populate(book, library)
        book.save(target)
    finally:
        book.close()


@pytest.fixture
def control_bytes():
    output = BytesIO()
    _save('openpyxl', output)
    return output.getvalue()


@pytest.fixture(params=['openpyxl', 'rpyxl'], ids=lambda name: 'writer-' + name)
def written_source(request, tmp_path, transport):
    if transport == 'path':
        source = tmp_path / (request.param + '.xlsx')
        _save(request.param, source)
    else:
        source = BytesIO()
        _save(request.param, source)
        source.seek(0)
    yield source
    if isinstance(source, BytesIO):
        source.close()


@pytest.fixture(params=['path', 'bytesio'])
def transport(request):
    return request.param


def _assert_value(actual, expected):
    # The control decides date vs datetime: a date-formatted Excel serial may
    # legitimately reload as datetime rather than the originally assigned date.
    if isinstance(expected, (date, time)):
        assert type(actual) is type(expected), (actual, expected)
        if isinstance(expected, datetime):
            assert abs(actual - expected) <= TEMPORAL_TOLERANCE
        elif isinstance(expected, time):
            day = date(2024, 1, 1)
            assert abs(datetime.combine(day, actual) - datetime.combine(day, expected)) <= TEMPORAL_TOLERANCE
        else:
            assert actual == expected
    else:
        assert actual == expected


def _style(cell):
    return (cell.number_format, cell.font.name, cell.font.bold, cell.font.color.rgb,
            cell.fill.fill_type, cell.fill.fgColor.rgb, cell.alignment.horizontal,
            cell.alignment.wrap_text, cell.border.left.style, cell.border.left.color.rgb)


@pytest.mark.parametrize('reader', ['openpyxl', 'rpyxl'], ids=lambda name: 'reader-' + name)
@pytest.mark.parametrize('read_only', [False, True], ids=['editable', 'read-only'])
def test_values_styles_and_temporals_match_openpyxl_control(
        written_source, control_bytes, reader, read_only):
    with ExitStack() as stack:
        control_source = stack.enter_context(BytesIO(control_bytes))
        control = openpyxl.load_workbook(control_source, read_only=read_only)
        stack.callback(control.close)
        actual = LIBRARIES[reader].load_workbook(written_source, read_only=read_only)
        stack.callback(actual.close)
        assert actual.sheetnames == control.sheetnames
        expected_rows = list(control.active.iter_rows(min_row=1, max_row=4, max_col=5, values_only=True))
        actual_rows = list(actual.active.iter_rows(min_row=1, max_row=4, max_col=5, values_only=True))
        assert len(actual_rows) == len(expected_rows)
        for actual_row, expected_row in zip(actual_rows, expected_rows):
            assert len(actual_row) == len(expected_row)
            for value, expected in zip(actual_row, expected_row):
                _assert_value(value, expected)
        # Only rpyxl defers metadata with a warning; openpyxl read-only cells
        # expose styles but not worksheet features such as comments or tables.
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', message='Accessing read-only metadata', category=RuntimeWarning)
            assert _style(actual.active['B2']) == _style(control.active['B2'])
            for coordinate in ('C2', 'D2', 'E2'):
                assert actual.active[coordinate].number_format == control.active[coordinate].number_format


def _features(book):
    sheet = book.active
    validation_items = sheet.data_validations.dataValidation
    cf = []
    for key in sheet.conditional_formatting:
        cell_range = str(getattr(key, 'sqref', key))
        cf.append((cell_range, [(rule.type, rule.operator, tuple(rule.formula))
                               for rule in sheet.conditional_formatting[key]]))
    chart = sheet._charts[0]
    series = chart.series[0]
    # rpyxl models the reference directly; openpyxl wraps it in NumDataSource.
    chart_values = series.values if hasattr(series, 'values') else series.val.numRef.f
    return {
        'merge': sorted(str(value) for value in sheet.merged_cells.ranges),
        'merge_values': (sheet['A4'].value, sheet['B4'].value),
        'dimensions': (sheet.max_row, sheet.max_column, sheet.column_dimensions['B'].width,
                       sheet.column_dimensions['C'].hidden, sheet.row_dimensions[2].height,
                       sheet.row_dimensions[3].hidden),
        'freeze': sheet.freeze_panes,
        'filter': sheet.auto_filter.ref,
        'link': sheet['A2'].hyperlink.target,
        'comment': (sheet['A2'].comment.text, sheet['A2'].comment.author),
        'properties': (book.properties.title, book.properties.creator, book.properties.description),
        'table_names': sorted(sheet.tables),
        'table': (sheet.tables['Entries'].ref, sheet.tables['Entries'].tableStyleInfo.name,
                  sheet.tables['Entries'].tableStyleInfo.showRowStripes),
        'validations': [(item.type, item.operator, str(item.sqref), str(item.formula1), str(item.formula2))
                        for item in validation_items],
        'conditional_formatting': cf,
        'chart': (len(sheet._charts), chart.type, len(chart.series), chart_values),
        'name': book.defined_names['Amounts'].attr_text,
    }


@pytest.mark.parametrize('reader', ['openpyxl', 'rpyxl'], ids=lambda name: 'reader-' + name)
def test_editable_features_match_openpyxl_control(written_source, control_bytes, reader):
    with ExitStack() as stack:
        control_source = stack.enter_context(BytesIO(control_bytes))
        control = openpyxl.load_workbook(control_source)
        stack.callback(control.close)
        actual = LIBRARIES[reader].load_workbook(written_source)
        stack.callback(actual.close)
        assert _features(actual) == _features(control)
