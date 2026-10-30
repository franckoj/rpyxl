"""Worksheet facade: native values, sparse feature sidecars."""
from .._compat import (
    coordinate,
    metadata_property,
    move_metadata,
    structural_metadata,
)
from .._core import Worksheet as NativeWorksheet
from .._models import coordinate_parts, range_bounds
from ..cell import Cell, MergedCell


class Worksheet:
    SHEETSTATE_VISIBLE = 'visible'
    SHEETSTATE_HIDDEN = 'hidden'
    SHEETSTATE_VERYHIDDEN = 'veryHidden'
    ORIENTATION_PORTRAIT = 'portrait'
    ORIENTATION_LANDSCAPE = 'landscape'
    PAPERSIZE_A4 = 9
    PAPERSIZE_LETTER = 1

    def __init__(self, parent, native):
        self.parent = parent
        self._native = native
        self._number_formats = None

    def __getattr__(self, name):
        return getattr(self._native, name)

    def _check_writable(self):
        self.parent._check_writable()

    @property
    def _metadata(self):
        return self.parent._sheet_metadata(self.title)

    @property
    def title(self):
        return self._native.title

    @title.setter
    def title(self, value):
        self.parent._rename_sheet(self, value)

    merged_cells = metadata_property('merged_cells')
    column_dimensions = metadata_property('column_dimensions')
    row_dimensions = metadata_property('row_dimensions')
    auto_filter = metadata_property('auto_filter')
    sheet_properties = metadata_property('sheet_properties')
    protection = metadata_property('protection')
    page_setup = metadata_property('page_setup')
    data_validations = metadata_property('data_validations')
    conditional_formatting = metadata_property('conditional_formatting')
    tables = metadata_property('tables')
    _charts = metadata_property('charts')

    @property
    def freeze_panes(self):
        return self._metadata.freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value):
        self._check_writable()
        value = getattr(value, 'coordinate', value)
        if value is not None:
            col, row = coordinate_parts(value)
            value = coordinate(col, row)
        self._metadata.freeze_panes = None if value == 'A1' else value

    @property
    def sheet_state(self):
        return self._metadata.sheet_state

    @sheet_state.setter
    def sheet_state(self, value):
        self._check_writable()
        if value not in ('visible', 'hidden', 'veryHidden'):
            raise ValueError('Invalid sheet state')
        self._metadata.sheet_state = value

    def _is_merged(self, col, row):
        # Ordinary value access must not trigger lazy read-only metadata loading.
        book = self.parent._book
        metadata = book.sheets.get(self.title) if book is not None else None
        if metadata is not None:
            for merged in metadata.merged_cells:
                c1, r1, c2, r2 = range_bounds(str(merged))
                if c1 <= col <= c2 and r1 <= row <= r2:
                    return (col, row) != (c1, r1)
        return False

    def _wrap_cell(self, native):
        cls = MergedCell if self._is_merged(native.column, native.row) else Cell
        return cls(self, native)

    def _wrap_result(self, result):
        if isinstance(result, (tuple, list)):
            return tuple(self._wrap_result(item) for item in result)
        return self._wrap_cell(result)

    def cell(self, row, column, value=None):
        coordinate(column, row)
        if self.parent.write_only:
            raise AttributeError("'WriteOnlyWorksheet' object has no attribute 'cell'")
        if self.parent.read_only and value is not None:
            self._check_writable()
        native = self._native.cell(row=row, column=column)
        cell = self._wrap_cell(native)
        if value is not None:
            cell.value = value
        return cell

    def __getitem__(self, key):
        return self._wrap_result(self._native[key])

    def __setitem__(self, key, value):
        self._check_writable()
        col, row = coordinate_parts(key)
        self.cell(row=row, column=col).value = value

    def __iter__(self):
        return self.iter_rows()

    def append(self, iterable):
        self._check_writable()
        return self._native.append(iterable)

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        if self.parent.write_only:
            raise AttributeError("'WriteOnlyWorksheet' object has no attribute 'iter_rows'")
        args = dict(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)
        if values_only:
            return iter(self._native.iter_rows(values_only=True, **args))
        if self.parent.read_only:
            return (tuple(self._wrap_cell(cell) for cell in row)
                    for row in self._native.iter_rows(values_only=False, **args))
        if self.dimensions == 'A1' and all(value is None for value in args.values()):
            return (tuple(self._wrap_cell(cell) for cell in row)
                    for row in self._native.iter_rows(values_only=False))
        # The editable native iterator currently builds a list. Request one row
        # at a time, so wrapping cells never builds a full Python worksheet grid.
        def rows():
            start, end = min_row or self.min_row or 1, max_row or self.max_row or 1
            for row in range(start, end + 1):
                for cells in self._native.iter_rows(min_row=row, max_row=row,
                        min_col=min_col, max_col=max_col, values_only=False):
                    yield tuple(self._wrap_cell(cell) for cell in cells)
        return rows()

    def iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None, values_only=False):
        if self.parent.write_only:
            raise AttributeError("'WriteOnlyWorksheet' object has no attribute 'iter_cols'")
        if self.parent.read_only:
            raise NotImplementedError('Column iteration is unavailable in read-only mode')
        args = dict(min_col=min_col, max_col=max_col, min_row=min_row, max_row=max_row)
        if values_only:
            return iter(self._native.iter_cols(values_only=True, **args))
        if self.dimensions == 'A1' and all(value is None for value in args.values()):
            return (tuple(self._wrap_cell(cell) for cell in col)
                    for col in self._native.iter_cols(values_only=False))
        def columns():
            for col in range(min_col or self.min_column or 1, (max_col or self.max_column or 1) + 1):
                for cells in self._native.iter_cols(min_col=col, max_col=col,
                        min_row=min_row, max_row=max_row, values_only=False):
                    yield tuple(self._wrap_cell(cell) for cell in cells)
        return columns()

    @property
    def rows(self):
        return self.iter_rows()

    @property
    def columns(self):
        return self.iter_cols()

    @property
    def values(self):
        return self.iter_rows(values_only=True)

    def calculate_dimension(self, force=False):
        return self._native.dimensions

    def add_data_validation(self, validation):
        self._check_writable()
        self.data_validations.add(validation)

    def add_table(self, table):
        self._check_writable()
        if any(table.displayName in sheet.tables for sheet in self.parent.worksheets):
            raise ValueError('Table name already exists in this workbook')
        self.tables.add(table)

    def add_chart(self, chart, anchor=None):
        self._check_writable()
        self._metadata.add_chart(chart, anchor)

    @staticmethod
    def _range(range_string, start_row, start_column, end_row, end_column):
        if range_string is None:
            range_string = coordinate(start_column, start_row) + ':' + coordinate(end_column, end_row)
        bounds = range_bounds(range_string)
        return coordinate(bounds[0], bounds[1]) + ':' + coordinate(bounds[2], bounds[3]), bounds

    def _clear_loaded_merged_values(self):
        # Work row-by-row over the native used bounds, not a Python cell grid.
        # Valid merged files normally have nothing to clear; malformed producers
        # can leave hidden values which openpyxl discards on editable load.
        for merged in self._metadata.merged_cells:
            c1, r1, c2, r2 = range_bounds(str(merged))
            low_col, high_col = max(c1, self.min_column), min(c2, self.max_column)
            if low_col > high_col:
                continue
            for row in range(max(r1, self.min_row), min(r2, self.max_row) + 1):
                for values in self._native.iter_rows(min_row=row, max_row=row,
                        min_col=low_col, max_col=high_col, values_only=True):
                    for col, value in enumerate(values, low_col):
                        if value is not None and (col, row) != (c1, r1):
                            self._native.cell(row=row, column=col).value = None

    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        self._check_writable()
        ref, (c1, r1, c2, r2) = self._range(range_string, start_row, start_column, end_row, end_column)
        metadata = self._metadata
        if any(str(merged) == ref for merged in metadata.merged_cells):
            return
        for merged in metadata.merged_cells:
            a, b, c, d = range_bounds(str(merged))
            if max(a, c1) <= min(c, c2) and max(b, r1) <= min(d, r2):
                raise ValueError('Overlapping merged ranges are not supported')
        for row in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                if (col, row) == (c1, r1):
                    continue
                native = self._native.cell(row=row, column=col)
                native.value = None
                native.number_format = 'General'
                native.comment = None
                native.hyperlink = None
                metadata.cells.pop(coordinate(col, row), None)
        metadata.merged_cells.add(ref)

    def unmerge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        self._check_writable()
        ref, (c1, r1, c2, r2) = self._range(range_string, start_row, start_column, end_row, end_column)
        metadata = self._metadata
        metadata.merged_cells.remove(ref)
        for key in list(metadata.cells):
            col, row = coordinate_parts(key)
            if c1 <= col <= c2 and r1 <= row <= r2 and (col, row) != (c1, r1):
                del metadata.cells[key]

    def _check_structural_metadata(self):
        self._check_writable()
        if self.parent.write_only:
            raise TypeError('Structural edits are unavailable in write-only mode')
        book = self.parent._book
        if book is None:
            return None
        sheet = book.sheets.get(self.title)
        if book.defined_names or any(s.charts for s in book.sheets.values()):
            raise NotImplementedError('Structural edits with defined names or chart references are unsafe')
        if sheet is not None and (sheet.merged_cells.ranges or sheet.freeze_panes or
                sheet.auto_filter.ref or sheet.data_validations or sheet.conditional_formatting or sheet.tables):
            raise NotImplementedError(
                'Structural edits with merges, panes, filters, validations, conditional formatting '
                'or tables are not supported; remove those features first')
        return sheet

    def _structural(self, method, axis, idx, amount):
        if not isinstance(idx, int) or idx < 1 or not isinstance(amount, int) or amount < 0:
            raise ValueError('Index must be positive and amount must be nonnegative')
        sheet = self._check_structural_metadata()
        shifted = structural_metadata(sheet, axis, idx, amount, method.startswith('delete')) if sheet else None
        getattr(self._native, method)(idx, amount)
        if shifted is not None:
            self.parent._book.sheets[self.title] = shifted

    def insert_rows(self, idx, amount=1):
        self._structural('insert_rows', 'row', idx, amount)

    def delete_rows(self, idx, amount=1):
        self._structural('delete_rows', 'row', idx, amount)

    def insert_cols(self, idx, amount=1):
        self._structural('insert_cols', 'col', idx, amount)

    def delete_cols(self, idx, amount=1):
        self._structural('delete_cols', 'col', idx, amount)

    def move_range(self, cell_range, rows=0, cols=0, translate=False):
        if translate:
            raise NotImplementedError('Formula translation is not supported')
        sheet = self._check_structural_metadata()
        shifted = move_metadata(sheet, cell_range, rows, cols) if sheet else None
        self._native.move_range(cell_range, rows=rows, cols=cols, translate=False)
        if shifted is not None:
            self.parent._book.sheets[self.title] = shifted

    def __repr__(self):
        return '<Worksheet %r>' % self.title


__all__ = ['Worksheet', 'NativeWorksheet']
