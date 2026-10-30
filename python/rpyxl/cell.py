"""Lightweight cell views. No Python cell grid is retained."""
from copy import deepcopy
from datetime import date, datetime, time

from ._compat import cell_metadata_property
from ._core import Cell as NativeCell
from ._models import CellMetadata, Hyperlink


class Cell:
    __slots__ = ('parent', '_native')

    def __init__(self, parent, native):
        self.parent = parent
        self._native = native

    def __getattr__(self, name):
        return getattr(self._native, name)

    @property
    def value(self):
        return self._native.value

    @value.setter
    def value(self, value):
        self.parent._check_writable()
        if self.parent._is_merged(self.column, self.row):
            raise AttributeError('MergedCell value is read-only')
        self._native.value = value

    @property
    def internal_value(self):
        return self.value

    @property
    def data_type(self):
        return self._native.data_type

    @data_type.setter
    def data_type(self, value):
        self.parent._check_writable()
        self._native.data_type = value

    @property
    def number_format(self):
        if not self.parent.parent.read_only:
            return self._native.number_format
        sheet = self.parent
        # Loading the shared metadata first gives the same warning and lifecycle
        # semantics as font/fill access. Only styled coordinates are retained.
        _ = sheet._metadata
        if sheet._number_formats is None:
            from ._compat import read_number_formats
            sheet._number_formats = read_number_formats(sheet.parent._source, sheet.title)
        return sheet._number_formats.get(self.coordinate, 'General')

    @number_format.setter
    def number_format(self, value):
        self.parent._check_writable()
        self._native.number_format = value

    @property
    def is_date(self):
        if hasattr(self._native, 'is_date'):
            return self._native.is_date
        return isinstance(self.value, (date, datetime, time))

    font = cell_metadata_property('font')
    fill = cell_metadata_property('fill')
    border = cell_metadata_property('border')
    alignment = cell_metadata_property('alignment')
    protection = cell_metadata_property('protection')

    @property
    def comment(self):
        return self.parent._metadata.cells.get(self.coordinate, CellMetadata()).comment

    @comment.setter
    def comment(self, value):
        self.parent._check_writable()
        metadata = self.parent._metadata.cells.setdefault(self.coordinate, CellMetadata())
        metadata.comment = deepcopy(value)
        # Clear old native annotations; the sidecar is authoritative (and can
        # represent authors and richer hyperlink fields that native cannot).
        self._native.comment = None

    @property
    def hyperlink(self):
        return self.parent._metadata.cells.get(self.coordinate, CellMetadata()).hyperlink

    @hyperlink.setter
    def hyperlink(self, value):
        self.parent._check_writable()
        if isinstance(value, str):
            value = Hyperlink(ref=self.coordinate, target=value)
        else:
            value = deepcopy(value)
            if value is not None:
                value.ref = self.coordinate
        self.parent._metadata.cells.setdefault(self.coordinate, CellMetadata()).hyperlink = value
        self._native.hyperlink = None
        if value is not None and self.value is None:
            self.value = value.target or value.location

    @property
    def has_style(self):
        metadata = self.parent._metadata.cells.get(self.coordinate)
        if metadata is not None:
            from ._compat import STYLE_DEFAULTS
            if any(getattr(metadata, name) is not None and getattr(metadata, name) != default()
                   for name, default in STYLE_DEFAULTS.items()):
                return True
        return self.number_format != 'General'

    def offset(self, row=0, column=0):
        return self.parent.cell(self.row + row, self.column + column)

    def __repr__(self):
        return '<Cell %r.%s>' % (self.parent.title, self.coordinate)


class MergedCell(Cell):
    @property
    def value(self):
        return None

    @value.setter
    def value(self, value):
        raise AttributeError('MergedCell value is read-only')


TYPE_STRING = 's'
TYPE_FORMULA = 'f'
TYPE_NUMERIC = 'n'
TYPE_BOOL = 'b'
TYPE_NULL = 'n'
TYPE_INLINE = 'inlineStr'
TYPE_ERROR = 'e'
TYPE_FORMULA_CACHE_STRING = 'str'
TYPE_DATE = 'd'

__all__ = ['Cell', 'MergedCell', 'NativeCell', 'TYPE_STRING', 'TYPE_FORMULA',
           'TYPE_NUMERIC', 'TYPE_BOOL', 'TYPE_NULL', 'TYPE_INLINE', 'TYPE_ERROR',
           'TYPE_FORMULA_CACHE_STRING', 'TYPE_DATE']
