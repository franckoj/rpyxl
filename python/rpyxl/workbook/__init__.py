"""Native workbooks with sparse, persistent compatibility metadata."""
import warnings
from copy import deepcopy
from tempfile import TemporaryFile

from .._core import ReadOnlyWorkbook as NativeReadOnlyWorkbook
from .._core import Workbook as NativeWorkbook
from .._core import WriteOnlyWorkbook as NativeWriteOnlyWorkbook
from .._models import MetadataBook, MetadataSheet


class _WorkbookMeta(type):
    def __instancecheck__(cls, instance):
        return type.__instancecheck__(cls, instance) or isinstance(
            instance, (NativeWorkbook, NativeWriteOnlyWorkbook, NativeReadOnlyWorkbook))


class Workbook(metaclass=_WorkbookMeta):
    """Keep values in Rust and materialize feature models only when needed.

    New value-only books save directly through Rust. Once mutable metadata is
    exposed, saves conservatively overlay it, including nested in-place edits.
    Loaded editable books always overlay their original metadata. This costs an
    additional ZIP/XML pass but avoids an O(cells) Python dirty snapshot.

    Read-only metadata is deferred, not sheet-selective: first access warns and
    materializes all metadata. Value iteration stays native. Structural edits
    move sparse cell styles/annotations and dimensions, but reject reference-
    bearing features rather than guessing how to translate them. Renames reject
    books with names/charts; sheet copies reject tables/charts and do not copy
    sheet-local names. Formula translation and full merged-border propagation
    are not implemented. Native* exports provide explicit unwrapped access.
    """

    def __init__(self, write_only=False, *, _native=None, _source=None, _owns_source=False):
        self._native = _native if _native is not None else (
            NativeWriteOnlyWorkbook() if write_only else NativeWorkbook())
        self._source = _source
        self._owns_source = _owns_source
        self._book = None
        self._sheets = {}
        self._overlay = False
        self._closed = False
        if _source is not None and not self.read_only:
            book = self._ensure_metadata()
            for title, metadata in book.sheets.items():
                if metadata.merged_cells.ranges:
                    self[title]._clear_loaded_merged_values()

    def __getattr__(self, name):
        return getattr(self._native, name)

    @property
    def read_only(self):
        return self._native.read_only

    @property
    def write_only(self):
        return getattr(self._native, 'write_only', False)

    def _check_writable(self):
        if self._closed:
            raise ValueError('Workbook is closed')
        if self.read_only:
            raise TypeError('Cannot modify a read-only workbook')

    def _ensure_metadata(self):
        if self._closed:
            raise ValueError('Workbook is closed')
        if self._book is None:
            if self._source is None:
                self._book = MetadataBook()
            else:
                if self.read_only:
                    warnings.warn(
                        'Accessing read-only metadata loads metadata for all sheets; '
                        'values_only iteration does not incur this cost.',
                        RuntimeWarning, stacklevel=3)
                from .._metadata import load_metadata
                self._book = load_metadata(self._source)
        self._overlay = True
        return self._book

    def _sheet_metadata(self, title):
        book = self._ensure_metadata()
        return book.sheets.setdefault(title, MetadataSheet(title=title))

    @property
    def properties(self):
        return self._ensure_metadata().properties

    @properties.setter
    def properties(self, value):
        self._check_writable()
        self._ensure_metadata().properties = value

    @property
    def defined_names(self):
        return self._ensure_metadata().defined_names

    @defined_names.setter
    def defined_names(self, value):
        self._check_writable()
        self._ensure_metadata().defined_names = value

    def _wrap_sheet(self, native):
        title = native.title
        if title not in self._sheets:
            from ..worksheet import Worksheet
            self._sheets[title] = Worksheet(self, native)
        return self._sheets[title]

    @property
    def sheetnames(self):
        return self._native.sheetnames

    @property
    def worksheets(self):
        return [self._wrap_sheet(sheet) for sheet in self._native.worksheets]

    @property
    def active(self):
        if self.write_only:
            return self.worksheets[0] if self.sheetnames else None
        return self._wrap_sheet(self._native.active)

    @active.setter
    def active(self, value):
        self._check_writable()
        if hasattr(value, '_native'):
            self._require_sheet(value)
            value = value._native
        self._native.active = value

    def __getitem__(self, title):
        return self._wrap_sheet(self._native[title])

    def __contains__(self, title):
        return title in self.sheetnames

    def __iter__(self):
        return iter(self.worksheets)

    def __len__(self):
        return len(self.sheetnames)

    def _require_sheet(self, worksheet):
        if getattr(worksheet, 'parent', None) is not self or worksheet.title not in self.sheetnames:
            raise ValueError('Worksheet does not belong to this workbook')

    def index(self, worksheet):
        self._require_sheet(worksheet)
        return self.sheetnames.index(worksheet.title)

    def _local_scopes(self):
        if self._book is None:
            return []
        titles = self.sheetnames
        return [(name, titles[name.localSheetId]) for name in self._book.defined_names.values()
                if name.localSheetId is not None and 0 <= name.localSheetId < len(titles)]

    def _restore_scopes(self, scopes):
        if self._book is None:
            return
        names = self._book.defined_names
        for key, name in list(names.items()):
            if name.localSheetId is not None:
                del names[key]
        for name, title in scopes:
            if title in self.sheetnames:
                name.localSheetId = self.sheetnames.index(title)
                names.add(name)

    def create_sheet(self, title=None, index=None):
        self._check_writable()
        scopes = self._local_scopes()
        result = self._wrap_sheet(self._native.create_sheet(title=title, index=index))
        self._restore_scopes(scopes)
        return result

    def copy_worksheet(self, worksheet):
        self._check_writable()
        self._require_sheet(worksheet)
        if self.write_only:
            raise TypeError('Cannot copy sheets in write-only mode')
        metadata = worksheet._metadata
        if metadata.tables or metadata.charts:
            raise NotImplementedError('Copying sheets with tables or charts is not supported')
        copied = deepcopy(metadata)
        result = self._wrap_sheet(self._native.copy_worksheet(worksheet._native))
        copied.title = result.title
        book = self._ensure_metadata()
        book.sheets[result.title] = copied
        if any(name.localSheetId == self.index(worksheet)
               for name in book.defined_names.values()):
            warnings.warn('Sheet-local defined names are not copied', RuntimeWarning, stacklevel=2)
        return result

    def remove(self, worksheet):
        self._check_writable()
        self._require_sheet(worksheet)
        title = worksheet.title
        if self._book is not None and (self._book.defined_names or
                any(s.charts for s in self._book.sheets.values())):
            raise NotImplementedError('Removing sheets with defined names or chart references is unsafe')
        self._native.remove(worksheet._native)
        self._sheets.pop(title, None)
        if self._book is not None:
            self._book.sheets.pop(title, None)

    def remove_sheet(self, worksheet):
        self.remove(self[worksheet] if isinstance(worksheet, str) else worksheet)

    def __delitem__(self, title):
        self.remove(self[title])

    def _rename_sheet(self, worksheet, title):
        self._check_writable()
        if title == worksheet.title:
            return
        book = self._book
        if book is not None and (book.defined_names or any(s.charts for s in book.sheets.values())):
            raise NotImplementedError('Renaming sheets with defined names or chart references is unsafe')
        old = worksheet.title
        worksheet._native.title = title
        self._sheets.pop(old, None)
        self._sheets[worksheet.title] = worksheet
        if book is not None and old in book.sheets:
            metadata = book.sheets.pop(old)
            metadata.title = worksheet.title
            book.sheets[worksheet.title] = metadata

    def save(self, filename):
        self._check_writable()
        if not self._overlay:
            return self._native.save(filename)
        from .._metadata import apply_metadata
        # Seekable disk-backed storage avoids a second Python copy of the XLSX.
        with TemporaryFile(mode='w+b') as generated:
            self._native.save(generated)
            generated.seek(0)
            apply_metadata(generated, filename, self._book)

    def close(self):
        if not self._closed:
            self._native.close()
            if self._owns_source and self._source is not None:
                self._source.close()
            self._closed = True

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def __repr__(self):
        return '<Workbook sheetnames=%r>' % self.sheetnames


__all__ = ['Workbook', 'NativeWorkbook', 'NativeReadOnlyWorkbook', 'NativeWriteOnlyWorkbook']
