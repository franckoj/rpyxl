"""Native workbooks with sparse, persistent compatibility metadata."""

import warnings
from copy import deepcopy
from io import BytesIO
from tempfile import TemporaryFile

from .._core import ReadOnlyWorkbook as NativeReadOnlyWorkbook
from .._core import Workbook as NativeWorkbook
from .._core import WriteOnlyWorkbook as NativeWriteOnlyWorkbook
from .._models import MetadataBook, MetadataSheet
from .._style_pool import StylePool


class _WorkbookMeta(type):
    def __instancecheck__(cls, instance):
        return type.__instancecheck__(cls, instance) or isinstance(
            instance, (NativeWorkbook, NativeWriteOnlyWorkbook, NativeReadOnlyWorkbook)
        )


class Workbook(metaclass=_WorkbookMeta):
    """Keep values in Rust and materialize feature models only when needed.

    New value-only and cell-style-only books save in one pass through Rust.
    Mutable style models are snapshotted and deduplicated before transfer, so
    nested in-place edits remain visible without serializing XML for each cell.
    Loaded styles decode lazily from native records; style/value-only edits use
    native package preservation when available. Rich features and unsupported
    templates retain the ZIP/XML overlay, without rebuilding cell styles.

    Read-only metadata is deferred, not sheet-selective: first access warns and
    materializes all metadata. Value iteration stays native. Structural edits
    move sparse cell styles/annotations and dimensions, but reject reference-
    bearing features rather than guessing how to translate them. Renames reject
    books with names/charts; sheet copies reject tables/charts and do not copy
    sheet-local names. Formula translation and full merged-border propagation
    are not implemented. Native* exports provide explicit unwrapped access.
    """

    def __init__(
        self, write_only=False, *, _native=None, _source=None, _owns_source=False
    ):
        self._native = (
            _native
            if _native is not None
            else (NativeWriteOnlyWorkbook() if write_only else NativeWorkbook())
        )
        self._source = _source
        self._owns_source = _owns_source
        self._book = None
        self._full_metadata_loaded = False
        self._structure_changed = False
        self._native_package_api = (
            not self.read_only
            and not self.write_only
            and all(
                hasattr(self._native, name)
                for name in (
                    "_has_original_package",
                    "_can_preserve_package",
                    "_original_package",
                    "_source_style_record",
                    "_update_cell_styles",
                )
            )
        )
        self._has_original = _source is not None or (
            self._native_package_api and self._native._has_original_package()
        )
        self._style_pool = StylePool()
        self._style_defaults = {}
        self._sheets = {}
        self._overlay = False
        self._closed = False
        self._native_style_reader = None
        if self._native_package_api and self._has_original:
            from .._native_style_read import NativeStyleReader

            self._native_style_reader = NativeStyleReader(self)
        if _source is not None and not self.read_only and not self._native_package_api:
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
        return getattr(self._native, "write_only", False)

    def _check_writable(self):
        if self._closed:
            raise ValueError("Workbook is closed")
        if self.read_only:
            raise TypeError("Cannot modify a read-only workbook")

    def _ensure_metadata(self, full=True, *, load_original=None):
        if self._closed:
            raise ValueError("Workbook is closed")
        if load_original is not None:
            full = load_original
        full = full or self.read_only
        if full and not self._full_metadata_loaded:
            if self._has_original:
                if self.read_only:
                    warnings.warn(
                        "Accessing read-only metadata loads metadata for all sheets; "
                        "values_only iteration does not incur this cost.",
                        RuntimeWarning,
                        stacklevel=3,
                    )
                from .._metadata import load_metadata

                if self._native_style_reader is not None:
                    self._native_style_reader.flush()
                source = self._source
                if self._native_package_api:
                    source = BytesIO(self._native._original_package())
                    loaded = load_metadata(source, cell_styles=False)
                else:
                    loaded = load_metadata(source)
                # Original annotations must survive promotion from a sparse
                # style sidecar to the full feature model.
                if self._book is not None:
                    from .._native_style_read import STYLE_FIELDS

                    for title, sheet in self._book.sheets.items():
                        target = loaded.sheets.setdefault(
                            title, MetadataSheet(title=title)
                        )
                        for coordinate, cell in sheet.cells.items():
                            original = target.cells.get(coordinate)
                            if original is None:
                                target.cells[coordinate] = cell
                            elif cell.style_dirty:
                                for name in STYLE_FIELDS:
                                    setattr(original, name, getattr(cell, name))
                                original.style_dirty = True
                self._book = loaded
            self._full_metadata_loaded = True
        if self._book is None:
            self._book = MetadataBook()
        self._overlay = True
        return self._book

    def _sheet_metadata(self, title, *, full=True):
        book = self._ensure_metadata(full=full)
        metadata = book.sheets.get(title)
        if metadata is None:
            metadata = book.sheets[title] = MetadataSheet(title=title)
        return metadata

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
        if hasattr(value, "_native"):
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
        if (
            getattr(worksheet, "parent", None) is not self
            or worksheet.title not in self.sheetnames
        ):
            raise ValueError("Worksheet does not belong to this workbook")

    def index(self, worksheet):
        self._require_sheet(worksheet)
        return self.sheetnames.index(worksheet.title)

    def _prepare_structure(self):
        if self._has_original:
            self._ensure_metadata()
        if self._native_style_reader is not None:
            # Callers can retain an opaque model and mutate it after a sheet
            # operation. Attach those exact objects before dropping title keys.
            self._native_style_reader.flush(retain_all=True)
            self._native_style_reader.opaque_cells.clear()

    def _local_scopes(self):
        self._prepare_structure()
        if self._book is None:
            return []
        titles = self.sheetnames
        return [
            (name, titles[name.localSheetId])
            for name in self._book.defined_names.values()
            if name.localSheetId is not None and 0 <= name.localSheetId < len(titles)
        ]

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
        self._structure_changed = self._has_original
        return result

    def copy_worksheet(self, worksheet):
        self._check_writable()
        self._require_sheet(worksheet)
        if self.write_only:
            raise TypeError("Cannot copy sheets in write-only mode")
        self._prepare_structure()
        metadata = worksheet._metadata
        if metadata.tables or metadata.charts:
            raise NotImplementedError(
                "Copying sheets with tables or charts is not supported"
            )
        copied = deepcopy(metadata)
        result = self._wrap_sheet(self._native.copy_worksheet(worksheet._native))
        self._structure_changed = self._has_original
        copied.title = result.title
        book = self._ensure_metadata()
        book.sheets[result.title] = copied
        if any(
            name.localSheetId == self.index(worksheet)
            for name in book.defined_names.values()
        ):
            warnings.warn(
                "Sheet-local defined names are not copied", RuntimeWarning, stacklevel=2
            )
        return result

    def remove(self, worksheet):
        self._check_writable()
        self._require_sheet(worksheet)
        self._prepare_structure()
        title = worksheet.title
        if self._book is not None and (
            self._book.defined_names
            or any(s.charts for s in self._book.sheets.values())
        ):
            raise NotImplementedError(
                "Removing sheets with defined names or chart references is unsafe"
            )
        self._native.remove(worksheet._native)
        self._structure_changed = self._has_original
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
        self._prepare_structure()
        book = self._book
        if book is not None and (
            book.defined_names or any(s.charts for s in book.sheets.values())
        ):
            raise NotImplementedError(
                "Renaming sheets with defined names or chart references is unsafe"
            )
        old = worksheet.title
        worksheet._native.title = title
        self._structure_changed = self._has_original
        self._sheets.pop(old, None)
        self._sheets[worksheet.title] = worksheet
        if book is not None and old in book.sheets:
            metadata = book.sheets.pop(old)
            metadata.title = worksheet.title
            book.sheets[worksheet.title] = metadata

    def save(self, filename):
        self._check_writable()
        if self._native_style_reader is not None:
            self._native_style_reader.flush()
        if self._native_package_api:
            from .._native_styles import requires_metadata_overlay

            needs_overlay = self._structure_changed or requires_metadata_overlay(
                self._book
            )
            can_preserve = (
                not self._has_original or self._native._can_preserve_package()
            )
            if self._has_original and (needs_overlay or not can_preserve):
                self._ensure_metadata()
            elif not self._overlay:
                return self._native.save(filename)
        elif not self._overlay:
            return self._native.save(filename)
        from .._metadata import apply_metadata

        native_styles = not self.write_only
        if native_styles:
            from .._native_styles import (
                compile_style_records,
                requires_metadata_overlay,
            )

            records, assignments = compile_style_records(self._book)
            if self._native_package_api:
                from .._native_style_read import preserve_opaque_records

                records = preserve_opaque_records(self._book, records, assignments)
                self._native._update_cell_styles(records, assignments)
            else:
                self._native._replace_cell_styles(records, assignments)
            if not self._structure_changed and not requires_metadata_overlay(
                self._book
            ):
                return self._native.save(filename)
        # Only features not yet handled by Rust need an additional ZIP/XML pass.
        with TemporaryFile(mode="w+b") as generated:
            if self._native_package_api:
                self._native.save(generated, preserve_original=False)
            else:
                self._native.save(generated)
            generated.seek(0)
            apply_metadata(generated, filename, self._book, native_styles=native_styles)

    def close(self):
        if not self._closed:
            self._native.close()
            if self._owns_source and self._source is not None:
                self._source.close()
            if self._native_style_reader is not None:
                self._native_style_reader.records.clear()
                self._native_style_reader.opaque_cells.clear()
            self._closed = True

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def __repr__(self):
        return f"<Workbook sheetnames={self.sheetnames!r}>"


__all__ = [
    "NativeReadOnlyWorkbook",
    "NativeWorkbook",
    "NativeWriteOnlyWorkbook",
    "Workbook",
]
