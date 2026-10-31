"""Prepare sparse cell styles for native replacement without touching native state."""

from dataclasses import fields, is_dataclass
from functools import lru_cache
from xml.etree import ElementTree as ET

from . import _metadata
from ._models import CellMetadata, MetadataBook, MetadataSheet, coordinate_parts
from ._style_pool import FrozenStyle, unwrap_style
from .styles import STYLE_INTERNAL_FIELDS, _StyleModel

_STYLE_FIELDS = ("font", "fill", "border", "alignment", "protection")


@lru_cache(maxsize=64)
def _field_layout(model_type):
    layout = fields(model_type)
    return layout, frozenset(field.name for field in layout)


def _snapshot(value):
    """Freeze mutable model values, including XML not represented by fields.

    XML is structural rather than serialized: equal copied models share a cache
    entry without serializing every cell. Never retain an identity-based cache;
    callers can edit nested Colors, Sides, or original XML between saves.
    """
    if type(value) is FrozenStyle:
        return value.fingerprint
    if type(value) in (type(None), bool, int, float, str, bytes):
        return type(value), value
    if isinstance(value, ET.Element):
        return (
            value.tag,
            tuple(sorted(value.attrib.items())),
            value.text,
            value.tail,
            tuple(_snapshot(child) for child in value),
        )
    resolved = unwrap_style(value)
    if resolved is not value:
        return _snapshot(resolved)
    if is_dataclass(value):
        layout, names = _field_layout(type(value))
        if isinstance(value, _StyleModel):
            names = names | STYLE_INTERNAL_FIELDS
        return (
            type(value),
            tuple(
                (field.name, _snapshot(getattr(value, field.name))) for field in layout
            ),
            tuple(
                sorted(
                    (key, _snapshot(item))
                    for key, item in vars(value).items()
                    if key not in names
                )
            ),
        )
    if isinstance(value, dict):
        return (
            type(value),
            frozenset((_snapshot(key), _snapshot(item)) for key, item in value.items()),
            tuple(
                sorted(
                    (key, _snapshot(item))
                    for key, item in getattr(value, "__dict__", {}).items()
                )
            ),
        )
    if isinstance(value, (tuple, list)):
        return (
            type(value),
            tuple(_snapshot(item) for item in value),
            tuple(
                sorted(
                    (key, _snapshot(item))
                    for key, item in getattr(value, "__dict__", {}).items()
                )
            ),
        )
    if isinstance(value, (set, frozenset)):
        return (type(value), frozenset(_snapshot(item) for item in value))
    if hasattr(value, "__dict__") and not isinstance(value, type):
        return (type(value), _snapshot(vars(value)))
    return (type(value), value)


def _component_xml(name, value):
    if name == "font":
        node = _metadata._font_xml(value)
    elif name == "fill":
        node = _metadata._fill_xml(value)
    elif name == "border":
        node = _metadata._border_xml(value)
    else:
        node = _metadata._model_xml(
            name, value, rename=_metadata._ALIGN_NAMES if name == "alignment" else None
        )
    return ET.tostring(node, encoding="unicode")


def compile_style_records(metadata_book):
    """Return ``(records, assignments)`` for ``_replace_cell_styles``.

    Records contain optional font/fill/border/alignment/protection XML strings,
    in that order. Assignments are (sheet title, one-based row, column, index).
    Number formats remain entirely native-owned. All caches are local to this
    invocation, and each distinct component value is serialized only once.
    Empty results are meaningful: native replacement must clear previous styles.
    """
    records, assignments = [], []
    components, indices = {}, {}
    if metadata_book is None:
        return records, assignments
    for title, sheet in metadata_book.sheets.items():
        for coordinate, cell in sheet.cells.items():
            values = tuple(getattr(cell, name) for name in _STYLE_FIELDS)
            if all(value is None for value in values) and not cell.style_dirty:
                continue
            key = tuple(_snapshot(value) for value in values)
            if key not in indices:
                record = []
                for name, value, snapshot in zip(_STYLE_FIELDS, values, key):
                    if value is None:
                        record.append(None)
                        continue
                    component_key = (name, snapshot)
                    if component_key not in components:
                        components[component_key] = _component_xml(name, value)
                    record.append(components[component_key])
                indices[key] = len(records)
                records.append(tuple(record))
            col, row = coordinate_parts(coordinate)
            assignments.append((title, row, col, indices[key]))
    return records, assignments


def _has_extra_attributes(model):
    # Even empty loaded snapshots/parts are evidence that a native-only save
    # cannot prove it preserves all original content.
    return bool(set(vars(model)) - {field.name for field in fields(model)})


def requires_metadata_overlay(metadata_book):
    """Conservatively exclude only new, cell-style-only books from the overlay.

    Default getters may populate empty models and dimensions. Those do not need
    an overlay, but original snapshots, relationships, unknown attributes, and
    every non-cell-style feature do. This is a capability check, not dirty state.
    """
    if metadata_book is None:
        return False
    if _has_extra_attributes(metadata_book):
        return True
    default_book = MetadataBook()
    for field in fields(metadata_book):
        if field.name != "sheets" and _snapshot(
            getattr(metadata_book, field.name)
        ) != _snapshot(getattr(default_book, field.name)):
            return True
    default_cell = CellMetadata()
    for sheet in metadata_book.sheets.values():
        if _has_extra_attributes(sheet):
            return True
        default_sheet = MetadataSheet(title=sheet.title)
        for field in fields(sheet):
            name = field.name
            value = getattr(sheet, name)
            if name == "cells":
                for cell in value.values():
                    if _has_extra_attributes(cell):
                        return True
                    for cell_field in fields(cell):
                        if cell_field.name not in (
                            *_STYLE_FIELDS,
                            "style_dirty",
                        ) and _snapshot(getattr(cell, cell_field.name)) != _snapshot(
                            getattr(default_cell, cell_field.name)
                        ):
                            return True
            elif name in ("column_dimensions", "row_dimensions"):
                default = getattr(default_sheet, name).factory()
                if any(
                    _snapshot(dimension) != _snapshot(default)
                    for dimension in value.values()
                ):
                    return True
                if vars(value) != vars(getattr(default_sheet, name)):
                    return True
            elif _snapshot(value) != _snapshot(getattr(default_sheet, name)):
                return True
    return False
