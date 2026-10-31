"""Lazy decoding of native source styles, without a Python worksheet grid."""

from copy import deepcopy
from dataclasses import fields
from xml.etree import ElementTree as ET

from ._compat import STYLE_DEFAULTS
from ._models import CellMetadata, column_letter
from ._native_styles import _snapshot
from ._style_pool import FrozenStyle
from .styles import _STYLE_TYPES, Alignment, Protection

STYLE_FIELDS = tuple(STYLE_DEFAULTS)
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_COLOR = {"rgb", "theme", "indexed", "auto", "tint"}
_SIDES = {
    "start",
    "end",
    "left",
    "right",
    "top",
    "bottom",
    "diagonal",
    "vertical",
    "horizontal",
}
_FONT = {
    "name",
    "sz",
    "b",
    "i",
    "u",
    "strike",
    "vertAlign",
    "charset",
    "family",
    "scheme",
    "outline",
    "shadow",
    "condense",
    "extend",
}
# Exact namespace, attribute and child checks are intentional. A familiar local
# name does not prove a producer's extension can be represented by our models.
_SCHEMA = {
    "font": (set(), _FONT | {"color"}),
    "fill": (set(), {"patternFill"}),
    "patternFill": ({"patternType"}, {"fgColor", "bgColor"}),
    "border": ({"diagonalUp", "diagonalDown", "outline"}, _SIDES),
    "alignment": (
        {
            "horizontal",
            "vertical",
            "textRotation",
            "wrapText",
            "shrinkToFit",
            "indent",
            "relativeIndent",
            "justifyLastLine",
            "readingOrder",
        },
        set(),
    ),
    "protection": ({"locked", "hidden"}, set()),
    **{tag: (_COLOR, set()) for tag in ("color", "fgColor", "bgColor")},
    **{tag: ({"style"}, {"color"}) for tag in _SIDES},
    **{tag: ({"val"}, set()) for tag in _FONT},
}


def _clean_xml(node):
    if not isinstance(node.tag, str) or not node.tag.startswith(_NS):
        return False
    shape = _SCHEMA.get(node.tag[len(_NS) :])
    if shape is None or set(node.attrib) - shape[0]:
        return False
    if (node.text or "").strip() or (node.tail or "").strip():
        return False
    seen = set()
    for child in node:
        if child.tag in seen or child.tag not in {_NS + tag for tag in shape[1]}:
            return False
        seen.add(child.tag)
        if not _clean_xml(child):
            return False
    return True


def _strip_xml(model):
    if type(model) not in _STYLE_TYPES:
        return
    vars(model).pop("_xml", None)
    for field in fields(model):
        _strip_xml(getattr(model, field.name))


def decode_record(record, pool):
    """Decode once per native registry index; freeze only fully modeled XML."""
    from . import _metadata as xml

    result = []
    parsers = (xml._read_font, xml._read_fill, xml._read_border)
    for index, (name, data) in enumerate(zip(STYLE_FIELDS, record)):
        if data is None:
            result.append(None)
            continue
        node = xml.xml(data.encode("utf-8") if isinstance(data, str) else data)
        if index < 3:
            value = parsers[index](node)
        else:
            cls = Alignment if name == "alignment" else Protection
            value = xml._model(
                cls, node, xml._ALIGN_NAMES if name == "alignment" else None
            )
        if node.tag == _NS + name and _clean_xml(node):
            _strip_xml(value)
        result.append(pool.assign(value))
    return tuple(result)


def _merge_unknown(original, generated):
    """Keep extension data on regenerated known children as well as roots."""
    shape = _SCHEMA.get(original.tag.removeprefix(_NS))
    if shape is None:
        return generated
    for key, value in original.attrib.items():
        if key not in shape[0]:
            generated.set(key, value)
    if (original.text or "").strip():
        generated.text = original.text
    if (original.tail or "").strip():
        generated.tail = original.tail
    seen = set()
    for child in original:
        if child.tag in seen or child.tag not in {_NS + tag for tag in shape[1]}:
            if not any(ET.tostring(item) == ET.tostring(child) for item in generated):
                generated.append(deepcopy(child))
        else:
            seen.add(child.tag)
            target = generated.find(child.tag)
            if target is not None:
                _merge_unknown(child, target)
    return generated


def preserve_opaque_records(book, records, assignments):
    """The legacy serializers rebuild known children; retain their extensions."""
    patched = list(records)
    done = set()
    for title, row, col, index in assignments:
        if index in done:
            continue
        done.add(index)
        cell = book.sheets[title].cells[column_letter(col) + str(row)]
        record = list(records[index])
        for position, name in enumerate(STYLE_FIELDS):
            value = getattr(cell, name)
            original = getattr(value, "_xml", None)
            if original is not None:
                generated = ET.fromstring(record[position])
                record[position] = ET.tostring(
                    _merge_unknown(original, generated), encoding="unicode"
                )
        patched[index] = tuple(record)
    return patched


class NativeStyleReader:
    def __init__(self, workbook):
        self.workbook = workbook
        self.records = {}
        self.opaque_cells = {}

    def source(self, cell):
        index = cell._native._source_style_id
        if index is None:
            return (None,) * len(STYLE_FIELDS)
        if index not in self.records:
            self.records[index] = decode_record(
                self.workbook._native._source_style_record(index),
                self.workbook._style_pool,
            )
        return self.records[index]

    def _metadata(self, cell):
        book = self.workbook._book
        sheet = book.sheets.get(cell.parent.title) if book is not None else None
        return sheet.cells.get(cell.coordinate) if sheet is not None else None

    def _default(self, name):
        book = self.workbook
        if name not in book._style_defaults:
            book._style_defaults[name] = book._style_pool.assign(STYLE_DEFAULTS[name]())
        return book._style_defaults[name]

    def get(self, cell, name):
        metadata = self._metadata(cell)
        if metadata is not None and metadata.style_dirty:
            value = getattr(metadata, name)
            return value if value is not None else self._default(name)
        value = self.source(cell)[STYLE_FIELDS.index(name)]
        if value is None:
            return self._default(name)
        if type(value) is FrozenStyle:
            return value
        key = (cell.parent.title, cell.coordinate, name)
        if key not in self.opaque_cells:
            copied = deepcopy(value)
            self.opaque_cells[key] = (cell, copied, _snapshot(copied))
        return self.opaque_cells[key][1]

    def materialize(self, cell):
        metadata = self._metadata(cell)
        if metadata is not None and metadata.style_dirty:
            return metadata
        sheet = self.workbook._sheet_metadata(cell.parent.title, full=False)
        metadata = sheet.cells.setdefault(cell.coordinate, CellMetadata())
        for name, value in zip(STYLE_FIELDS, self.source(cell)):
            cached = self.opaque_cells.get((cell.parent.title, cell.coordinate, name))
            setattr(metadata, name, cached[1] if cached else deepcopy(value))
        metadata.style_dirty = True
        return metadata

    def set(self, cell, name, value):
        setattr(self.materialize(cell), name, value)

    def flush(self, *, retain_all=False):
        for cell, value, baseline in list(self.opaque_cells.values()):
            metadata = self._metadata(cell)
            if metadata is not None and metadata.style_dirty:
                continue
            if retain_all or _snapshot(value) != baseline:
                self.materialize(cell)
