"""OOXML metadata overlay, independent of the native value reader/writer.

Contract::

    book = load_metadata(path_or_seekable_binary_stream)
    book.sheets["Sheet1"].cells.setdefault("A1", CellMetadata()).font = Font(b=True)
    apply_metadata(native_xlsx_path_or_stream, output_path_or_binary_stream, book)

For ``apply_metadata``, ``source`` is a complete native-generated XLSX;
``load_metadata`` instead reads the original workbook.
The overlay never changes cell values, formulas, shared strings or numFmtIds.
Sheets are matched by title; metadata never creates or renames native sheets.
A MetadataBook is authoritative for its modeled fields on included sheets;
omitted sheets are untouched. None clears an optional field (except cell style
components, where None means do not override that component).

Loaded models retain private XML snapshots for unmodeled attributes/elements.
Drawing dependencies are retained without decoding image bytes. Unsupported
charts are opaque objects with an ``anchor`` and can be retained or removed, not
edited as BarCharts. This module does not resolve external relationships.

Public model contract (all feature classes live in ``rpyxl._models``):

* MetadataBook.properties is DocumentProperties; defined_names is a mapping
  with add(DefinedName); sheet-local keys are (name, localSheetId).
* MetadataBook.sheets maps exact native titles to MetadataSheet instances.
* MetadataSheet.cells maps A1 coordinates to CellMetadata. Optional style
  components are Font, PatternFill, Border, Alignment, and Protection;
  hyperlinks accept a URL string or Hyperlink, comments accept Comment.
* merged_cells.add(range) and merged_cells.ranges expose merged ranges.
  column_dimensions[letter] and row_dimensions[int] auto-create dimensions.
* freeze_panes accepts an A1 coordinate; auto_filter.ref and
  sheet_properties.tabColor set the filter range and tab color.
* protection is SheetProtection (set_password hashes legacy passwords);
  page_setup is PageSetup; sheet_state is visible, hidden, or veryHidden.
* data_validations.add(DataValidation), conditional_formatting.add(range,
  Rule), tables.add(Table), and charts.append(BarChart) declare features.
  MetadataSheet.add_chart(chart, anchor) is also available.

Limitations: only transitional OOXML worksheets are supported. No formula
calculation, chart caches, sheet creation/renaming, threaded comments, new
images, or full openpyxl object graph is provided. Built-in table styles work;
custom table-style definitions and advanced extension features with references
into the original style table are not translated to native style indices.
Original themes are not substituted, so theme colors use the native theme.
Unmodeled XML is retained on a best-effort basis, not a lossless XLSX editor.
Unused source package parts are preserved; this is not a secure-data scrubber.

ZIP paths are validated, decompression is bounded, DTDs/entities are rejected,
and each input archive is opened once per operation. Caller streams stay open
and their input positions are restored. Output streams must support seek and
truncate; source and target may be the same path or stream.
"""
import os
import posixpath
import re
from copy import deepcopy
from dataclasses import fields
from datetime import datetime, timezone
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from ._models import (
    AutoFilter,
    BarChart,
    CellMetadata,
    ColumnDimension,
    Comment,
    DataValidation,
    DefinedName,
    Hyperlink,
    MetadataBook,
    MetadataSheet,
    PageSetup,
    RowDimension,
    Rule,
    Series,
    SheetProtection,
    Table,
    TableColumn,
    TableStyleInfo,
    column_index,
    column_letter,
    coordinate_parts,
    range_bounds,
)
from .styles import (
    Alignment,
    Border,
    Color,
    Font,
    PatternFill,
    Protection,
    Side,
    _color,
)

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
X = "urn:schemas-microsoft-com:office:excel"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCT = "http://purl.org/dc/terms/"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
for prefix, uri in (("", S), ("r", R), ("c", C), ("a", A), ("xdr", XDR),
                    ("v", V), ("o", O), ("x", X), ("cp", CP), ("dc", DC),
                    ("dcterms", DCT), ("xsi", XSI), ("vt", VT)):
    ET.register_namespace(prefix, uri)

MAX_PART_BYTES = 256 * 1024 * 1024
MAX_ARCHIVE_BYTES = 1024 * 1024 * 1024
MAX_MEMBERS = 100000


def q(tag, ns=S):
    return "{%s}%s" % (ns, tag)


def local(node):
    return node.tag.rsplit("}", 1)[-1]


def _validate_xml(data):
    # Reject UTF-16/32 DTDs too, rather than depending on parser defaults.
    probe = data.replace(b"\x00", b"").upper()
    if b"<!DOCTYPE" in probe or b"<!ENTITY" in probe:
        raise ValueError("DTD/entity declarations are not allowed in XLSX metadata")


def xml(data):
    _validate_xml(data)
    return ET.fromstring(data)


def _worksheet_metadata_xml(data):
    """Parse worksheet metadata without building a cell-value DOM.

    Retain only r/s attributes on styled cells, all row attributes, and all
    non-cell metadata. Completed elements are detached from the parser tree;
    retained stubs are buffered separately until their row/sheetData closes.
    This also avoids scanning past every previously retained cell/row during
    parent.remove(). XML/security errors deliberately propagate to the caller.
    """
    _validate_xml(data)
    sheet_data_tag, row_tag, cell_tag = q("sheetData"), q("row"), q("c")
    stack = []
    cell = None
    row_children = []
    sheet_children = []
    with BytesIO(data) as stream:
        parser = ET.iterparse(stream, events=("start", "end"))
        for event, node in parser:
            if event == "start":
                stack.append(node)
                if len(stack) == 2 and node.tag == sheet_data_tag:
                    sheet_children = []
                elif len(stack) == 3 and stack[1].tag == sheet_data_tag and node.tag == row_tag:
                    row_children = []
                elif (len(stack) == 4 and stack[1].tag == sheet_data_tag
                      and stack[2].tag == row_tag and node.tag == cell_tag):
                    cell = node
                continue

            if cell is not None:
                if node is cell:
                    if "s" in node.attrib:
                        row_children.append(ET.Element(cell_tag, {
                            key: value for key, value in node.attrib.items() if key in ("r", "s")
                        }))
                    cell = None
                # Drop payload descendants as they finish, not just at </c>.
                stack[-2].remove(node)
                node.clear()
            elif len(stack) == 4 and stack[1].tag == sheet_data_tag and stack[2].tag == row_tag:
                # Preserve unknown non-cell row children, including extensions.
                stack[-2].remove(node)
                row_children.append(node)
            elif len(stack) == 3 and stack[1].tag == sheet_data_tag:
                stack[-2].remove(node)
                if node.tag == row_tag:
                    node.extend(row_children)
                    row_children = []
                    if not (len(node) or any(key != "r" for key in node.attrib)
                            or (node.text and node.text.strip())):
                        node.clear()
                        stack.pop()
                        continue
                sheet_children.append(node)
            elif len(stack) == 2 and node.tag == sheet_data_tag:
                node.extend(sheet_children)
                sheet_children = []
            stack.pop()
        return parser.root


def dump(root):
    # ElementTree rewrites extension prefixes. Keep mc:Ignorable consistent
    # with the namespace declarations actually emitted, not stale input names.
    root = deepcopy(root)
    ignorable = q("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006")
    for node in root.iter():
        node.attrib.pop(ignorable, None)
    # Include mc during both serializations so auto-generated nsN prefixes do
    # not shift when the final Ignorable attribute is added.
    root.set(ignorable, "")
    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    extensions = []
    for prefix, uri in re.findall(rb'xmlns:([^= ]+)="([^"]+)"', data):
        if (b"/office/spreadsheetml/" in uri or b"/office/drawing/" in uri
                or b"/office/excel/" in uri):
            extensions.append(prefix.decode("ascii"))
    if extensions:
        root.set(ignorable, " ".join(extensions))
    else:
        root.attrib.pop(ignorable)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def scalar(value):
    if isinstance(value, bool):
        return "1" if value else "0"
    return str(value)


def boolean(value):
    return str(value).lower() in ("1", "true", "on")


def attrs(obj, exclude=(), rename=None):
    rename = rename or {}
    return {rename.get(f.name, f.name): scalar(getattr(obj, f.name)) for f in fields(obj)
            if f.name not in exclude and getattr(obj, f.name) is not None}


def _required(parent, tag) -> ET.Element:
    node = parent.find(q(tag))
    if node is None:
        raise ValueError("Missing required OOXML element: " + tag)
    return node


def _remember(obj, node):
    obj._xml = deepcopy(node)
    return obj


def _model(cls, node, rename=None, exclude=()):
    rename = rename or {}
    result = cls()
    for f in fields(result):
        if f.name in exclude:
            continue
        value = node.get(rename.get(f.name, f.name))
        if value is not None:
            annotation = str(f.type)
            if "bool" in annotation:
                value = boolean(value)
            elif "int" in annotation:
                value = int(value)
            elif "float" in annotation:
                value = float(value)
            setattr(result, f.name, value)
    return _remember(result, node)


def _model_xml(tag, obj, exclude=(), rename=None):
    node = deepcopy(getattr(obj, "_xml", ET.Element(q(tag))))
    rename = rename or {}
    for f in fields(obj):
        if f.name not in exclude:
            node.attrib.pop(rename.get(f.name, f.name), None)
    node.attrib.update(attrs(obj, exclude, rename))
    return node


def _path(base, target):
    if "\\" in target or "?" in target or "#" in target:
        raise ValueError("Invalid package relationship target: " + target)
    value = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(base), target))
    if value == ".." or value.startswith("../") or ":" in value:
        raise ValueError("Relationship escapes package")
    return value


def _rels_path(part):
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


class _Archive:
    def __init__(self, source):
        self.stream = source if hasattr(source, "read") else None
        self.position = self.stream.tell() if self.stream is not None else None
        try:
            self.zip = ZipFile(source, "r")
        except BaseException:
            if self.stream is not None:
                self.stream.seek(self.position)
            raise
        try:
            infos = self.zip.infolist()
            if len(infos) > MAX_MEMBERS or sum(i.file_size for i in infos) > MAX_ARCHIVE_BYTES:
                raise ValueError("XLSX archive exceeds metadata resource limits")
            names = set()
            for info in infos:
                name = info.filename
                if (name.startswith("/") or "\\" in name or ":" in name
                        or ".." in name.split("/") or name in names
                        or info.file_size > MAX_PART_BYTES):
                    raise ValueError("Unsafe, duplicate, or oversized ZIP member: " + name)
                names.add(name)
            self.names = names
            self.cache = {}
        except BaseException:
            self.close()
            raise

    def read(self, name):
        if name not in self.cache:
            self.cache[name] = self.zip.read(name)
        return self.cache[name]

    def root(self, name, default=None):
        return xml(self.read(name)) if name in self.names else default

    def close(self):
        self.zip.close()
        if self.stream is not None:
            self.stream.seek(self.position)


def _relationships(archive, part):
    root = archive.root(_rels_path(part), ET.Element(q("Relationships", P)))
    return root, {n.get("Id"): n for n in root}


def _target(part, rel):
    if rel is None or rel.get("TargetMode") == "External":
        return None
    return _path(part, rel.get("Target", ""))


def _find_rel(rels, suffix):
    return next((rel for rel in rels.values() if rel.get("Type", "").endswith("/" + suffix)), None)


def _book_parts(archive):
    _, rels = _relationships(archive, "")
    part = _target("", _find_rel(rels, "officeDocument")) or "xl/workbook.xml"
    workbook = archive.root(part)
    if workbook is None:
        raise ValueError("XLSX workbook part is missing")
    if workbook.tag != q("workbook"):
        raise ValueError("Only transitional OOXML workbooks are supported")
    _, workbook_rels = _relationships(archive, part)
    sheets = []
    for node in workbook.findall(q("sheets") + "/" + q("sheet")):
        rel = workbook_rels.get(node.get(q("id", R)))
        if rel is not None and not rel.get("Type", "").endswith("/worksheet"):
            raise ValueError("Metadata supports worksheets, not chart/dialog sheets")
        path = _target(part, rel)
        if path is not None:
            sheets.append((node, path))
    style_part = _target(part, _find_rel(workbook_rels, "styles"))
    return part, workbook, sheets, style_part


def _read_color(node):
    if node is None:
        return None
    return _remember(Color(rgb=node.get("rgb"), theme=int(node.get("theme")) if node.get("theme") is not None else None,
                           indexed=int(node.get("indexed")) if node.get("indexed") is not None else None,
                           auto=boolean(node.get("auto")) if node.get("auto") is not None else None,
                           tint=float(node.get("tint", 0))), node)


def _color_xml(tag, value):
    color = _color(value)
    node = ET.Element(q(tag))
    if color is not None:
        node.attrib.update(attrs(color))
        if not color.tint:
            node.attrib.pop("tint", None)
    return node


_FONT_TAGS = {"name": "name", "size": "sz", "bold": "b", "italic": "i", "underline": "u",
              "strike": "strike", "vertAlign": "vertAlign", "charset": "charset", "family": "family",
              "scheme": "scheme", "outline": "outline", "shadow": "shadow", "condense": "condense", "extend": "extend"}
_FONT_BOOLS = {"bold", "italic", "strike", "outline", "shadow", "condense", "extend"}
_ALIGN_NAMES = {"text_rotation": "textRotation", "wrap_text": "wrapText", "shrink_to_fit": "shrinkToFit"}
_SIDES = ("start", "end", "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal")


def _read_font(node):
    font = Font(name=None, size=None)
    for name, tag in _FONT_TAGS.items():
        child = node.find(q(tag))
        if child is not None:
            value = child.get("val", "single" if name == "underline" else "1")
            if name in _FONT_BOOLS:
                value = boolean(value)
            elif name in ("size", "family"):
                value = float(value)
            elif name == "charset":
                value = int(value)
            setattr(font, name, value)
    font.color = _read_color(node.find(q("color")))
    return _remember(font, node)


def _font_xml(font):
    node = deepcopy(getattr(font, "_xml", ET.Element(q("font"))))
    known = set(_FONT_TAGS.values()) | {"color"}
    node[:] = [n for n in node if local(n) not in known]
    for name, tag in _FONT_TAGS.items():
        value = getattr(font, name)
        if value is not None and not (name in _FONT_BOOLS and not value):
            ET.SubElement(node, q(tag), {"val": scalar(value)})
    if font.color is not None:
        node.append(_color_xml("color", font.color))
    return node


def _read_fill(node):
    pattern = node.find(q("patternFill"))
    fill = PatternFill()
    if pattern is not None:
        fill.fill_type = pattern.get("patternType")
        fill.fgColor, fill.bgColor = _read_color(pattern.find(q("fgColor"))), _read_color(pattern.find(q("bgColor")))
    return _remember(fill, node)


def _fill_xml(fill):
    node = deepcopy(getattr(fill, "_xml", ET.Element(q("fill"))))
    if fill.fill_type is None and fill.fgColor is None and fill.bgColor is None and node.find(q("gradientFill")) is not None:
        return node
    node[:] = []
    pattern = ET.SubElement(node, q("patternFill"))
    if fill.fill_type is not None:
        pattern.set("patternType", fill.fill_type)
    for tag, value in (("fgColor", fill.fgColor), ("bgColor", fill.bgColor)):
        if value is not None:
            pattern.append(_color_xml(tag, value))
    return node


def _read_border(node):
    border = _model(Border, node, exclude=_SIDES + ("diagonal_direction",))
    for name in _SIDES:
        side = node.find(q(name))
        if side is not None:
            setattr(border, name, _remember(Side(side.get("style"), _read_color(side.find(q("color")))), side))
    return border


def _border_xml(border):
    node = _model_xml("border", border, exclude=_SIDES + ("diagonal_direction",))
    node[:] = [n for n in node if local(n) not in _SIDES]
    for name in _SIDES:
        side = getattr(border, name)
        if side is not None:
            child = ET.SubElement(node, q(name))
            if side.style is not None:
                child.set("style", side.style)
            if side.color is not None:
                child.append(_color_xml("color", side.color))
    return node


class _Styles:
    def __init__(self, root=None):
        self.root = root if root is not None else ET.Element(q("styleSheet"))
        for tag in ("fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs"):
            if self.root.find(q(tag)) is None:
                self.root.append(ET.Element(q(tag), {"count": "0"}))
        for tag, child in (("fonts", _font_xml(Font())), ("fills", _fill_xml(PatternFill())),
                           ("borders", _border_xml(Border())), ("cellStyleXfs", ET.Element(q("xf"))),
                           ("cellXfs", ET.Element(q("xf"), {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"}))):
            if not len(_required(self.root, tag)):
                _required(self.root, tag).append(child)
        fills = _required(self.root, "fills")
        if len(fills) == 1:
            fills.append(_fill_xml(PatternFill(patternType="gray125")))
        self.indices = {}

    def add(self, tag, node):
        parent = _required(self.root, tag)
        key = (tag, ET.tostring(node))
        if key not in self.indices:
            for index, child in enumerate(parent):
                if ET.tostring(child) == key[1]:
                    self.indices[key] = index
                    break
            else:
                self.indices[key] = len(parent)
                parent.append(node)
        return self.indices[key]

    def read(self, index):
        xfs = _required(self.root, "cellXfs")
        if not 0 <= index < len(xfs):
            raise ValueError("Cell references invalid style index")
        xf = xfs[index]
        result = CellMetadata()
        for name, group, parser in (("font", "fonts", _read_font), ("fill", "fills", _read_fill), ("border", "borders", _read_border)):
            nodes = _required(self.root, group)
            style_index = int(xf.get(name + "Id", 0))
            if 0 <= style_index < len(nodes):
                setattr(result, name, parser(nodes[style_index]))
        for name, cls in (("alignment", Alignment), ("protection", Protection)):
            child = xf.find(q(name))
            if child is not None:
                setattr(result, name, _model(cls, child, _ALIGN_NAMES if name == "alignment" else None))
        return result

    def overlay(self, cell, metadata):
        if not any(getattr(metadata, name) is not None for name in ("font", "fill", "border", "alignment", "protection")):
            return
        xfs = _required(self.root, "cellXfs")
        index = int(cell.get("s", "0"))
        if not 0 <= index < len(xfs):
            raise ValueError("Native cell references invalid style index")
        # Crucially retain the native numFmtId, quotePrefix, xfId, etc.
        xf = deepcopy(xfs[index])
        for name, group, writer in (("font", "fonts", _font_xml), ("fill", "fills", _fill_xml), ("border", "borders", _border_xml)):
            value = getattr(metadata, name)
            if value is not None:
                xf.set(name + "Id", str(self.add(group, writer(value))))
                xf.set("apply" + name.capitalize(), "1")
        for name in ("alignment", "protection"):
            value = getattr(metadata, name)
            if value is not None:
                for old in xf.findall(q(name)):
                    xf.remove(old)
                xf.append(_model_xml(name, value, rename=_ALIGN_NAMES if name == "alignment" else None))
                xf.set("apply" + name.capitalize(), "1")
        # alignment precedes protection, extLst comes last.
        xf[:] = sorted(xf, key=lambda n: {"alignment": 0, "protection": 1, "extLst": 2}.get(local(n), 2))
        cell.set("s", str(self.add("cellXfs", xf)))

    def finish(self):
        order = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst")
        self.root[:] = sorted(self.root, key=lambda n: order.index(local(n)) if local(n) in order else len(order))
        for tag in ("fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs"):
            parent = _required(self.root, tag)
            parent.set("count", str(len(parent)))
        return dump(self.root)


def _read_rule(node, styles):
    rule = Rule(node.get("type", "expression"), operator=node.get("operator"),
                formula=[n.text or "" for n in node.findall(q("formula"))],
                stopIfTrue=boolean(node.get("stopIfTrue")) if "stopIfTrue" in node.attrib else None,
                priority=int(node.get("priority", "0")))
    dxfs = _required(styles.root, "dxfs")
    if node.get("dxfId") is not None:
        index = int(node.get("dxfId"))
        if 0 <= index < len(dxfs):
            rule._dxf_xml = deepcopy(dxfs[index])
            for name, parser in (("font", _read_font), ("fill", _read_fill), ("border", _read_border)):
                child = dxfs[index].find(q(name))
                if child is not None:
                    setattr(rule, name, parser(child))
    scale = node.find(q("colorScale"))
    if scale is not None:
        stops, colors = scale.findall(q("cfvo")), scale.findall(q("color"))
        rule.colorScale = [(stop.get("type"), stop.get("val"), _read_color(color)) for stop, color in zip(stops, colors)]
    return _remember(rule, node)


def _rule_xml(rule, styles, priority):
    node = deepcopy(getattr(rule, "_xml", ET.Element(q("cfRule"))))
    for key, value in (("type", rule.type), ("operator", rule.operator), ("stopIfTrue", rule.stopIfTrue)):
        node.attrib.pop(key, None)
        if value is not None:
            node.set(key, scalar(value))
    node.set("priority", str(rule.priority or priority))
    node.attrib.pop("dxfId", None)
    dxf = deepcopy(getattr(rule, "_dxf_xml", ET.Element(q("dxf"))))
    for name, writer in (("font", _font_xml), ("fill", _fill_xml), ("border", _border_xml)):
        for child in dxf.findall(q(name)):
            dxf.remove(child)
        if getattr(rule, name) is not None:
            dxf.append(writer(getattr(rule, name)))
    if len(dxf):
        dxf_order = ("font", "numFmt", "fill", "alignment", "border", "protection", "extLst")
        dxf[:] = sorted(dxf, key=lambda n: dxf_order.index(local(n)) if local(n) in dxf_order else len(dxf_order))
        node.set("dxfId", str(styles.add("dxfs", dxf)))
    for child in node.findall(q("formula")) + node.findall(q("colorScale")):
        node.remove(child)
    if rule.colorScale is not None:
        scale = ET.SubElement(node, q("colorScale"))
        for type_, value, color in rule.colorScale:
            stop = ET.SubElement(scale, q("cfvo"), {"type": type_})
            if value is not None:
                stop.set("val", scalar(value))
        for type_, value, color in rule.colorScale:
            scale.append(_color_xml("color", color))
    for formula in rule.formula:
        ET.SubElement(node, q("formula")).text = str(formula)
    return node


_CORE_FIELDS = {"creator": (DC, "creator"), "title": (DC, "title"), "subject": (DC, "subject"),
                "description": (DC, "description"), "identifier": (DC, "identifier"), "language": (DC, "language"),
                "created": (DCT, "created"), "modified": (DCT, "modified")}
for _key in ("keywords", "lastModifiedBy", "revision", "category", "contentStatus", "version", "lastPrinted"):
    _CORE_FIELDS[_key] = (CP, _key)
_APP_FIELDS = {"application": "Application", "appVersion": "AppVersion", "company": "Company", "manager": "Manager"}


def _read_properties(archive, book):
    _, rels = _relationships(archive, "")
    for kind, default, mapping, ns in (("core-properties", "docProps/core.xml", _CORE_FIELDS, CP),
                                      ("extended-properties", "docProps/app.xml", _APP_FIELDS, EP)):
        path = _target("", _find_rel(rels, kind)) or default
        root = archive.root(path)
        if root is None:
            continue
        setattr(book, "_" + kind.replace("-", "_"), deepcopy(root))
        for key, tag in mapping.items():
            child = root.find(q(tag[1], tag[0]) if isinstance(tag, tuple) else q(tag, ns))
            if child is not None:
                value = child.text
                if key in ("created", "modified", "lastPrinted") and value:
                    try:
                        value = datetime.fromisoformat(value.replace("Z", "+00:00"))
                    except ValueError:
                        pass
                setattr(book.properties, key, value)


def _read_table(root):
    table = Table(root.get("displayName") or root.get("name"), root.get("ref"))
    for name in ("name", "id", "headerRowCount", "totalsRowCount", "totalsRowShown"):
        if root.get(name) is not None:
            value = root.get(name)
            setattr(table, name, boolean(value) if name == "totalsRowShown" else int(value) if name in ("id", "headerRowCount", "totalsRowCount") else value)
    style = root.find(q("tableStyleInfo"))
    if style is not None:
        table.tableStyleInfo = _model(TableStyleInfo, style)
    table.tableColumns = [_remember(TableColumn(int(n.get("id")), n.get("name", "")), n)
                          for n in root.findall(q("tableColumns") + "/" + q("tableColumn"))]
    return _remember(table, root)


class _OpaqueChart:
    """Loaded unsupported chart; its XML/dependencies are preserved verbatim."""
    _part: str
    _drawing_part: str
    _anchor_xml: ET.Element

    def __init__(self, anchor):
        self.anchor = anchor


def _chart_text(node):
    if node is None:
        return None
    return "".join(n.text or "" for n in node.iter(q("t", A))) or None


def _read_chart(root, anchor):
    bar = root.find(".//" + q("barChart", C))
    plot = root.find(".//" + q("plotArea", C))
    chart_types = [n for n in plot if local(n).endswith("Chart")] if plot is not None else []
    if bar is None or len(chart_types) != 1:
        return _OpaqueChart(anchor)
    chart = BarChart(anchor=anchor, title=_chart_text(root.find(".//" + q("title", C))))
    chart._chart_xml = deepcopy(root)
    for tag, name, convert in (("barDir", "type", str), ("grouping", "grouping", str), ("overlap", "overlap", int), ("gapWidth", "gapWidth", int)):
        child = bar.find(q(tag, C))
        if child is not None:
            setattr(chart, name, convert(child.get("val")))
    style = root.find(q("style", C))
    if style is not None:
        chart.style = int(style.get("val"))
    for node in bar.findall(q("ser", C)):
        values = node.find(q("val", C) + "/" + q("numRef", C) + "/" + q("f", C))
        if values is None:
            # Literal data and uncommon series layouts must remain opaque.
            return _OpaqueChart(anchor)
        title_ref = node.find(q("tx", C) + "/" + q("strRef", C) + "/" + q("f", C))
        title = node.find(q("tx", C) + "/" + q("v", C))
        category = node.find(q("cat", C) + "/" + q("strRef", C) + "/" + q("f", C))
        if category is None:
            category = node.find(q("cat", C) + "/" + q("numRef", C) + "/" + q("f", C))
        series = Series(values.text or "", title=title.text if title is not None else None,
                        title_reference=title_ref.text if title_ref is not None else None,
                        categories=category.text if category is not None else None)
        series._xml = deepcopy(node)
        chart.series.append(series)
    axes = root.findall(".//" + q("catAx", C)) + root.findall(".//" + q("valAx", C))
    if axes:
        chart.x_axis.title = _chart_text(axes[0].find(q("title", C)))
    if len(axes) > 1:
        chart.y_axis.title = _chart_text(axes[1].find(q("title", C)))
    return chart


def _capture_dependencies(archive, book, part, seen=None):
    seen = set() if seen is None else seen
    if part in seen or part not in archive.names:
        return
    seen.add(part)
    book._parts[part] = archive.read(part)
    relpath = _rels_path(part)
    if relpath in archive.names:
        book._parts[relpath] = archive.read(relpath)
        _, rels = _relationships(archive, part)
        for rel in rels.values():
            path = _target(part, rel)
            if path:
                _capture_dependencies(archive, book, path, seen)


def load_metadata(source):
    """Read editable metadata without decoding or retaining native cell values."""
    archive = _Archive(source)
    try:
        book = MetadataBook()
        book._parts = {}
        book._content_types = archive.root("[Content_Types].xml")
        _book_part, workbook, sheets, style_part = _book_parts(archive)
        styles = _Styles(archive.root(style_part) if style_part else None)
        _read_properties(archive, book)
        names = workbook.find(q("definedNames"))
        if names is not None:
            for node in names:
                name = DefinedName(node.get("name"), node.text)
                for f in fields(name):
                    if f.name not in ("name", "attr_text") and node.get(f.name) is not None:
                        value = node.get(f.name)
                        annotation = str(f.type)
                        setattr(name, f.name, boolean(value) if "bool" in annotation else int(value) if "int" in annotation else value)
                book.defined_names.add(_remember(name, node))
        for sheet_node, part in sheets:
            if part not in archive.names:
                continue
            root = _worksheet_metadata_xml(archive.read(part))
            sheet = MetadataSheet(title=sheet_node.get("name"), sheet_state=sheet_node.get("state", "visible"))
            book.sheets[sheet.title] = sheet
            # Never copy sheetData, even its reduced styled-cell/row stubs.
            sheet._xml = ET.Element(root.tag, dict(root.attrib))
            sheet._xml.text, sheet._xml.tail = root.text, root.tail
            sheet._xml.extend(deepcopy(node) for node in root if node.tag != q("sheetData"))
            sheet._part = part
            sheet._rels, rels = _relationships(archive, part)
            for cell in root.findall(q("sheetData") + "/" + q("row") + "/" + q("c")):
                if "s" in cell.attrib:
                    sheet.cells[cell.get("r")] = styles.read(int(cell.get("s")))
            for merged in root.findall(q("mergeCells") + "/" + q("mergeCell")):
                sheet.merged_cells.add(merged.get("ref"))
            for col in root.findall(q("cols") + "/" + q("col")):
                dimension = _model(ColumnDimension, col)
                sheet.column_dimensions[column_letter(int(col.get("min")))] = dimension
            for row in root.findall(q("sheetData") + "/" + q("row")):
                if any(key in row.attrib for key in ("ht", "hidden", "outlineLevel", "collapsed")):
                    dimension = _model(RowDimension, ET.Element(row.tag, dict(row.attrib)), {"height": "ht"})
                    sheet.row_dimensions[int(row.get("r"))] = dimension
            pane = root.find(q("sheetViews") + "/" + q("sheetView") + "/" + q("pane"))
            if pane is not None and pane.get("state") in ("frozen", "frozenSplit"):
                sheet.freeze_panes = pane.get("topLeftCell") or "%s%s" % (column_letter(int(float(pane.get("xSplit", "0"))) + 1), int(float(pane.get("ySplit", "0"))) + 1)
            filter_node = root.find(q("autoFilter"))
            if filter_node is not None:
                sheet.auto_filter = _remember(AutoFilter(filter_node.get("ref")), filter_node)
            props = root.find(q("sheetPr"))
            if props is not None:
                sheet.sheet_properties._xml = deepcopy(props)
                sheet.sheet_properties.tabColor = _read_color(props.find(q("tabColor")))
            for tag, cls, name in (("sheetProtection", SheetProtection, "protection"), ("pageSetup", PageSetup, "page_setup")):
                node = root.find(q(tag))
                if node is not None:
                    setattr(sheet, name, _model(cls, node))
            validations = root.find(q("dataValidations"))
            if validations is not None:
                sheet.data_validations._xml = deepcopy(validations)
                for node in validations.findall(q("dataValidation")):
                    validation = _model(DataValidation, node, exclude=("formula1", "formula2"))
                    for key in ("formula1", "formula2"):
                        formula = node.find(q(key))
                        if formula is not None:
                            setattr(validation, key, formula.text)
                    sheet.data_validations.add(validation)
            for conditional in root.findall(q("conditionalFormatting")):
                for rule in conditional.findall(q("cfRule")):
                    sheet.conditional_formatting.add(conditional.get("sqref"), _read_rule(rule, styles))
            for hyperlink in root.findall(q("hyperlinks") + "/" + q("hyperlink")):
                link = _model(Hyperlink, hyperlink)
                rel = rels.get(hyperlink.get(q("id", R)))
                if rel is not None:
                    link.target = rel.get("Target")
                c1, r1, _, _ = range_bounds(hyperlink.get("ref"))
                sheet.cells.setdefault(column_letter(c1) + str(r1), CellMetadata()).hyperlink = link
            comment_rel = _find_rel(rels, "comments")
            comment_path = _target(part, comment_rel)
            if comment_path:
                comments = archive.root(comment_path)
                if comments is not None:
                    authors = [n.text or "" for n in comments.findall(q("authors") + "/" + q("author"))]
                    for node in comments.findall(q("commentList") + "/" + q("comment")):
                        author_id = int(node.get("authorId", "0"))
                        text = "".join(n.text or "" for n in node.iter(q("t")))
                        comment = _remember(Comment(text, authors[author_id] if author_id < len(authors) else ""), node)
                        comment._original_text = text
                        sheet.cells.setdefault(node.get("ref"), CellMetadata()).comment = comment
            vml_path = _target(part, _find_rel(rels, "vmlDrawing"))
            if vml_path:
                vml = archive.root(vml_path)
                for shape in vml if vml is not None else ():
                    client = shape.find(q("ClientData", X))
                    if client is None or client.get("ObjectType") != "Note":
                        continue
                    row = int(client.findtext(q("Row", X), "0")) + 1
                    col = int(client.findtext(q("Column", X), "0")) + 1
                    cell = sheet.cells.get(column_letter(col) + str(row))
                    if cell is None or cell.comment is None:
                        continue
                    for name in ("width", "height"):
                        match = re.search(r"(?:^|;)\s*" + name + r"\s*:\s*([0-9.]+)(pt|px)?", shape.get("style", ""))
                        if match:
                            setattr(cell.comment, name, float(match[1]) * (4 / 3 if match[2] == "pt" else 1))
            for node in root.findall(q("tableParts") + "/" + q("tablePart")):
                table_path = _target(part, rels.get(node.get(q("id", R))))
                if table_path:
                    table_root = archive.root(table_path)
                    if table_root is not None:
                        table = _read_table(table_root)
                        table._part = table_path
                        sheet.tables.add(table)
                        _capture_dependencies(archive, book, table_path)
            for drawing_node in root.findall(q("drawing")):
                drawing_path = _target(part, rels.get(drawing_node.get(q("id", R))))
                if not drawing_path:
                    continue
                _capture_dependencies(archive, book, drawing_path)
                drawing = archive.root(drawing_path)
                _, drawing_rels = _relationships(archive, drawing_path)
                if drawing is None:
                    continue
                for anchor_node in drawing:
                    chart_node = anchor_node.find(".//" + q("chart", C))
                    if chart_node is None:
                        continue
                    chart_path = _target(drawing_path, drawing_rels.get(chart_node.get(q("id", R))))
                    if not chart_path:
                        continue
                    chart_root = archive.root(chart_path)
                    if chart_root is None:
                        continue
                    origin = anchor_node.find(q("from", XDR))
                    anchor = "A1"
                    if origin is not None:
                        anchor = column_letter(int(origin.findtext(q("col", XDR), "0")) + 1) + str(int(origin.findtext(q("row", XDR), "0")) + 1)
                    chart = _read_chart(chart_root, anchor)
                    chart._part, chart._drawing_part = chart_path, drawing_path
                    chart._anchor_xml = deepcopy(anchor_node)
                    extent = anchor_node.find(q("ext", XDR))
                    if extent is not None and isinstance(chart, BarChart):
                        chart.width, chart.height = int(extent.get("cx")) / 360000, int(extent.get("cy")) / 360000
                    sheet.charts.append(chart)
            # Keep VML, images, printer settings, and other internal dependencies.
            for rel in rels.values():
                path = _target(part, rel)
                if path and not rel.get("Type", "").endswith(("/hyperlink", "/comments")):
                    _capture_dependencies(archive, book, path)
        return book
    finally:
        archive.close()


class _Package:
    def __init__(self, archive, metadata):
        self.archive = archive
        self.parts = {}
        self.used = set(archive.names) | set(getattr(metadata, "_parts", {}))
        self.content_types = archive.root("[Content_Types].xml", ET.Element(q("Types", CT)))
        self.original_types = getattr(metadata, "_content_types", None)
        self.original_parts = getattr(metadata, "_parts", {})
        self.imports = {}

    def allocate(self, directory, stem, extension):
        index = 1
        while "%s/%s%d.%s" % (directory, stem, index, extension) in self.used:
            index += 1
        path = "%s/%s%d.%s" % (directory, stem, index, extension)
        self.used.add(path)
        return path

    def put(self, path, root, content_type=None):
        self.parts[path] = dump(root) if isinstance(root, ET.Element) else root
        self.used.add(path)
        if content_type:
            self.override(path, content_type)

    def override(self, path, type_):
        for node in self.content_types:
            if node.tag == q("Override", CT) and node.get("PartName") == "/" + path:
                node.set("ContentType", type_)
                return
        ET.SubElement(self.content_types, q("Override", CT), {"PartName": "/" + path, "ContentType": type_})

    def default(self, extension, type_):
        if not any(n.tag == q("Default", CT) and n.get("Extension") == extension for n in self.content_types):
            ET.SubElement(self.content_types, q("Default", CT), {"Extension": extension, "ContentType": type_})

    def import_part(self, path):
        if path in self.imports:
            return self.imports[path]
        if path not in self.original_parts:
            return path
        directory, filename = posixpath.split(path)
        stem, extension = posixpath.splitext(filename)
        # Always allocate imported parts, avoiding collisions with native assets.
        new = self.allocate(directory, re.sub(r"\d+$", "", stem) + "Metadata", extension.lstrip("."))
        self.imports[path] = new
        self.parts[new] = self.original_parts[path]
        if self.original_types is not None:
            for node in self.original_types:
                if node.get("PartName") == "/" + path:
                    self.override(new, node.get("ContentType"))
                elif node.tag == q("Default", CT) and node.get("Extension") == extension.lstrip("."):
                    self.default(node.get("Extension"), node.get("ContentType"))
        relpath = _rels_path(path)
        if relpath in self.original_parts:
            rels = xml(self.original_parts[relpath])
            for rel in rels:
                target = _target(path, rel)
                if target:
                    rel.set("Target", posixpath.relpath(self.import_part(target), posixpath.dirname(new)))
            self.put(_rels_path(new), rels)
        return new


def _add_rel(root, type_, target, external=False):
    used = {n.get("Id") for n in root}
    index = 1
    while "rId%d" % index in used:
        index += 1
    id_ = "rId%d" % index
    attributes = {"Id": id_, "Type": R + "/" + type_, "Target": target}
    if external:
        attributes["TargetMode"] = "External"
    ET.SubElement(root, q("Relationship", P), attributes)
    return id_


def _remove_rels(root, suffixes):
    for node in list(root):
        if any(node.get("Type", "").endswith("/" + suffix) for suffix in suffixes):
            root.remove(node)


def _replace(root, tag, nodes):
    for node in list(root):
        if node.tag == q(tag):
            root.remove(node)
    root.extend(nodes)


def _template(sheet, root, tag):
    original = getattr(sheet, "_xml", None)
    node = original.find(q(tag)) if original is not None else None
    if node is None:
        node = root.find(q(tag))
    return deepcopy(node) if node is not None else ET.Element(q(tag))


_SHEET_ORDER = ("sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
                "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
                "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
                "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
                "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
                "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst")


def _sheet_basic(root, sheet, styles):
    data = root.find(q("sheetData"))
    if data is None:
        data = ET.SubElement(root, q("sheetData"))
    rows = {int(n.get("r")): n for n in data.findall(q("row"))}
    cells = {n.get("r"): n for row in rows.values() for n in row.findall(q("c"))}
    for coordinate, metadata in sheet.cells.items():
        col, row_index = coordinate_parts(coordinate)
        coordinate = column_letter(col) + str(row_index)
        if coordinate not in cells:
            row = rows.get(row_index)
            if row is None:
                row = rows[row_index] = ET.SubElement(data, q("row"), {"r": str(row_index)})
            cells[coordinate] = ET.SubElement(row, q("c"), {"r": coordinate})
        styles.overlay(cells[coordinate], metadata)
    for index, dimension in sheet.row_dimensions.items():
        if not isinstance(index, int) or not 1 <= index <= 1048576:
            raise ValueError("Invalid row dimension index")
        row = rows.get(index)
        if row is None:
            row = rows[index] = ET.SubElement(data, q("row"), {"r": str(index)})
        for key in ("ht", "customHeight", "hidden", "outlineLevel", "collapsed"):
            row.attrib.pop(key, None)
        row.attrib.update(attrs(dimension, rename={"height": "ht"}))
        if dimension.height is not None:
            row.set("customHeight", "1")
    data[:] = sorted(data, key=lambda n: int(n.get("r", "0")))
    for row in data:
        row[:] = sorted(row, key=lambda n: coordinate_parts(n.get("r"))[0] if local(n) == "c" else 16385)
    # Expand dimension for metadata-only cells; do not contract the native range.
    if cells:
        coordinates = [coordinate_parts(c) for c in cells]
        existing = root.find(q("dimension"))
        if existing is not None:
            try:
                c1, r1, c2, r2 = range_bounds(existing.get("ref"))
                coordinates.extend(((c1, r1), (c2, r2)))
            except (ValueError, TypeError):
                pass
        c1, r1 = min(c for c, r in coordinates), min(r for c, r in coordinates)
        c2, r2 = max(c for c, r in coordinates), max(r for c, r in coordinates)
        _replace(root, "dimension", [ET.Element(q("dimension"), {"ref": "%s%d:%s%d" % (column_letter(c1), r1, column_letter(c2), r2)})])
    cols = ET.Element(q("cols"))
    native_styles, resolved = {}, {}
    for native in root.findall(q("cols") + "/" + q("col")):
        low, high = int(native.get("min", "1")), int(native.get("max", "1"))
        column_letter(low)
        column_letter(high)
        if native.get("style") is not None:
            for index in range(low, high + 1):
                native_styles[index] = native.get("style")
                resolved[index] = ET.Element(q("col"), {"style": native.get("style")})
    # Resolve overlapping grouped columns so an explicit B dimension overrides
    # a loaded A:C group. Native style IDs survive even on unmodeled columns.
    for letter, dimension in sorted(sheet.column_dimensions.items(), key=lambda item: column_index(item[0])):
        index = column_index(letter)
        low, high = dimension.min or index, dimension.max or index
        column_letter(low)
        column_letter(high)
        if high < low:
            raise ValueError("Reversed column dimension")
        node = deepcopy(getattr(dimension, "_xml", ET.Element(q("col"))))
        for key in ("min", "max", "style", "width", "hidden", "bestFit", "outlineLevel", "collapsed", "customWidth"):
            node.attrib.pop(key, None)
        node.attrib.update(attrs(dimension, exclude=("min", "max")))
        if dimension.width is not None:
            node.set("customWidth", "1")
        for index in range(low, high + 1):
            child = deepcopy(node)
            if index in native_styles:
                child.set("style", native_styles[index])
            resolved[index] = child
    previous, previous_key, previous_index = None, None, 0
    for index, node in sorted(resolved.items()):
        key = ET.tostring(node)
        if previous is not None and index == previous_index + 1 and key == previous_key:
            previous.set("max", str(index))
        else:
            node.set("min", str(index))
            node.set("max", str(index))
            cols.append(node)
            previous = node
        previous_index, previous_key = index, key
    _replace(root, "cols", [cols] if len(cols) else [])
    merged = ET.Element(q("mergeCells"), {"count": str(len(sheet.merged_cells.ranges))})
    for value in sheet.merged_cells.ranges:
        range_bounds(str(value))
        ET.SubElement(merged, q("mergeCell"), {"ref": str(value)})
    _replace(root, "mergeCells", [merged] if len(merged) else [])
    views = _template(sheet, root, "sheetViews")
    if not len(views):
        ET.SubElement(views, q("sheetView"), {"workbookViewId": "0"})
    view = views.find(q("sheetView"))
    if view is None:
        view = ET.SubElement(views, q("sheetView"), {"workbookViewId": "0"})
    old_pane = view.find(q("pane"))
    was_frozen = old_pane is not None and old_pane.get("state") in ("frozen", "frozenSplit")
    if was_frozen or sheet.freeze_panes:
        for node in view.findall(q("pane")) + view.findall(q("selection")):
            view.remove(node)
    if sheet.freeze_panes and sheet.freeze_panes != "A1":
        x, y = coordinate_parts(getattr(sheet.freeze_panes, "coordinate", sheet.freeze_panes))
        pane = ET.Element(q("pane"), {"state": "frozen", "topLeftCell": column_letter(x) + str(y)})
        if x > 1:
            pane.set("xSplit", str(x - 1))
        if y > 1:
            pane.set("ySplit", str(y - 1))
        active = "bottomRight" if x > 1 and y > 1 else "topRight" if x > 1 else "bottomLeft"
        pane.set("activePane", active)
        view.insert(0, pane)
        ET.SubElement(view, q("selection"), {"pane": active, "activeCell": column_letter(x) + str(y), "sqref": column_letter(x) + str(y)})
    _replace(root, "sheetViews", [views])
    filter_node = deepcopy(getattr(sheet.auto_filter, "_xml", ET.Element(q("autoFilter"))))
    if sheet.auto_filter.ref:
        range_bounds(sheet.auto_filter.ref)
        filter_node.set("ref", sheet.auto_filter.ref)
    _replace(root, "autoFilter", [filter_node] if sheet.auto_filter.ref else [])
    props = _template(sheet, root, "sheetPr")
    for node in props.findall(q("tabColor")):
        props.remove(node)
    if sheet.sheet_properties.tabColor is not None:
        props.insert(0, _color_xml("tabColor", sheet.sheet_properties.tabColor))
    if sheet.page_setup.fitToWidth is not None or sheet.page_setup.fitToHeight is not None:
        setup_props = props.find(q("pageSetUpPr"))
        if setup_props is None:
            setup_props = ET.SubElement(props, q("pageSetUpPr"))
        setup_props.set("fitToPage", "1")
    _replace(root, "sheetPr", [props] if len(props) or props.attrib else [])
    for tag, model in (("sheetProtection", sheet.protection), ("pageSetup", sheet.page_setup)):
        node = _model_xml(tag, model)
        _replace(root, tag, [node] if node.attrib or len(node) else [])
    validations = deepcopy(getattr(sheet.data_validations, "_xml", ET.Element(q("dataValidations"))))
    validations[:] = [n for n in validations if local(n) != "dataValidation"]
    count = 0
    for validation in sheet.data_validations:
        if not str(validation.sqref).strip():
            continue
        for value in str(validation.sqref).split():
            range_bounds(value)
        node = _model_xml("dataValidation", validation, exclude=("formula1", "formula2"))
        for key in ("formula1", "formula2"):
            for old in node.findall(q(key)):
                node.remove(old)
            if getattr(validation, key) is not None:
                ET.SubElement(node, q(key)).text = str(getattr(validation, key))
        validations.append(node)
        count += 1
    validations.set("count", str(count))
    _replace(root, "dataValidations", [validations] if count else [])
    conditionals = []
    priority = 1
    for cell_range in sheet.conditional_formatting:
        node = ET.Element(q("conditionalFormatting"), {"sqref": str(cell_range)})
        for rule in sheet.conditional_formatting[cell_range]:
            node.append(_rule_xml(rule, styles, priority))
            priority += 1
        conditionals.append(node)
    _replace(root, "conditionalFormatting", conditionals)
    return cells


def _hyperlinks(root, rels, sheet):
    _remove_rels(rels, ("hyperlink",))
    hyperlinks = ET.Element(q("hyperlinks"))
    for coordinate, metadata in sheet.cells.items():
        value = metadata.hyperlink
        if value is None:
            continue
        link = Hyperlink(target=value) if isinstance(value, str) else value
        node = deepcopy(getattr(link, "_xml", ET.Element(q("hyperlink"))))
        for key in ("target", "id", q("id", R), "location", "tooltip", "display"):
            node.attrib.pop(key, None)
        range_bounds(link.ref or coordinate)
        node.set("ref", link.ref or coordinate)
        for key in ("location", "tooltip", "display"):
            if getattr(link, key) is not None:
                node.set(key, str(getattr(link, key)))
        if link.target:
            if link.target.startswith("#"):
                node.set("location", link.target[1:])
            else:
                node.set(q("id", R), _add_rel(rels, "hyperlink", link.target, external=True))
        hyperlinks.append(node)
    _replace(root, "hyperlinks", [hyperlinks] if len(hyperlinks) else [])


def _comments(package, root, rels, part, sheet):
    _remove_rels(rels, ("comments",))
    comments = [(coord, cell.comment) for coord, cell in sheet.cells.items() if cell.comment is not None]
    # Preserve non-note VML (buttons and other controls) through import below.
    legacy = _template(sheet, root, "legacyDrawing")
    legacy_id = legacy.get(q("id", R))
    original_rels = {n.get("Id"): n for n in getattr(sheet, "_rels", ())}
    legacy_rel = original_rels.get(legacy_id)
    vml = None
    native_path = None
    vml_path = _target(getattr(sheet, "_part", part), legacy_rel)
    if vml_path in package.original_parts:
        vml = xml(package.original_parts[vml_path])
    if vml is None:
        # New metadata can also be overlaid on a native archive with VML.
        native_rel = next((n for n in rels if n.get("Id") == legacy_id), None)
        native_path = _target(part, native_rel)
        if native_path and native_path in package.archive.names:
            vml = package.archive.root(native_path)
    if vml is not None:
        for shape in list(vml):
            if any(n.get("ObjectType") == "Note" for n in shape.iter(q("ClientData", X))):
                vml.remove(shape)
    if not comments:
        if vml is not None:
            if any(local(n) == "shape" for n in vml):
                out_vml = package.import_part(vml_path) if vml_path in package.original_parts else native_path
                if out_vml:
                    package.put(out_vml, vml)
            else:
                current = root.find(q("legacyDrawing"))
                if current is not None:
                    current_id = current.get(q("id", R))
                    rels[:] = [n for n in rels if n.get("Id") != current_id]
                _replace(root, "legacyDrawing", [])
        return
    if vml is None:
        vml = ET.Element("xml")
    if not any(n.get("id") == "_x0000_t202" for n in vml):
        shape_type = ET.SubElement(vml, q("shapetype", V), {"id": "_x0000_t202", "coordsize": "21600,21600", q("spt", O): "202", "path": "m,l,21600r21600,l21600,xe"})
        ET.SubElement(shape_type, q("stroke", V), {"joinstyle": "miter"})
        ET.SubElement(shape_type, q("path", V), {"gradientshapeok": "t", q("connecttype", O): "rect"})
    comments_root = ET.Element(q("comments"))
    authors_node, list_node = ET.SubElement(comments_root, q("authors")), ET.SubElement(comments_root, q("commentList"))
    authors = []
    used_shape_ids = {n.get("id") for n in vml}
    shape_index = 1025
    for coordinate, comment in comments:
        col, row = coordinate_parts(coordinate)
        if comment.author not in authors:
            authors.append(comment.author)
            ET.SubElement(authors_node, q("author")).text = comment.author
        node = deepcopy(getattr(comment, "_xml", ET.Element(q("comment"))))
        node.set("ref", coordinate)
        node.set("authorId", str(authors.index(comment.author)))
        if comment.text != getattr(comment, "_original_text", None):
            for old in node.findall(q("text")):
                node.remove(old)
            text = ET.SubElement(node, q("text"))
            ET.SubElement(text, q("t"), {q("space", "http://www.w3.org/XML/1998/namespace"): "preserve"}).text = comment.text
        list_node.append(node)
        while "_x0000_s%d" % shape_index in used_shape_ids:
            shape_index += 1
        shape = ET.SubElement(vml, q("shape", V), {"id": "_x0000_s%d" % shape_index, "type": "#_x0000_t202",
                              "style": "position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:%spt;height:%spt;z-index:1;visibility:hidden" % (comment.width * 0.75, comment.height * 0.75),
                              "fillcolor": "#ffffe1", q("insetmode", O): "auto"})
        shape_index += 1
        ET.SubElement(shape, q("fill", V), {"color2": "#ffffe1"})
        ET.SubElement(shape, q("shadow", V), {"on": "t", "color": "black", "obscured": "t"})
        ET.SubElement(shape, q("path", V), {q("connecttype", O): "none"})
        ET.SubElement(ET.SubElement(shape, q("textbox", V), {"style": "mso-direction-alt:auto"}), "div", {"style": "text-align:left"})
        client = ET.SubElement(shape, q("ClientData", X), {"ObjectType": "Note"})
        ET.SubElement(client, q("MoveWithCells", X))
        ET.SubElement(client, q("SizeWithCells", X))
        ET.SubElement(client, q("AutoFill", X)).text = "False"
        ET.SubElement(client, q("Row", X)).text = str(row - 1)
        ET.SubElement(client, q("Column", X)).text = str(col - 1)
    comment_part = package.allocate("xl/comments", "comment", "xml")
    package.put(comment_part, comments_root, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
    _add_rel(rels, "comments", posixpath.relpath(comment_part, posixpath.dirname(part)))
    # Use imported VML to retain control dependencies, if present.
    if vml_path in package.original_parts:
        out_vml = package.import_part(vml_path)
    else:
        out_vml = package.allocate("xl/drawings", "commentsDrawing", "vml")
    package.put(out_vml, vml)
    package.default("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")
    current_legacy = root.find(q("legacyDrawing"))
    if current_legacy is not None:
        current_id = current_legacy.get(q("id", R))
        rels[:] = [n for n in rels if n.get("Id") != current_id]
    id_ = _add_rel(rels, "vmlDrawing", posixpath.relpath(out_vml, posixpath.dirname(part)))
    _replace(root, "legacyDrawing", [ET.Element(q("legacyDrawing"), {q("id", R): id_})])


def _table_xml(table, id_, cells, archive):
    node = _model_xml("table", table, exclude=("tableStyleInfo", "tableColumns"))
    node.set("id", str(id_))
    c1, r1, c2, r2 = range_bounds(table.ref)
    columns = list(table.tableColumns)
    if columns and len(columns) != c2 - c1 + 1:
        raise ValueError("Table column count does not match its range")
    if not columns:
        shared = None
        columns = []
        for index, col in enumerate(range(c1, c2 + 1), 1):
            cell = cells.get(column_letter(col) + str(r1))
            name = None
            if cell is not None and table.headerRowCount:
                if cell.get("t") == "inlineStr":
                    name = "".join(n.text or "" for n in cell.iter(q("t")))
                elif cell.get("t") == "s":
                    if shared is None:
                        # Shared-string path follows workbook relationships, not an assumed filename.
                        workbook_part, _, _, _ = _book_parts(archive)
                        _, rels = _relationships(archive, workbook_part)
                        path = _target(workbook_part, _find_rel(rels, "sharedStrings"))
                        shared_root = archive.root(path) if path else None
                        shared = ["".join(t.text or "" for t in n.iter(q("t"))) for n in shared_root] if shared_root is not None else []
                    value = cell.findtext(q("v"))
                    if value is not None and 0 <= int(value) < len(shared):
                        name = shared[int(value)]
                else:
                    name = cell.findtext(q("v"))
            columns.append(TableColumn(index, str(name) if name is not None else "Column%d" % index))
    names = [c.name.casefold() for c in columns]
    if len(set(names)) != len(names):
        raise ValueError("Table column names must be unique")
    col_node = ET.Element(q("tableColumns"), {"count": str(len(columns))})
    for index, col in enumerate(columns, 1):
        child = deepcopy(getattr(col, "_xml", ET.Element(q("tableColumn"))))
        child.set("id", str(index))
        child.set("name", col.name)
        col_node.append(child)
    for old in node.findall(q("tableColumns")) + node.findall(q("tableStyleInfo")):
        node.remove(old)
    filter_node = node.find(q("autoFilter"))
    if table.headerRowCount:
        if filter_node is None:
            filter_node = ET.SubElement(node, q("autoFilter"))
        bottom = r2 - (table.totalsRowCount or (1 if table.totalsRowShown else 0))
        if bottom < r1:
            raise ValueError("Table has no header/data rows")
        filter_node.set("ref", "%s%d:%s%d" % (column_letter(c1), r1, column_letter(c2), bottom))
    elif filter_node is not None:
        node.remove(filter_node)
    node.append(col_node)
    if table.tableStyleInfo is not None:
        node.append(_model_xml("tableStyleInfo", table.tableStyleInfo))
    order = ("autoFilter", "sortState", "tableColumns", "tableStyleInfo", "extLst")
    node[:] = sorted(node, key=lambda n: order.index(local(n)) if local(n) in order else len(order))
    return node


def _tables(package, root, rels, part, sheet, cells, used_names, used_ids):
    _remove_rels(rels, ("table",))
    parts = ET.Element(q("tableParts"))
    for table in sheet.tables.values():
        name = table.displayName
        if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", name) or name.casefold() in used_names:
            raise ValueError("Invalid or duplicate table name: " + name)
        if re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]*", name):
            try:
                coordinate_parts(name)
            except ValueError:
                pass
            else:
                raise ValueError("Table name cannot be a cell reference")
        used_names.add(name.casefold())
        id_ = table.id or 1
        while id_ in used_ids:
            id_ += 1
        used_ids.add(id_)
        original = getattr(table, "_part", None)
        table_part = package.import_part(original) if original in package.original_parts else package.allocate("xl/tables", "table", "xml")
        package.put(table_part, _table_xml(table, id_, cells, package.archive), "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
        id_ = _add_rel(rels, "table", posixpath.relpath(table_part, posixpath.dirname(part)))
        ET.SubElement(parts, q("tablePart"), {q("id", R): id_})
    parts.set("count", str(len(parts)))
    _replace(root, "tableParts", [parts] if len(parts) else [])


def _chart_title(value):
    title = ET.Element(q("title", C))
    rich = ET.SubElement(ET.SubElement(title, q("tx", C)), q("rich", C))
    ET.SubElement(rich, q("bodyPr", A))
    ET.SubElement(rich, q("lstStyle", A))
    paragraph = ET.SubElement(rich, q("p", A))
    ET.SubElement(ET.SubElement(paragraph, q("r", A)), q("t", A)).text = str(value)
    ET.SubElement(title, q("overlay", C), {"val": "0"})
    return title


def _chart_xml(chart):
    root = deepcopy(getattr(chart, "_chart_xml", ET.Element(q("chartSpace", C))))
    body = root.find(q("chart", C))
    if body is None:
        body = ET.SubElement(root, q("chart", C))
    for old in root.findall(q("style", C)):
        root.remove(old)
    if chart.style is not None:
        root.insert(0, ET.Element(q("style", C), {"val": str(chart.style)}))
    for old in body.findall(q("title", C)):
        body.remove(old)
    if chart.title is not None:
        body.insert(0, _chart_title(chart.title))
    plot = body.find(q("plotArea", C))
    if plot is None:
        plot = ET.SubElement(body, q("plotArea", C))
        ET.SubElement(plot, q("layout", C))
    bar = plot.find(q("barChart", C))
    if bar is None:
        bar = ET.SubElement(plot, q("barChart", C))
    # Preserve series formatting but replace the declared data references.
    for tag in ("barDir", "grouping", "ser", "gapWidth", "overlap"):
        for old in bar.findall(q(tag, C)):
            bar.remove(old)
    bar.insert(0, ET.Element(q("barDir", C), {"val": chart.type}))
    bar.insert(1, ET.Element(q("grouping", C), {"val": chart.grouping}))
    for index, series in enumerate(chart.series):
        node = deepcopy(getattr(series, "_xml", ET.Element(q("ser", C))))
        for tag in ("idx", "order", "tx", "cat", "val"):
            for old in node.findall(q(tag, C)):
                node.remove(old)
        node.insert(0, ET.Element(q("idx", C), {"val": str(index)}))
        node.insert(1, ET.Element(q("order", C), {"val": str(index)}))
        if series.title is not None or series.title_reference is not None:
            tx = ET.Element(q("tx", C))
            if series.title_reference is not None:
                ET.SubElement(ET.SubElement(tx, q("strRef", C)), q("f", C)).text = series.title_reference
            else:
                ET.SubElement(tx, q("v", C)).text = series.title
            node.insert(2, tx)
        if series.categories:
            ET.SubElement(ET.SubElement(ET.SubElement(node, q("cat", C)), q("strRef", C)), q("f", C)).text = series.categories
        ET.SubElement(ET.SubElement(ET.SubElement(node, q("val", C)), q("numRef", C)), q("f", C)).text = series.values
        series_order = ("idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "shape", "extLst")
        node[:] = sorted(node, key=lambda n: series_order.index(local(n)) if local(n) in series_order else len(series_order))
        bar.append(node)
    ET.SubElement(bar, q("gapWidth", C), {"val": str(chart.gapWidth)})
    if chart.overlap is not None:
        ET.SubElement(bar, q("overlap", C), {"val": str(chart.overlap)})
    ids = [n.get("val") for n in bar.findall(q("axId", C))]
    if len(ids) < 2:
        for old in bar.findall(q("axId", C)):
            bar.remove(old)
        ids = ["10", "100"]
        for id_ in ids:
            ET.SubElement(bar, q("axId", C), {"val": id_})
    for axis_tag, id_, cross, position, title in (("catAx", ids[0], ids[1], "l" if chart.type == "bar" else "b", chart.x_axis.title),
                                                ("valAx", ids[1], ids[0], "b" if chart.type == "bar" else "l", chart.y_axis.title)):
        axis = plot.find(q(axis_tag, C))
        if axis is None:
            axis = ET.SubElement(plot, q(axis_tag, C))
            ET.SubElement(axis, q("axId", C), {"val": id_})
            ET.SubElement(ET.SubElement(axis, q("scaling", C)), q("orientation", C), {"val": "minMax"})
            ET.SubElement(axis, q("delete", C), {"val": "0"})
            ET.SubElement(axis, q("axPos", C), {"val": position})
            if axis_tag == "valAx":
                ET.SubElement(axis, q("numFmt", C), {"formatCode": "General", "sourceLinked": "1"})
            ET.SubElement(axis, q("majorTickMark", C), {"val": "none"})
            ET.SubElement(axis, q("minorTickMark", C), {"val": "none"})
            ET.SubElement(axis, q("tickLblPos", C), {"val": "nextTo"})
            ET.SubElement(axis, q("crossAx", C), {"val": cross})
            ET.SubElement(axis, q("crosses", C), {"val": "autoZero"})
        for old in axis.findall(q("title", C)):
            axis.remove(old)
        if title is not None:
            axis.append(_chart_title(title))
        axis_order = ("axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst")
        axis[:] = sorted(axis, key=lambda n: axis_order.index(local(n)) if local(n) in axis_order else len(axis_order))
    bar_order = ("barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "overlap", "serLines", "axId", "extLst")
    bar[:] = sorted(bar, key=lambda n: bar_order.index(local(n)) if local(n) in bar_order else len(bar_order))
    body_order = ("title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall", "backWall", "plotArea", "legend", "plotVisOnly", "dispBlanksAs", "showDLblsOverMax", "extLst")
    body[:] = sorted(body, key=lambda n: body_order.index(local(n)) if local(n) in body_order else len(body_order))
    root_order = ("date1904", "lang", "roundedCorners", "AlternateContent", "style", "clrMapOvr", "pivotSource", "protection", "chart", "spPr", "txPr", "externalData", "printSettings", "userShapes", "extLst")
    root[:] = sorted(root, key=lambda n: root_order.index(local(n)) if local(n) in root_order else len(root_order))
    return root


def _chart_anchor(chart, id_, object_id):
    col, row = coordinate_parts(getattr(chart.anchor, "coordinate", chart.anchor))
    original = getattr(chart, "_anchor_xml", None)
    if original is not None:
        anchor = deepcopy(original)
        origin = anchor.find(q("from", XDR))
        if origin is not None:
            old_col = int(origin.findtext(q("col", XDR), "0"))
            old_row = int(origin.findtext(q("row", XDR), "0"))
            origin.find(q("col", XDR)).text = str(col - 1)
            origin.find(q("row", XDR)).text = str(row - 1)
            end = anchor.find(q("to", XDR))
            if end is not None:
                end.find(q("col", XDR)).text = str(int(end.findtext(q("col", XDR))) + col - 1 - old_col)
                end.find(q("row", XDR)).text = str(int(end.findtext(q("row", XDR))) + row - 1 - old_row)
        extent = anchor.find(q("ext", XDR))
        if extent is not None and isinstance(chart, BarChart):
            extent.set("cx", str(round(chart.width * 360000)))
            extent.set("cy", str(round(chart.height * 360000)))
        anchor.find(".//" + q("chart", C)).set(q("id", R), id_)
        for node in anchor.iter(q("cNvPr", XDR)):
            node.set("id", str(object_id))
        return anchor
    anchor = ET.Element(q("oneCellAnchor", XDR))
    origin = ET.SubElement(anchor, q("from", XDR))
    for name, value in (("col", col - 1), ("colOff", 0), ("row", row - 1), ("rowOff", 0)):
        ET.SubElement(origin, q(name, XDR)).text = str(value)
    ET.SubElement(anchor, q("ext", XDR), {"cx": str(round(chart.width * 360000)), "cy": str(round(chart.height * 360000))})
    frame = ET.SubElement(anchor, q("graphicFrame", XDR), {"macro": ""})
    properties = ET.SubElement(frame, q("nvGraphicFramePr", XDR))
    ET.SubElement(properties, q("cNvPr", XDR), {"id": str(object_id), "name": "Chart %d" % object_id})
    ET.SubElement(properties, q("cNvGraphicFramePr", XDR))
    transform = ET.SubElement(frame, q("xfrm", XDR))
    ET.SubElement(transform, q("off", A), {"x": "0", "y": "0"})
    ET.SubElement(transform, q("ext", A), {"cx": "0", "cy": "0"})
    graphic = ET.SubElement(frame, q("graphic", A))
    graphic_data = ET.SubElement(graphic, q("graphicData", A), {"uri": C})
    ET.SubElement(graphic_data, q("chart", C), {q("id", R): id_})
    ET.SubElement(anchor, q("clientData", XDR))
    return anchor


def _drawings(package, root, rels, part, sheet):
    drawing = ET.Element(q("wsDr", XDR))
    drawing_rels = ET.Element(q("Relationships", P))
    # Retain non-chart shapes/images from original drawings, or native drawings
    # when this is a newly constructed MetadataBook.
    original_rels = {n.get("Id"): n for n in getattr(sheet, "_rels", ())}
    template = getattr(sheet, "_xml", root)
    original_part = getattr(sheet, "_part", part)
    for node in template.findall(q("drawing")):
        rel = original_rels.get(node.get(q("id", R)))
        loaded = rel is not None
        if rel is None:
            rel = next((n for n in rels if n.get("Id") == node.get(q("id", R))), None)
        drawing_part = _target(original_part if loaded else part, rel)
        if not drawing_part:
            continue
        if loaded and drawing_part in package.original_parts:
            source_drawing = xml(package.original_parts[drawing_part])
            source_rel_data = package.original_parts.get(_rels_path(drawing_part))
            source_rels = xml(source_rel_data) if source_rel_data else ET.Element(q("Relationships", P))
        elif drawing_part in package.archive.names:
            source_drawing = package.archive.root(drawing_part)
            source_rels, _ = _relationships(package.archive, drawing_part)
        else:
            continue
        relmap = {}
        for source_rel in source_rels:
            target = _target(drawing_part, source_rel)
            if target:
                target = package.import_part(target) if loaded else target
                # Patched to a relative path after allocating the destination.
            else:
                target = source_rel.get("Target")
            new_id = _add_rel(drawing_rels, source_rel.get("Type").rsplit("/", 1)[-1], target, source_rel.get("TargetMode") == "External")
            relmap[source_rel.get("Id")] = new_id
        for anchor in source_drawing:
            if anchor.find(".//" + q("chart", C)) is not None:
                continue
            anchor = deepcopy(anchor)
            for child in anchor.iter():
                for key, value in list(child.attrib.items()):
                    if key.startswith("{" + R + "}") and value in relmap:
                        child.set(key, relmap[value])
            drawing.append(anchor)
    _remove_rels(rels, ("drawing",))
    _replace(root, "drawing", [])
    if not sheet.charts and not len(drawing):
        return
    drawing_part = package.allocate("xl/drawings", "drawing", "xml")
    for rel in drawing_rels:
        if rel.get("TargetMode") != "External":
            rel.set("Target", posixpath.relpath(rel.get("Target"), posixpath.dirname(drawing_part)))
    object_ids = [int(n.get("id", "0")) for n in drawing.iter(q("cNvPr", XDR))]
    object_id = max(object_ids, default=0) + 1
    for chart in sheet.charts:
        original = getattr(chart, "_part", None)
        if original in package.original_parts:
            chart_part = package.import_part(original)
        else:
            chart_part = package.allocate("xl/charts", "chart", "xml")
        if isinstance(chart, BarChart):
            package.put(chart_part, _chart_xml(chart), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
        elif original not in package.original_parts:
            raise TypeError("Only BarChart or loaded opaque chart objects can be serialized")
        id_ = _add_rel(drawing_rels, "chart", posixpath.relpath(chart_part, posixpath.dirname(drawing_part)))
        drawing.append(_chart_anchor(chart, id_, object_id))
        object_id += 1
    package.put(drawing_part, drawing, "application/vnd.openxmlformats-officedocument.drawing+xml")
    package.put(_rels_path(drawing_part), drawing_rels)
    id_ = _add_rel(rels, "drawing", posixpath.relpath(drawing_part, posixpath.dirname(part)))
    _replace(root, "drawing", [ET.Element(q("drawing"), {q("id", R): id_})])


def _sheet_other_relationships(package, root, rels, part, sheet):
    original_rels = {n.get("Id"): n for n in getattr(sheet, "_rels", ())}
    template = getattr(sheet, "_xml", None)
    if template is None:
        return
    # Restore unmodeled worksheet features, rebasing all relationship IDs.
    modeled = {"sheetData", "dimension", "sheetPr", "sheetViews", "cols", "mergeCells", "autoFilter",
               "sheetProtection", "pageSetup", "dataValidations", "conditionalFormatting", "hyperlinks", "tableParts", "drawing"}
    for original in template:
        tag = local(original)
        if tag in modeled:
            continue
        node = deepcopy(original)
        for child in node.iter():
            for key, value in list(child.attrib.items()):
                if key.startswith("{" + R + "}") and value in original_rels:
                    rel = original_rels[value]
                    path = _target(sheet._part, rel)
                    target = posixpath.relpath(package.import_part(path), posixpath.dirname(part)) if path else rel.get("Target")
                    child.set(key, _add_rel(rels, rel.get("Type").rsplit("/", 1)[-1], target, rel.get("TargetMode") == "External"))
        _replace(root, tag, [node])
    # pageSetup is modeled, but its printer-settings relationship is not.
    for node in root.findall(q("pageSetup")):
        old_id = node.get(q("id", R))
        if old_id in original_rels:
            rel = original_rels[old_id]
            path = _target(sheet._part, rel)
            if path:
                target = posixpath.relpath(package.import_part(path), posixpath.dirname(part))
                node.set(q("id", R), _add_rel(rels, "printerSettings", target))


def _write_properties(package, metadata, root_rels, titles):
    props = metadata.properties
    for kind, path, ns, tag, mapping, private in (("core-properties", "docProps/core.xml", CP, "coreProperties", _CORE_FIELDS, "_core_properties"),
                                               ("extended-properties", "docProps/app.xml", EP, "Properties", _APP_FIELDS, "_extended_properties")):
        existing = next((n for n in root_rels if n.get("Type", "").endswith("/" + kind)), None)
        if existing is not None:
            path = _target("", existing) or path
        root = deepcopy(getattr(metadata, private, None))
        if root is None:
            root = package.archive.root(path, ET.Element(q(tag, ns)))
        for key, field_tag in mapping.items():
            qualified = q(field_tag[1], field_tag[0]) if isinstance(field_tag, tuple) else q(field_tag, ns)
            for old in root.findall(qualified):
                root.remove(old)
            value = getattr(props, key)
            if value is not None:
                node = ET.SubElement(root, qualified)
                if isinstance(value, datetime):
                    if value.tzinfo is None:
                        value = value.replace(tzinfo=timezone.utc)
                    value = value.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
                node.text = str(value)
                if key in ("created", "modified"):
                    node.set(q("type", XSI), "dcterms:W3CDTF")
        if ns == EP:
            for old in root.findall(q("TitlesOfParts", EP)) + root.findall(q("HeadingPairs", EP)):
                root.remove(old)
            vector = ET.SubElement(ET.SubElement(root, q("TitlesOfParts", EP)), q("vector", VT), {"size": str(len(titles)), "baseType": "lpstr"})
            for title in titles:
                ET.SubElement(vector, q("lpstr", VT)).text = title
            vector = ET.SubElement(ET.SubElement(root, q("HeadingPairs", EP)), q("vector", VT), {"size": "2", "baseType": "variant"})
            ET.SubElement(ET.SubElement(vector, q("variant", VT)), q("lpstr", VT)).text = "Worksheets"
            ET.SubElement(ET.SubElement(vector, q("variant", VT)), q("i4", VT)).text = str(len(titles))
        package.put(path, root, "application/vnd.openxmlformats-package.core-properties+xml" if ns == CP else "application/vnd.openxmlformats-officedocument.extended-properties+xml")
        if existing is None:
            id_ = _add_rel(root_rels, kind, path)
            if ns == CP:
                next(n for n in root_rels if n.get("Id") == id_).set("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")


def apply_metadata(source, target, metadata):
    """Overlay metadata on native XLSX bytes, then write target; return None.

    Paths and seekable binary streams are supported (not raw bytes). Sheet
    metadata is matched by exact title. Unknown titles and an all-hidden
    workbook raise ValueError before the destination is opened.
    """
    if not isinstance(metadata, MetadataBook):
        raise TypeError("metadata must be a MetadataBook")
    archive = _Archive(source)
    try:
        book_part, workbook, sheets, style_part = _book_parts(archive)
        titles = [n.get("name") for n, _ in sheets]
        unknown = set(metadata.sheets) - set(titles)
        if unknown:
            raise ValueError("Metadata refers to missing native sheets: " + ", ".join(sorted(unknown)))
        package = _Package(archive, metadata)
        styles = _Styles(archive.root(style_part) if style_part else None)
        book_rels, _ = _relationships(archive, book_part)
        if not style_part:
            style_part = package.allocate(posixpath.dirname(book_part), "styles", "xml")
            _add_rel(book_rels, "styles", posixpath.relpath(style_part, posixpath.dirname(book_part)))
        # Reserve table IDs/names from sheets not being overlaid.
        used_names, used_ids = set(), set()
        for node, part in sheets:
            if node.get("name") in metadata.sheets:
                continue
            _, rels = _relationships(archive, part)
            for rel in rels.values():
                if rel.get("Type", "").endswith("/table"):
                    path = _target(part, rel)
                    table = archive.root(path) if path else None
                    if table is not None:
                        used_names.add(table.get("displayName", "").casefold())
                        used_ids.add(int(table.get("id", "0")))
        for sheet_node, part in sheets:
            sheet = metadata.sheets.get(sheet_node.get("name"))
            if sheet is None:
                continue
            if sheet.sheet_state not in ("visible", "hidden", "veryHidden"):
                raise ValueError("Invalid sheet_state: " + sheet.sheet_state)
            sheet_node.set("state", sheet.sheet_state)
            root = archive.root(part)
            if root is None:
                raise ValueError("Native worksheet is missing: " + part)
            rels, _ = _relationships(archive, part)
            cells = _sheet_basic(root, sheet, styles)
            _sheet_other_relationships(package, root, rels, part, sheet)
            _hyperlinks(root, rels, sheet)
            _comments(package, root, rels, part, sheet)
            _tables(package, root, rels, part, sheet, cells, used_names, used_ids)
            _drawings(package, root, rels, part, sheet)
            root[:] = sorted(root, key=lambda n: _SHEET_ORDER.index(local(n)) if local(n) in _SHEET_ORDER else len(_SHEET_ORDER))
            package.put(part, root)
            if len(rels):
                package.put(_rels_path(part), rels)
        if sheets and not any(node.get("state", "visible") == "visible" for node, _ in sheets):
            raise ValueError("At least one worksheet must be visible")
        # Ensure workbook activeTab does not point to a hidden or missing sheet.
        visible = [index for index, (node, _) in enumerate(sheets) if node.get("state", "visible") == "visible"]
        for view in workbook.findall(q("bookViews") + "/" + q("workbookView")):
            if visible and int(view.get("activeTab", "0")) not in visible:
                view.set("activeTab", str(visible[0]))
        names = ET.Element(q("definedNames"))
        for name in metadata.defined_names.values():
            node = deepcopy(getattr(name, "_xml", ET.Element(q("definedName"))))
            for f in fields(name):
                if f.name != "attr_text":
                    node.attrib.pop(f.name, None)
            node.attrib.update(attrs(name, exclude=("attr_text",)))
            node.text = name.attr_text
            names.append(node)
        _replace(workbook, "definedNames", [names] if len(names) else [])
        order = ("fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst")
        workbook[:] = sorted(workbook, key=lambda n: order.index(local(n)) if local(n) in order else len(order))
        package.put(book_part, workbook)
        package.put(_rels_path(book_part), book_rels)
        package.put(style_part, styles.finish(), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        root_rels, _ = _relationships(archive, "")
        _write_properties(package, metadata, root_rels, titles)
        package.put("_rels/.rels", root_rels)
        package.default("rels", "application/vnd.openxmlformats-package.relationships+xml")
        package.default("xml", "application/xml")
        package.put("[Content_Types].xml", package.content_types)
        # Stage before opening target: safe for in-place saves and failed overlays.
        output = BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as result:
            for info in archive.zip.infolist():
                if info.filename not in package.parts:
                    result.writestr(info, archive.read(info.filename))
            for path, data in package.parts.items():
                result.writestr(path, data)
        payload = output.getvalue()
    finally:
        archive.close()
    if hasattr(target, "write"):
        target.seek(0)
        target.write(payload)
        target.truncate()
    else:
        with open(os.fspath(target), "wb") as stream:
            stream.write(payload)


__all__ = ["MetadataBook", "MetadataSheet", "CellMetadata", "load_metadata", "apply_metadata"]
