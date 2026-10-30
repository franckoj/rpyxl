"""Small wrapper helpers; cell values and number formats remain native-owned."""

from copy import deepcopy
from io import BytesIO
from xml.etree import ElementTree as ET

from ._models import CellMetadata, column_letter, coordinate_parts, range_bounds
from .styles import Alignment, Border, Color, Font, PatternFill, Protection, Side


def default_font():
    return Font(color=Color(theme=1), family=2, scheme="minor")


def default_fill():
    return PatternFill(fgColor=Color(rgb="00000000"), bgColor=Color(rgb="00000000"))


def default_border():
    return Border(left=Side(), right=Side(), top=Side(), bottom=Side(), diagonal=Side())


STYLE_DEFAULTS = {
    "font": default_font,
    "fill": default_fill,
    "border": default_border,
    "alignment": Alignment,
    "protection": Protection,
}


def metadata_property(name):
    def get(self):
        return getattr(self._metadata, name)

    def set(self, value):
        self._check_writable()
        setattr(self._metadata, name, value)

    return property(get, set)


def cell_metadata_property(name):
    def get(self):
        cells = self.parent._metadata.cells
        coordinate = self.coordinate
        metadata = cells.get(coordinate)
        if metadata is None:
            metadata = cells[coordinate] = CellMetadata()
        value = getattr(metadata, name)
        workbook = self.parent.parent
        pool = workbook._style_pool
        if value is None and name in STYLE_DEFAULTS:
            if name not in workbook._style_defaults:
                workbook._style_defaults[name] = pool.assign(STYLE_DEFAULTS[name]())
            value = workbook._style_defaults[name]
            setattr(metadata, name, value)
        return pool.view(
            lambda: getattr(metadata, name),
            lambda assigned: setattr(metadata, name, assigned),
            self.parent._check_writable,
        )

    def set(self, value):
        self.parent._check_writable()
        pool = self.parent.parent._style_pool
        cells = self.parent._metadata.cells
        coordinate = self.coordinate
        metadata = cells.get(coordinate)
        if metadata is None:
            metadata = cells[coordinate] = CellMetadata()
        setattr(metadata, name, pool.assign(value))

    return property(get, set)


def coordinate(col, row):
    if not 1 <= row <= 1048576:
        raise ValueError("Row outside Excel limits")
    return column_letter(col) + str(row)


def structural_metadata(sheet, axis, idx, amount, delete=False):
    """Prepare a sparse sidecar transaction before touching native values."""
    result = deepcopy(sheet)
    if not amount:
        return result

    def shift(value):
        if delete and idx <= value < idx + amount:
            return None
        return value + (-amount if delete else amount) if value >= idx else value

    def point(col, row):
        value = shift(row if axis == "row" else col)
        if value is None:
            return None
        return coordinate(
            col if axis == "row" else value, value if axis == "row" else row
        )

    result.cells = {}
    for ref, metadata in sheet.cells.items():
        new_ref = point(*coordinate_parts(ref))
        if new_ref:
            if metadata.hyperlink is not None and not isinstance(
                metadata.hyperlink, str
            ):
                metadata = deepcopy(metadata)
                metadata.hyperlink.ref = new_ref
            result.cells[new_ref] = metadata

    dimensions = result.row_dimensions if axis == "row" else result.column_dimensions
    items = list(dimensions.items())
    dimensions.clear()
    for key, dimension in items:
        old = key if axis == "row" else coordinate_parts(str(key) + "1")[0]
        # Grouped columns need range-aware rewriting, not just a key shift.
        if axis == "col" and dimension.max and dimension.min != dimension.max:
            raise NotImplementedError(
                "Structural edits of grouped column dimensions are not supported"
            )
        new = shift(old)
        if new is not None:
            coordinate(1, new) if axis == "row" else column_letter(new)
            if axis == "col":
                dimension.min = dimension.max = new
            dimensions[new if axis == "row" else column_letter(new)] = dimension
    return result


def move_metadata(sheet, cell_range, rows, cols):
    result = deepcopy(sheet)
    c1, r1, c2, r2 = range_bounds(cell_range)
    coordinate(c1 + cols, r1 + rows)
    coordinate(c2 + cols, r2 + rows)
    moving = {}
    for ref, metadata in sheet.cells.items():
        col, row = coordinate_parts(ref)
        if c1 <= col <= c2 and r1 <= row <= r2:
            new_ref = coordinate(col + cols, row + rows)
            metadata = deepcopy(metadata)
            if metadata.hyperlink is not None and not isinstance(
                metadata.hyperlink, str
            ):
                metadata.hyperlink.ref = new_ref
            moving[new_ref] = metadata
            result.cells.pop(ref, None)
        elif c1 + cols <= col <= c2 + cols and r1 + rows <= row <= r2 + rows:
            result.cells.pop(ref, None)
    result.cells.update(moving)
    return result


# Keep this aligned with the native built-in format spellings. Read-only native
# cells currently omit their format; this sparse fallback is loaded on demand.
BUILTIN_FORMATS = dict(
    enumerate(
        [
            "General",
            "0",
            "0.00",
            "#,##0",
            "#,##0.00",
            '"$"#,##0_);("$"#,##0)',
            '"$"#,##0_);[Red]("$"#,##0)',
            '"$"#,##0.00_);("$"#,##0.00)',
            '"$"#,##0.00_);[Red]("$"#,##0.00)',
            "0%",
            "0.00%",
            "0.00E+00",
            "# ?/?",
            "# ??/??",
            "mm-dd-yy",
            "d-mmm-yy",
            "d-mmm",
            "mmm-yy",
            "h:mm AM/PM",
            "h:mm:ss AM/PM",
            "h:mm",
            "h:mm:ss",
            "m/d/yy h:mm",
        ]
    )
)
BUILTIN_FORMATS.update(
    {
        37: "#,##0 ;(#,##0)",
        38: "#,##0 ;[Red](#,##0)",
        39: "#,##0.00;(#,##0.00)",
        40: "#,##0.00;[Red](#,##0.00)",
        41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
        42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
        43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
        44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
        45: "mm:ss",
        46: "[h]:mm:ss",
        47: "mmss.0",
        48: "##0.0E+0",
        49: "@",
    }
)


def read_number_formats(source, title):
    from ._metadata import _Archive, _book_parts, q

    archive = _Archive(source)
    try:
        _, _, sheets, styles_part = _book_parts(archive)
        styles = archive.root(styles_part) if styles_part else None
        formats = dict(BUILTIN_FORMATS)
        xfs = ["General"]
        if styles is not None:
            formats.update(
                {
                    int(n.attrib["numFmtId"]): n.attrib["formatCode"]
                    for n in styles.findall(q("numFmts") + "/" + q("numFmt"))
                }
            )
            xfs = [
                formats.get(int(n.get("numFmtId", "0")), "General")
                for n in styles.findall(q("cellXfs") + "/" + q("xf"))
            ]
        part = next(part for node, part in sheets if node.get("name") == title)
        data = archive.read(part)
        # Reuse the metadata parser's DTD/entity rejection before iterparse.
        probe = data.replace(b"\x00", b"").upper()
        if b"<!DOCTYPE" in probe or b"<!ENTITY" in probe:
            raise ValueError("DTD/entity declarations are not allowed in XLSX metadata")
        result = {}
        stack = []
        for event, node in ET.iterparse(BytesIO(data), events=("start", "end")):
            if event == "start":
                stack.append(node)
                continue
            if node.tag == q("c") and "s" in node.attrib:
                index = int(node.attrib["s"])
                result[node.attrib["r"]] = (
                    xfs[index] if 0 <= index < len(xfs) else "General"
                )
            if len(stack) > 1:
                stack[-2].remove(node)
            stack.pop()
            node.clear()
        return result
    finally:
        archive.close()
