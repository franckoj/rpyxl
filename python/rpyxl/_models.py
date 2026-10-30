"""Pure Python feature objects shared by wrappers and the OOXML overlay.

These objects do not own native cell values. Ranges and formula strings use A1
notation; callers may populate them before passing a MetadataBook to the writer.
"""
import re
from dataclasses import dataclass, field
from typing import Any, Optional

from .styles import Alignment, Border, Font, PatternFill, Protection


class _XMLModel:
    # Snapshots are annotations, not dataclass fields: they must never become
    # OOXML attributes or affect public constructor/copy/equality behavior.
    _xml: Any
    _part: str
    _rels: Any
    _parts: dict
    _content_types: Any
    _dxf_xml: Any
    _chart_xml: Any
    _drawing_part: str
    _anchor_xml: Any
    _original_text: str


def column_index(letter):
    result = 0
    for char in letter.upper():
        if not "A" <= char <= "Z":
            raise ValueError("Invalid column letter")
        result = result * 26 + ord(char) - 64
    if not 1 <= result <= 16384:
        raise ValueError("Column outside Excel limits")
    return result


def column_letter(index):
    if not 1 <= index <= 16384:
        raise ValueError("Column outside Excel limits")
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def coordinate_parts(value):
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)", str(value))
    if not match or int(match[2]) > 1048576:
        raise ValueError("Invalid cell coordinate: %r" % value)
    return column_index(match[1]), int(match[2])


def range_bounds(value):
    ends = str(value).split(":")
    if len(ends) not in (1, 2):
        raise ValueError("Invalid cell range: %r" % value)
    c1, r1 = coordinate_parts(ends[0])
    c2, r2 = coordinate_parts(ends[-1])
    if c2 < c1 or r2 < r1:
        raise ValueError("Reversed cell range")
    return c1, r1, c2, r2


@dataclass
class DefinedName(_XMLModel):
    name: str
    attr_text: Optional[str] = None
    localSheetId: Optional[int] = None
    hidden: Optional[bool] = None
    comment: Optional[str] = None
    description: Optional[str] = None
    customMenu: Optional[str] = None
    function: Optional[bool] = None
    vbProcedure: Optional[bool] = None
    xlm: Optional[bool] = None
    functionGroupId: Optional[int] = None
    shortcutKey: Optional[str] = None
    publishToServer: Optional[bool] = None
    workbookParameter: Optional[bool] = None

    @property
    def value(self):
        return self.attr_text

    @value.setter
    def value(self, value):
        self.attr_text = value

    @property
    def destinations(self):
        for part in re.finditer(r"(?:'((?:[^']|'')+)'|([^,!' ]+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)", self.attr_text or ""):
            yield (part[1] or part[2]).replace("''", "'"), part[3]


class DefinedNames(dict):
    """Global names use string keys; sheet-local names use (name, sheet_id)."""
    def add(self, value):
        self[value.name if value.localSheetId is None else (value.name, value.localSheetId)] = value

    append = add


@dataclass(init=False)
class DataValidation(_XMLModel):
    type: Optional[str] = None
    formula1: Any = None
    formula2: Any = None
    showErrorMessage: Optional[bool] = None
    showInputMessage: Optional[bool] = None
    showDropDown: Optional[bool] = None
    allowBlank: Optional[bool] = None
    operator: Optional[str] = None
    sqref: str = ""
    errorStyle: Optional[str] = None
    errorTitle: Optional[str] = None
    error: Optional[str] = None
    promptTitle: Optional[str] = None
    prompt: Optional[str] = None
    imeMode: Optional[str] = None

    def __init__(self, type=None, formula1=None, formula2=None,
                 showErrorMessage=None, showInputMessage=None, showDropDown=None,
                 allowBlank=None, operator=None, sqref="", errorStyle=None,
                 errorTitle=None, error=None, promptTitle=None, prompt=None,
                 imeMode=None, allow_blank=None):
        self.type, self.formula1, self.formula2 = type, formula1, formula2
        self.showErrorMessage, self.showInputMessage = showErrorMessage, showInputMessage
        self.showDropDown = showDropDown
        self.allowBlank = allowBlank if allow_blank is None else allow_blank
        self.operator, self.sqref = operator, sqref
        self.errorStyle, self.errorTitle, self.error = errorStyle, errorTitle, error
        self.promptTitle, self.prompt, self.imeMode = promptTitle, prompt, imeMode

    def add(self, cell):
        value = getattr(cell, "coordinate", cell)
        range_bounds(value)
        self.sqref = (str(self.sqref) + " " + value).strip()

    @property
    def allow_blank(self):
        return self.allowBlank

    @allow_blank.setter
    def allow_blank(self, value):
        self.allowBlank = value

    @property
    def ranges(self):
        return self.sqref


class DataValidations(list, _XMLModel):
    def add(self, value):
        self.append(value)

    @property
    def dataValidation(self):
        return self


@dataclass
class TableStyleInfo(_XMLModel):
    name: Optional[str] = None
    showFirstColumn: Optional[bool] = None
    showLastColumn: Optional[bool] = None
    showRowStripes: Optional[bool] = None
    showColumnStripes: Optional[bool] = None


@dataclass
class TableColumn(_XMLModel):
    id: int
    name: str


@dataclass
class Table(_XMLModel):
    displayName: str
    ref: str
    name: Optional[str] = None
    id: Optional[int] = None
    headerRowCount: int = 1
    totalsRowCount: Optional[int] = None
    totalsRowShown: Optional[bool] = None
    tableStyleInfo: Optional[TableStyleInfo] = None
    tableColumns: list = field(default_factory=list)

    def __post_init__(self):
        if self.name is None:
            self.name = self.displayName
        range_bounds(self.ref)


class Tables(dict):
    def add(self, table):
        if table.displayName in self:
            raise ValueError("Table already exists: " + table.displayName)
        self[table.displayName] = table


@dataclass
class Comment(_XMLModel):
    text: str
    author: str
    height: float = 79
    width: float = 144


@dataclass
class Hyperlink(_XMLModel):
    ref: Optional[str] = None
    target: Optional[str] = None
    location: Optional[str] = None
    tooltip: Optional[str] = None
    display: Optional[str] = None
    id: Optional[str] = None


@dataclass
class Rule(_XMLModel):
    type: str
    operator: Optional[str] = None
    formula: list = field(default_factory=list)
    stopIfTrue: Optional[bool] = None
    priority: int = 0
    font: Optional[Font] = None
    fill: Optional[PatternFill] = None
    border: Optional[Border] = None
    colorScale: Any = None


def CellIsRule(operator=None, formula=None, stopIfTrue=None, font=None, border=None, fill=None):
    return Rule("cellIs", operator, list(formula or []), stopIfTrue, font=font, border=border, fill=fill)


def FormulaRule(formula=None, stopIfTrue=None, font=None, border=None, fill=None):
    return Rule("expression", formula=list(formula or []), stopIfTrue=stopIfTrue, font=font, border=border, fill=fill)


def ColorScaleRule(start_type=None, start_value=None, start_color=None,
                   mid_type=None, mid_value=None, mid_color=None,
                   end_type=None, end_value=None, end_color=None):
    stops = [(start_type, start_value, start_color)]
    if mid_type is not None:
        stops.append((mid_type, mid_value, mid_color))
    stops.append((end_type, end_value, end_color))
    return Rule("colorScale", colorScale=stops)


class ConditionalFormatting:
    def __init__(self):
        self._rules = {}

    def add(self, cell_range, rule):
        for value in str(cell_range).split():
            range_bounds(value)
        self._rules.setdefault(str(cell_range), []).append(rule)

    def __iter__(self):
        return iter(self._rules)

    def __getitem__(self, key):
        return self._rules[str(key)]

    def __len__(self):
        return len(self._rules)


@dataclass(init=False)
class Reference:
    sheetname: str
    min_col: int
    min_row: int
    max_col: int
    max_row: int

    def __init__(self, worksheet=None, min_col=1, min_row=1, max_col=None, max_row=None, range_string=None):
        if range_string is not None:
            sheet, cells = range_string.rsplit("!", 1)
            self.sheetname = sheet.strip("'").replace("''", "'")
            self.min_col, self.min_row, self.max_col, self.max_row = range_bounds(cells)
        else:
            self.sheetname = getattr(worksheet, "title", worksheet)
            if not isinstance(self.sheetname, str):
                raise ValueError("Reference requires a worksheet or sheet title")
            self.min_col, self.min_row = min_col, min_row
            self.max_col = min_col if max_col is None else max_col
            self.max_row = min_row if max_row is None else max_row
            range_bounds(self.cells)

    @property
    def cells(self):
        return "$%s$%s:$%s$%s" % (column_letter(self.min_col), self.min_row, column_letter(self.max_col), self.max_row)

    def __str__(self):
        return "'%s'!%s" % (self.sheetname.replace("'", "''"), self.cells)


@dataclass
class Series(_XMLModel):
    values: str
    title: Optional[str] = None
    title_reference: Optional[str] = None
    categories: Optional[str] = None


@dataclass
class BarChart(_XMLModel):
    title: Optional[str] = None
    type: str = "col"
    style: Optional[int] = None
    grouping: str = "clustered"
    overlap: Optional[int] = None
    gapWidth: int = 150
    width: float = 15
    height: float = 7.5
    anchor: str = "A1"
    series: list = field(default_factory=list)
    x_axis: Any = field(default_factory=lambda: ChartAxis())
    y_axis: Any = field(default_factory=lambda: ChartAxis())

    def add_data(self, data, titles_from_data=False, from_rows=False):
        outer = range(data.min_row, data.max_row + 1) if from_rows else range(data.min_col, data.max_col + 1)
        for index in outer:
            c1, r1 = (data.min_col, index) if from_rows else (index, data.min_row)
            c2, r2 = (data.max_col, index) if from_rows else (index, data.max_row)
            title_ref = None
            if titles_from_data:
                title_ref = str(Reference(data.sheetname, c1, r1, c1, r1))
                if from_rows:
                    c1 += 1
                else:
                    r1 += 1
            self.series.append(Series(str(Reference(data.sheetname, c1, r1, c2, r2)), title_reference=title_ref))

    def set_categories(self, labels):
        for series in self.series:
            series.categories = str(labels)


@dataclass
class ChartAxis:
    title: Optional[str] = None


@dataclass
class CellMetadata:
    font: Optional[Font] = None
    fill: Optional[PatternFill] = None
    border: Optional[Border] = None
    alignment: Optional[Alignment] = None
    protection: Optional[Protection] = None
    hyperlink: Any = None
    comment: Optional[Comment] = None


@dataclass(frozen=True)
class CellRange:
    coord: str

    def __post_init__(self):
        range_bounds(self.coord)

    def __str__(self):
        return self.coord


class MergedCells:
    def __init__(self):
        self.ranges = []

    def add(self, value):
        cell_range = value if isinstance(value, CellRange) else CellRange(str(value))
        if cell_range not in self.ranges:
            self.ranges.append(cell_range)

    def remove(self, value):
        self.ranges.remove(value if isinstance(value, CellRange) else CellRange(str(value)))

    def __iter__(self):
        return iter(self.ranges)


@dataclass
class ColumnDimension(_XMLModel):
    width: Optional[float] = None
    hidden: Optional[bool] = None
    min: Optional[int] = None
    max: Optional[int] = None
    bestFit: Optional[bool] = None
    outlineLevel: Optional[int] = None
    collapsed: Optional[bool] = None


@dataclass
class RowDimension(_XMLModel):
    height: Optional[float] = None
    hidden: Optional[bool] = None
    outlineLevel: Optional[int] = None
    collapsed: Optional[bool] = None


class Dimensions(dict):
    def __init__(self, factory):
        super().__init__()
        self.factory = factory

    def __missing__(self, key):
        value = self[key] = self.factory()
        return value


@dataclass
class AutoFilter(_XMLModel):
    ref: Optional[str] = None


@dataclass
class SheetProperties(_XMLModel):
    tabColor: Any = None


@dataclass
class SheetProtection(_XMLModel):
    sheet: Optional[bool] = None
    objects: Optional[bool] = None
    scenarios: Optional[bool] = None
    formatCells: Optional[bool] = None
    formatColumns: Optional[bool] = None
    formatRows: Optional[bool] = None
    insertColumns: Optional[bool] = None
    insertRows: Optional[bool] = None
    insertHyperlinks: Optional[bool] = None
    deleteColumns: Optional[bool] = None
    deleteRows: Optional[bool] = None
    selectLockedCells: Optional[bool] = None
    sort: Optional[bool] = None
    autoFilter: Optional[bool] = None
    pivotTables: Optional[bool] = None
    selectUnlockedCells: Optional[bool] = None
    password: Optional[str] = None
    algorithmName: Optional[str] = None
    hashValue: Optional[str] = None
    saltValue: Optional[str] = None
    spinCount: Optional[int] = None

    def set_password(self, value, already_hashed=False):
        if already_hashed:
            self.password = value
        else:
            result = 0
            for index, char in enumerate(value, 1):
                rotated = ord(char) << index
                result ^= (rotated & 0x7fff) | (rotated >> 15)
            self.password = "%04X" % (result ^ len(value) ^ 0xCE4B)
        self.sheet = True

    def enable(self):
        self.sheet = True

    def disable(self):
        self.sheet = False


@dataclass
class PageSetup(_XMLModel):
    orientation: Optional[str] = None
    paperSize: Optional[int] = None
    scale: Optional[int] = None
    fitToHeight: Optional[int] = None
    fitToWidth: Optional[int] = None
    firstPageNumber: Optional[int] = None
    useFirstPageNumber: Optional[bool] = None
    blackAndWhite: Optional[bool] = None
    draft: Optional[bool] = None
    cellComments: Optional[str] = None
    errors: Optional[str] = None
    horizontalDpi: Optional[int] = None
    verticalDpi: Optional[int] = None
    copies: Optional[int] = None


@dataclass
class DocumentProperties:
    creator: Optional[str] = None
    title: Optional[str] = None
    subject: Optional[str] = None
    description: Optional[str] = None
    keywords: Optional[str] = None
    lastModifiedBy: Optional[str] = None
    revision: Optional[str] = None
    category: Optional[str] = None
    contentStatus: Optional[str] = None
    version: Optional[str] = None
    identifier: Optional[str] = None
    language: Optional[str] = None
    created: Any = None
    modified: Any = None
    lastPrinted: Any = None
    application: Optional[str] = None
    appVersion: Optional[str] = None
    company: Optional[str] = None
    manager: Optional[str] = None


@dataclass
class MetadataSheet(_XMLModel):
    title: str = "Sheet"
    cells: dict = field(default_factory=dict)
    merged_cells: MergedCells = field(default_factory=MergedCells)
    column_dimensions: Dimensions = field(default_factory=lambda: Dimensions(ColumnDimension))
    row_dimensions: Dimensions = field(default_factory=lambda: Dimensions(RowDimension))
    freeze_panes: Optional[str] = None
    auto_filter: AutoFilter = field(default_factory=AutoFilter)
    sheet_properties: SheetProperties = field(default_factory=SheetProperties)
    sheet_state: str = "visible"
    protection: SheetProtection = field(default_factory=SheetProtection)
    page_setup: PageSetup = field(default_factory=PageSetup)
    data_validations: DataValidations = field(default_factory=DataValidations)
    conditional_formatting: ConditionalFormatting = field(default_factory=ConditionalFormatting)
    tables: Tables = field(default_factory=Tables)
    charts: list = field(default_factory=list)

    def add_chart(self, chart, anchor=None):
        if anchor is not None:
            chart.anchor = anchor
        self.charts.append(chart)


@dataclass
class MetadataBook(_XMLModel):
    properties: DocumentProperties = field(default_factory=DocumentProperties)
    defined_names: DefinedNames = field(default_factory=DefinedNames)
    sheets: dict = field(default_factory=dict)


__all__ = [name for name, value in list(globals().items())
           if not name.startswith("_") and getattr(value, "__module__", None) == __name__]
