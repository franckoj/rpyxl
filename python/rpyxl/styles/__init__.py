"""Dependency-free, copyable style objects compatible with common openpyxl usage."""

from dataclasses import dataclass
from typing import Any, Optional


def _alias(name):
    return property(lambda self: getattr(self, name), lambda self, value: setattr(self, name, value))


@dataclass
class Color:
    rgb: Optional[str] = None
    theme: Optional[int] = None
    auto: Optional[bool] = None
    indexed: Optional[int] = None
    tint: float = 0.0

    def __post_init__(self):
        if self.rgb is not None:
            value = self.rgb.lstrip("#")
            if len(value) not in (6, 8) or any(c not in "0123456789abcdefABCDEF" for c in value):
                raise ValueError("RGB colors must contain six or eight hexadecimal digits")
            self.rgb = ("00" + value if len(value) == 6 else value).upper()

    @property
    def type(self):
        for key in ("rgb", "indexed", "theme", "auto"):
            if getattr(self, key) is not None:
                return key
        return "rgb"

    @property
    def value(self):
        return getattr(self, self.type)


def _color(value):
    return Color(rgb=value) if isinstance(value, str) else value


@dataclass(init=False)
class Font:
    name: Optional[str] = "Calibri"
    size: Optional[float] = 11.0
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None
    strike: bool = False
    color: Any = None
    vertAlign: Optional[str] = None
    charset: Optional[int] = None
    family: Optional[float] = None
    scheme: Optional[str] = None
    outline: Optional[bool] = None
    shadow: Optional[bool] = None
    condense: Optional[bool] = None
    extend: Optional[bool] = None

    def __init__(self, name: Optional[str] = "Calibri", size: Optional[float] = 11.0, bold=False, italic=False,
                 underline=None, strike=False, color=None, vertAlign=None,
                 charset=None, family=None, scheme=None, outline=None,
                 shadow=None, condense=None, extend=None, b=None, i=None,
                 sz=None, u=None):
        self.name, self.size = name, size if sz is None else sz
        self.bold, self.italic = bold if b is None else b, italic if i is None else i
        self.underline, self.strike = underline if u is None else u, strike
        self.color, self.vertAlign = _color(color), vertAlign
        self.charset, self.family, self.scheme = charset, family, scheme
        self.outline, self.shadow, self.condense, self.extend = outline, shadow, condense, extend

    b = _alias("bold")
    i = _alias("italic")
    sz = _alias("size")
    u = _alias("underline")


@dataclass(init=False)
class PatternFill:
    fill_type: Optional[str] = None
    fgColor: Any = None
    bgColor: Any = None

    def __init__(self, fill_type=None, start_color=None, end_color=None,
                 fgColor=None, bgColor=None, patternType=None):
        self.fill_type = fill_type if patternType is None else patternType
        self.fgColor = _color(fgColor if fgColor is not None else start_color)
        self.bgColor = _color(bgColor if bgColor is not None else end_color)

    start_color = _alias("fgColor")
    end_color = _alias("bgColor")
    patternType = _alias("fill_type")


@dataclass(init=False)
class Alignment:
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    text_rotation: int = 0
    wrap_text: Optional[bool] = None
    shrink_to_fit: Optional[bool] = None
    indent: float = 0
    relativeIndent: float = 0
    justifyLastLine: Optional[bool] = None
    readingOrder: float = 0

    def __init__(self, horizontal=None, vertical=None, text_rotation=0,
                 wrap_text=None, shrink_to_fit=None, indent=0, relativeIndent=0,
                 justifyLastLine=None, readingOrder=0, textRotation=None,
                 wrapText=None, shrinkToFit=None):
        self.horizontal, self.vertical = horizontal, vertical
        self.text_rotation = text_rotation if textRotation is None else textRotation
        self.wrap_text = wrap_text if wrapText is None else wrapText
        self.shrink_to_fit = shrink_to_fit if shrinkToFit is None else shrinkToFit
        self.indent, self.relativeIndent = indent, relativeIndent
        self.justifyLastLine, self.readingOrder = justifyLastLine, readingOrder

    textRotation = _alias("text_rotation")
    wrapText = _alias("wrap_text")
    shrinkToFit = _alias("shrink_to_fit")


@dataclass(init=False)
class Side:
    style: Optional[str] = None
    color: Any = None

    def __init__(self, style=None, color=None, border_style=None):
        self.style = style if border_style is None else border_style
        self.color = _color(color)

    border_style = _alias("style")


@dataclass
class Border:
    left: Optional[Side] = None
    right: Optional[Side] = None
    top: Optional[Side] = None
    bottom: Optional[Side] = None
    diagonal: Optional[Side] = None
    diagonal_direction: Optional[int] = None
    diagonalUp: bool = False
    diagonalDown: bool = False
    outline: bool = True
    vertical: Optional[Side] = None
    horizontal: Optional[Side] = None
    start: Optional[Side] = None
    end: Optional[Side] = None


@dataclass
class Protection:
    locked: bool = True
    hidden: bool = False


__all__ = ["Font", "PatternFill", "Alignment", "Side", "Border", "Color", "Protection"]
