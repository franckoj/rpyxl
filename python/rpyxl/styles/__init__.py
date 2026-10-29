"""Styles foundation module compatible with openpyxl.styles."""

from dataclasses import dataclass
from typing import Optional, Any


@dataclass
class Font:
    name: Optional[str] = "Calibri"
    size: Optional[float] = 11.0
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None
    strike: bool = False
    color: Optional[Any] = None
    vertAlign: Optional[str] = None


@dataclass
class PatternFill:
    fill_type: Optional[str] = None
    start_color: Optional[Any] = None
    end_color: Optional[Any] = None
    fgColor: Optional[Any] = None
    bgColor: Optional[Any] = None


@dataclass
class Alignment:
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    text_rotation: int = 0
    wrap_text: Optional[bool] = None
    shrink_to_fit: Optional[bool] = None
    indent: int = 0


@dataclass
class Side:
    style: Optional[str] = None
    color: Optional[Any] = None


@dataclass
class Border:
    left: Optional[Side] = None
    right: Optional[Side] = None
    top: Optional[Side] = None
    bottom: Optional[Side] = None
    diagonal: Optional[Side] = None


@dataclass
class Color:
    rgb: Optional[str] = None
    theme: Optional[int] = None
    auto: Optional[bool] = None


__all__ = [
    "Font",
    "PatternFill",
    "Alignment",
    "Side",
    "Border",
    "Color",
]
