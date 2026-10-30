"""Dependency-free, copyable style objects compatible with common openpyxl usage."""

from __future__ import annotations

from dataclasses import dataclass, fields
from functools import lru_cache
from types import MappingProxyType, MemberDescriptorType
from typing import Any

# Exclude these from generic model snapshots/extra-attribute checks. In
# particular, the source cache and child index are implementation state, not XML.
STYLE_INTERNAL_FIELDS = frozenset(
    {"_style_revision", "_style_children", "_style_opaque", "_style_cache"}
)
_STYLE_PRIMITIVES = (type(None), bool, int, float, str, bytes)


@lru_cache(maxsize=32)
def _style_fields(model_type):
    return MappingProxyType(
        {field.name: index for index, field in enumerate(fields(model_type))}
    )


class _StyleModel:
    """Revision tracking for ordinary, independently writable style models.

    Tracking starts only when a clean model is first assigned to a pool. Plain
    construction, loaded XML, and thawing do not allocate per-instance tracking
    state. Afterwards only child models are visited to check cached snapshots.
    Opaque values are never cached, since mutations can bypass __setattr__.
    Direct edits to the private instance dictionary are not a supported API.
    """

    # An inherited sentinel avoids materializing instance dictionaries or raising
    # AttributeError on every constructor field assignment.
    _style_revision = None

    def __setattr__(self, name, value):
        if self._style_revision is None or name in STYLE_INTERNAL_FIELDS:
            object.__setattr__(self, name, value)
            return
        names = _style_fields(type(self))
        if name not in names and isinstance(getattr(type(self), name, None), property):
            # Alias setters re-enter here with the canonical dataclass field.
            object.__setattr__(self, name, value)
            return
        state = vars(self)
        object.__setattr__(self, name, value)
        state["_style_revision"] += 1
        state["_style_cache"] = None
        state["_style_children"].pop(name, None)
        state["_style_opaque"].discard(name)
        if name not in names:
            state["_style_opaque"].add(name)
        elif type(value) in _STYLE_TYPES:
            state["_style_children"][name] = value
        elif type(value) not in _STYLE_PRIMITIVES:
            state["_style_opaque"].add(name)

    def __delattr__(self, name):
        object.__delattr__(self, name)
        if self._style_revision is None:
            return
        state = vars(self)
        state["_style_revision"] += 1
        state["_style_cache"] = None
        state["_style_children"].pop(name, None)
        if name in _style_fields(type(self)):
            state["_style_opaque"].add(name)
        else:
            state["_style_opaque"].discard(name)

    def _style_token(self, ancestors=None):
        if type(self) not in _STYLE_TYPES:
            return None
        untracked = self._style_revision is None
        if untracked:
            state = vars(self)
            if state.keys() != _style_fields(type(self)).keys():
                return None
            model_children = {}
            for name, value in state.items():
                if type(value) in _STYLE_TYPES:
                    model_children[name] = value
                elif type(value) not in _STYLE_PRIMITIVES:
                    return None
        else:
            if self._style_opaque:
                return None
            model_children = self._style_children
        if ancestors is None:
            ancestors = set()
        identity = id(self)
        if identity in ancestors:
            return None
        ancestors.add(identity)
        try:
            children = []
            for name, child in model_children.items():
                token = child._style_token(ancestors)
                if token is None:
                    return None
                children.append((name, token))
            if untracked:
                # Enable children before the parent can cache a snapshot. Defer
                # parent activation until opaque descendants/cycles are excluded.
                state.update(
                    _style_revision=0,
                    _style_children=model_children,
                    _style_opaque=set(),
                    _style_cache=None,
                )
            return self._style_revision, tuple(children)
        finally:
            ancestors.remove(identity)

    def __copy__(self):
        result = type(self).__new__(type(self))
        for name, value in vars(self).items():
            if name not in STYLE_INTERNAL_FIELDS:
                setattr(result, name, value)
        for owner in type(self).__mro__:
            for descriptor in vars(owner).values():
                if isinstance(descriptor, MemberDescriptorType):
                    try:
                        value = descriptor.__get__(self, type(self))
                    except AttributeError:
                        continue
                    descriptor.__set__(result, value)
        return result


def _alias(name):
    return property(
        lambda self: getattr(self, name), lambda self, value: setattr(self, name, value)
    )


@dataclass
class Color(_StyleModel):
    rgb: str | None = None
    theme: int | None = None
    auto: bool | None = None
    indexed: int | None = None
    tint: float = 0.0

    def __post_init__(self):
        if self.rgb is not None:
            value = self.rgb.lstrip("#")
            if len(value) not in (6, 8) or any(
                c not in "0123456789abcdefABCDEF" for c in value
            ):
                raise ValueError(
                    "RGB colors must contain six or eight hexadecimal digits"
                )
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
class Font(_StyleModel):
    name: str | None = "Calibri"
    size: float | None = 11.0
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strike: bool = False
    color: Any = None
    vertAlign: str | None = None
    charset: int | None = None
    family: float | None = None
    scheme: str | None = None
    outline: bool | None = None
    shadow: bool | None = None
    condense: bool | None = None
    extend: bool | None = None

    def __init__(
        self,
        name: str | None = "Calibri",
        size: float | None = 11.0,
        bold=False,
        italic=False,
        underline=None,
        strike=False,
        color=None,
        vertAlign=None,
        charset=None,
        family=None,
        scheme=None,
        outline=None,
        shadow=None,
        condense=None,
        extend=None,
        b=None,
        i=None,
        sz=None,
        u=None,
    ):
        self.name, self.size = name, size if sz is None else sz
        self.bold, self.italic = bold if b is None else b, italic if i is None else i
        self.underline, self.strike = underline if u is None else u, strike
        self.color, self.vertAlign = _color(color), vertAlign
        self.charset, self.family, self.scheme = charset, family, scheme
        self.outline, self.shadow, self.condense, self.extend = (
            outline,
            shadow,
            condense,
            extend,
        )

    b = _alias("bold")
    i = _alias("italic")
    sz = _alias("size")
    u = _alias("underline")


@dataclass(init=False)
class PatternFill(_StyleModel):
    fill_type: str | None = None
    fgColor: Any = None
    bgColor: Any = None

    def __init__(
        self,
        fill_type=None,
        start_color=None,
        end_color=None,
        fgColor=None,
        bgColor=None,
        patternType=None,
    ):
        self.fill_type = fill_type if patternType is None else patternType
        self.fgColor = _color(fgColor if fgColor is not None else start_color)
        self.bgColor = _color(bgColor if bgColor is not None else end_color)

    start_color = _alias("fgColor")
    end_color = _alias("bgColor")
    patternType = _alias("fill_type")


@dataclass(init=False)
class Alignment(_StyleModel):
    horizontal: str | None = None
    vertical: str | None = None
    text_rotation: int = 0
    wrap_text: bool | None = None
    shrink_to_fit: bool | None = None
    indent: float = 0
    relativeIndent: float = 0
    justifyLastLine: bool | None = None
    readingOrder: float = 0

    def __init__(
        self,
        horizontal=None,
        vertical=None,
        text_rotation=0,
        wrap_text=None,
        shrink_to_fit=None,
        indent=0,
        relativeIndent=0,
        justifyLastLine=None,
        readingOrder=0,
        textRotation=None,
        wrapText=None,
        shrinkToFit=None,
    ):
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
class Side(_StyleModel):
    style: str | None = None
    color: Any = None

    def __init__(self, style=None, color=None, border_style=None):
        self.style = style if border_style is None else border_style
        self.color = _color(color)

    border_style = _alias("style")


@dataclass
class Border(_StyleModel):
    left: Side | None = None
    right: Side | None = None
    top: Side | None = None
    bottom: Side | None = None
    diagonal: Side | None = None
    diagonal_direction: int | None = None
    diagonalUp: bool = False
    diagonalDown: bool = False
    outline: bool = True
    vertical: Side | None = None
    horizontal: Side | None = None
    start: Side | None = None
    end: Side | None = None


@dataclass
class Protection(_StyleModel):
    locked: bool = True
    hidden: bool = False


_STYLE_TYPES = frozenset(
    {Font, PatternFill, Border, Alignment, Protection, Color, Side}
)


__all__ = ["Alignment", "Border", "Color", "Font", "PatternFill", "Protection", "Side"]
