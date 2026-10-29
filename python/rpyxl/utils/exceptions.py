"""Exceptions matching openpyxl.utils.exceptions hierarchy."""

class OpenpyxlException(Exception):
    """Base class for all rpyxl and openpyxl-compatible exceptions."""
    pass


class InvalidFileException(OpenpyxlException, ValueError):
    """Raised when an XLSX file is invalid, corrupted, or not an archive."""
    pass


class SheetTitleException(OpenpyxlException, ValueError):
    """Raised when a worksheet title is invalid."""
    pass


class CellCoordinatesException(OpenpyxlException, ValueError):
    """Raised when a cell coordinate string is malformed."""
    pass


class NamedRangeException(OpenpyxlException):
    """Raised when an error occurs with a named range."""
    pass


__all__ = [
    "OpenpyxlException",
    "InvalidFileException",
    "SheetTitleException",
    "CellCoordinatesException",
    "NamedRangeException",
]
