"""Coordinate utilities for openpyxl compatibility."""

import re


def get_column_letter(col_idx: int) -> str:
    """Convert a 1-based column index to an Excel column letter."""
    if not (1 <= col_idx <= 16384):
        raise ValueError(f"Column index {col_idx} is out of bounds (allowed: 1-16384)")
    col = col_idx
    letters = []
    while col > 0:
        rem = (col - 1) % 26
        letters.append(chr(ord("A") + rem))
        col = (col - 1) // 26
    return "".join(reversed(letters))


def column_index_from_string(col_str: str) -> int:
    """Convert an Excel column letter to a 1-based column index."""
    clean = col_str.strip().lstrip("$").upper()
    if not clean or not clean.isalpha():
        raise ValueError(f"Invalid column string: {col_str}")
    col = 0
    for char in clean:
        col = col * 26 + (ord(char) - ord("A") + 1)
    if not (1 <= col <= 16384):
        raise ValueError(f"Column index for {col_str} is out of bounds (allowed: 1-16384)")
    return col


def coordinate_to_tuple(coord: str) -> tuple[int, int]:
    """Convert a coordinate string (e.g. 'A1', '$B$10') to (row, col) (1-based)."""
    s = coord.strip()
    if s and s.isascii():
        i = 1 if s[0] == "$" else 0
        col_start = i
        while i < len(s) and ("A" <= s[i] <= "Z" or "a" <= s[i] <= "z"):
            i += 1
        col_end = i
        if col_end > col_start and col_end - col_start <= 3:
            if i < len(s) and s[i] == "$":
                i += 1
            row_start = i
            while i < len(s) and "0" <= s[i] <= "9":
                i += 1
            if i == len(s) and i > row_start:
                row = int(s[row_start:i])
                if 1 <= row <= 1048576:
                    col = 0
                    for ch in s[col_start:col_end].upper():
                        col = col * 26 + (ord(ch) - 64)
                    if 1 <= col <= 16384:
                        return (row, col)
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?([0-9]+)", s)
    if match is None:
        raise ValueError(f"Invalid coordinate format: {coord}")
    letters, digits = match.groups()
    row = int(digits)
    col = column_index_from_string(letters)
    if not (1 <= row <= 1048576):
        raise ValueError(f"Row {row} is out of bounds (allowed: 1-1048576)")
    return (row, col)


def range_boundaries(range_string: str) -> tuple[int, int, int, int]:
    """
    Convert a range string (e.g. 'A1:C5' or 'A1') to (min_col, min_row, max_col, max_row).
    Matches openpyxl.utils.range_boundaries.
    """
    clean = range_string.strip()
    if ":" in clean:
        parts = clean.split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid range format: {range_string}")
        min_row, min_col = coordinate_to_tuple(parts[0])
        max_row, max_col = coordinate_to_tuple(parts[1])
        return (
            min(min_col, max_col),
            min(min_row, max_row),
            max(min_col, max_col),
            max(min_row, max_row),
        )
    else:
        row, col = coordinate_to_tuple(clean)
        return (col, row, col, row)


def quote_sheetname(sheetname: str) -> str:
    """Add quotes around sheetnames matching openpyxl.utils.quote_sheetname."""
    if "'" in sheetname:
        sheetname = sheetname.replace("'", "''")
    return f"'{sheetname}'"
