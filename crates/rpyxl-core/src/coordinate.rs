use crate::error::{RpyxlError, RpyxlResult};

pub const MIN_ROW: u32 = 1;
pub const MAX_ROW: u32 = 1_048_576;
pub const MIN_COLUMN: u32 = 1;
pub const MAX_COLUMN: u32 = 16_384;

const INVALID_TITLE_CHARS: [char; 7] = ['\\', '/', '?', '*', ':', '[', ']'];

/// Write column letters directly into an existing buffer without heap allocation.
#[inline]
pub fn write_column_letter(col_idx: u32, out: &mut String) -> RpyxlResult<()> {
    if !(MIN_COLUMN..=MAX_COLUMN).contains(&col_idx) {
        return Err(RpyxlError::ColumnOutOfBounds(col_idx));
    }

    let mut col = col_idx;
    let mut buf = [0u8; 3];
    let mut idx = 3;
    while col > 0 {
        idx -= 1;
        let rem = ((col - 1) % 26) as u8;
        buf[idx] = b'A' + rem;
        col = (col - 1) / 26;
    }
    // Safety: buf[idx..3] only contains ASCII bytes b'A'..=b'Z'
    let s = unsafe { std::str::from_utf8_unchecked(&buf[idx..3]) };
    out.push_str(s);
    Ok(())
}

/// Convert a column index (1-based) to an Excel column letter (e.g., 1 -> "A", 27 -> "AA").
pub fn get_column_letter(col_idx: u32) -> RpyxlResult<String> {
    let mut s = String::with_capacity(3);
    write_column_letter(col_idx, &mut s)?;
    Ok(s)
}

/// Convert an Excel column letter (e.g., "A", "AA", "xfd") to a 1-based column index.
pub fn column_index_from_string(col_str: &str) -> RpyxlResult<u32> {
    let clean = col_str.trim().trim_start_matches('$');
    if clean.is_empty() {
        return Err(RpyxlError::InvalidCoordinate(col_str.to_string()));
    }

    let mut col: u64 = 0;
    for c in clean.chars() {
        if !c.is_ascii_alphabetic() {
            return Err(RpyxlError::InvalidCoordinate(col_str.to_string()));
        }
        let val = (c.to_ascii_uppercase() as u64) - ('A' as u64) + 1;
        col = col * 26 + val;
        if col > MAX_COLUMN as u64 {
            return Err(RpyxlError::ColumnOutOfBounds(col as u32));
        }
    }

    if col < MIN_COLUMN as u64 {
        return Err(RpyxlError::ColumnOutOfBounds(col as u32));
    }

    Ok(col as u32)
}

/// Parse a cell coordinate string (e.g., "A1", "$B$12", "c3") into (row, col) (1-based).
pub fn coordinate_to_tuple(coord: &str) -> RpyxlResult<(u32, u32)> {
    let clean = coord.trim();
    if clean.is_empty() {
        return Err(RpyxlError::InvalidCoordinate(coord.to_string()));
    }

    let mut col_str = String::new();
    let mut row_str = String::new();

    for c in clean.chars() {
        if c == '$' {
            continue;
        } else if c.is_ascii_alphabetic() {
            if !row_str.is_empty() {
                // Cannot have letters after digits (e.g., "A1B")
                return Err(RpyxlError::InvalidCoordinate(coord.to_string()));
            }
            col_str.push(c);
        } else if c.is_ascii_digit() {
            row_str.push(c);
        } else {
            return Err(RpyxlError::InvalidCoordinate(coord.to_string()));
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return Err(RpyxlError::InvalidCoordinate(coord.to_string()));
    }

    let col = column_index_from_string(&col_str)?;
    let row: u32 = row_str
        .parse()
        .map_err(|_| RpyxlError::InvalidCoordinate(coord.to_string()))?;

    if !(MIN_ROW..=MAX_ROW).contains(&row) {
        return Err(RpyxlError::RowOutOfBounds(row));
    }

    Ok((row, col))
}

/// Write (row, col) (1-based) directly into an existing buffer as coordinate (e.g., (1, 1) -> "A1").
#[inline]
pub fn write_coordinate(row: u32, column: u32, out: &mut String) -> RpyxlResult<()> {
    if !(MIN_ROW..=MAX_ROW).contains(&row) {
        return Err(RpyxlError::RowOutOfBounds(row));
    }
    write_column_letter(column, out)?;
    let mut buf = itoa::Buffer::new();
    out.push_str(buf.format(row));
    Ok(())
}

/// Convert (row, col) (1-based) to an Excel coordinate string (e.g., (1, 1) -> "A1").
pub fn tuple_to_coordinate(row: u32, column: u32) -> RpyxlResult<String> {
    let mut s = String::with_capacity(10);
    write_coordinate(row, column, &mut s)?;
    Ok(s)
}

/// Parse a range string (e.g. "A1:B10" or "A1") into (min_row, min_col, max_row, max_col) 1-based bounds.
pub fn range_boundaries(range_str: &str) -> RpyxlResult<(u32, u32, u32, u32)> {
    let clean = range_str.trim();
    if clean.is_empty() {
        return Err(RpyxlError::InvalidRange(range_str.to_string()));
    }

    let parts: Vec<&str> = clean.split(':').collect();
    match parts.len() {
        1 => {
            let (row, col) = coordinate_to_tuple(parts[0])?;
            Ok((row, col, row, col))
        }
        2 => {
            let (r1, c1) = coordinate_to_tuple(parts[0])?;
            let (r2, c2) = coordinate_to_tuple(parts[1])?;
            let min_row = r1.min(r2);
            let max_row = r1.max(r2);
            let min_col = c1.min(c2);
            let max_col = c1.max(c2);
            Ok((min_row, min_col, max_row, max_col))
        }
        _ => Err(RpyxlError::InvalidRange(range_str.to_string())),
    }
}

/// Validate worksheet title according to Excel rules.
pub fn validate_sheet_title(title: &str) -> RpyxlResult<()> {
    if title.is_empty() {
        return Err(RpyxlError::InvalidSheetTitle(
            "Title cannot be empty".to_string(),
        ));
    }

    if title.chars().count() > 31 {
        return Err(RpyxlError::InvalidSheetTitle(format!(
            "Title '{}' exceeds maximum allowed length of 31 characters (length: {})",
            title,
            title.chars().count()
        )));
    }

    if title.starts_with('\'') || title.ends_with('\'') {
        return Err(RpyxlError::InvalidSheetTitle(format!(
            "Title '{}' cannot start or end with an apostrophe",
            title
        )));
    }

    for invalid_char in INVALID_TITLE_CHARS {
        if title.contains(invalid_char) {
            return Err(RpyxlError::InvalidSheetTitle(format!(
                "Title '{}' contains invalid character '{}'",
                title, invalid_char
            )));
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_letter_conversions() {
        assert_eq!(get_column_letter(1).unwrap(), "A");
        assert_eq!(get_column_letter(26).unwrap(), "Z");
        assert_eq!(get_column_letter(27).unwrap(), "AA");
        assert_eq!(get_column_letter(28).unwrap(), "AB");
        assert_eq!(get_column_letter(702).unwrap(), "ZZ");
        assert_eq!(get_column_letter(703).unwrap(), "AAA");
        assert_eq!(get_column_letter(16384).unwrap(), "XFD");
        assert!(get_column_letter(0).is_err());
        assert!(get_column_letter(16385).is_err());
    }

    #[test]
    fn test_column_index_from_string() {
        assert_eq!(column_index_from_string("A").unwrap(), 1);
        assert_eq!(column_index_from_string("a").unwrap(), 1);
        assert_eq!(column_index_from_string("Z").unwrap(), 26);
        assert_eq!(column_index_from_string("AA").unwrap(), 27);
        assert_eq!(column_index_from_string("ab").unwrap(), 28);
        assert_eq!(column_index_from_string("ZZ").unwrap(), 702);
        assert_eq!(column_index_from_string("AAA").unwrap(), 703);
        assert_eq!(column_index_from_string("XFD").unwrap(), 16384);
        assert_eq!(column_index_from_string("$A").unwrap(), 1);
        assert!(column_index_from_string("XFE").is_err());
        assert!(column_index_from_string("").is_err());
        assert!(column_index_from_string("123").is_err());
    }

    #[test]
    fn test_coordinate_to_tuple() {
        assert_eq!(coordinate_to_tuple("A1").unwrap(), (1, 1));
        assert_eq!(coordinate_to_tuple("$B$10").unwrap(), (10, 2));
        assert_eq!(coordinate_to_tuple("AA100").unwrap(), (100, 27));
        assert!(coordinate_to_tuple("A0").is_err());
        assert!(coordinate_to_tuple("A1048577").is_err());
        assert!(coordinate_to_tuple("invalid").is_err());
    }

    #[test]
    fn test_tuple_to_coordinate() {
        assert_eq!(tuple_to_coordinate(1, 1).unwrap(), "A1");
        assert_eq!(tuple_to_coordinate(10, 2).unwrap(), "B10");
        assert_eq!(tuple_to_coordinate(100, 27).unwrap(), "AA100");
    }

    #[test]
    fn test_range_boundaries() {
        assert_eq!(range_boundaries("A1").unwrap(), (1, 1, 1, 1));
        assert_eq!(range_boundaries("A1:B10").unwrap(), (1, 1, 10, 2));
        assert_eq!(range_boundaries("B10:A1").unwrap(), (1, 1, 10, 2));
    }

    #[test]
    fn test_sheet_title_validation() {
        assert!(validate_sheet_title("Sheet1").is_ok());
        assert!(validate_sheet_title("").is_err());
        assert!(validate_sheet_title("a".repeat(32).as_str()).is_err());
        assert!(validate_sheet_title("Sheet:1").is_err());
        assert!(validate_sheet_title("Sheet/1").is_err());
        assert!(validate_sheet_title("Sheet?").is_err());
        assert!(validate_sheet_title("'Sheet'").is_err());
    }
}
