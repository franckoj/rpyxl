use std::collections::BTreeMap;

use crate::cell::{Cell, CellValue};
use crate::coordinate::{
    range_boundaries, tuple_to_coordinate, validate_sheet_title, MAX_COLUMN, MAX_ROW, MIN_COLUMN,
    MIN_ROW,
};
use crate::error::{RpyxlError, RpyxlResult};

#[derive(Debug, Clone, PartialEq)]
pub struct Worksheet {
    pub title: String,
    cells: BTreeMap<(u32, u32), Cell>,
    cached_min_row: Option<u32>,
    cached_max_row: Option<u32>,
    cached_min_col: Option<u32>,
    cached_max_col: Option<u32>,
}

impl Worksheet {
    pub fn new(title: impl Into<String>) -> RpyxlResult<Self> {
        let title_str = title.into();
        validate_sheet_title(&title_str)?;
        Ok(Self {
            title: title_str,
            cells: BTreeMap::new(),
            cached_min_row: None,
            cached_max_row: None,
            cached_min_col: None,
            cached_max_col: None,
        })
    }

    pub fn set_title(&mut self, new_title: &str) -> RpyxlResult<()> {
        validate_sheet_title(new_title)?;
        self.title = new_title.to_string();
        Ok(())
    }

    pub fn min_row(&self) -> u32 {
        self.cached_min_row.unwrap_or(MIN_ROW)
    }

    pub fn max_row(&self) -> u32 {
        self.cached_max_row.unwrap_or(MIN_ROW)
    }

    pub fn min_column(&self) -> u32 {
        self.cached_min_col.unwrap_or(MIN_COLUMN)
    }

    pub fn max_column(&self) -> u32 {
        self.cached_max_col.unwrap_or(MIN_COLUMN)
    }

    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    pub fn cells(&self) -> &BTreeMap<(u32, u32), Cell> {
        &self.cells
    }

    pub fn dimensions(&self) -> RpyxlResult<String> {
        let min_r = self.min_row();
        let max_r = self.max_row();
        let min_c = self.min_column();
        let max_c = self.max_column();

        let top_left = tuple_to_coordinate(min_r, min_c)?;
        let bottom_right = tuple_to_coordinate(max_r, max_c)?;

        if top_left == bottom_right {
            Ok(top_left)
        } else {
            Ok(format!("{}:{}", top_left, bottom_right))
        }
    }

    pub fn get_cell(&self, row: u32, column: u32) -> Option<&Cell> {
        self.cells.get(&(row, column))
    }

    pub fn get_cell_value(&self, row: u32, column: u32) -> CellValue {
        self.cells
            .get(&(row, column))
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty)
    }

    pub fn get_resolved_cell_value(&self, row: u32, column: u32, data_only: bool) -> CellValue {
        match self.cells.get(&(row, column)) {
            Some(c) => match &c.value {
                CellValue::Formula {
                    formula,
                    cached_value,
                } => {
                    if data_only {
                        cached_value
                            .as_ref()
                            .map(|v| *v.clone())
                            .unwrap_or(CellValue::Empty)
                    } else {
                        CellValue::String(formula.clone())
                    }
                }
                other => other.clone(),
            },
            None => CellValue::Empty,
        }
    }

    pub fn get_or_create_cell(&mut self, row: u32, column: u32) -> RpyxlResult<&mut Cell> {
        if !(MIN_ROW..=MAX_ROW).contains(&row) {
            return Err(RpyxlError::RowOutOfBounds(row));
        }
        if !(MIN_COLUMN..=MAX_COLUMN).contains(&column) {
            return Err(RpyxlError::ColumnOutOfBounds(column));
        }

        self.update_bounds(row, column);
        Ok(self
            .cells
            .entry((row, column))
            .or_insert_with(|| Cell::new(row, column, CellValue::Empty)))
    }

    pub fn set_cell_value(
        &mut self,
        row: u32,
        column: u32,
        value: CellValue,
    ) -> RpyxlResult<&mut Cell> {
        if !(MIN_ROW..=MAX_ROW).contains(&row) {
            return Err(RpyxlError::RowOutOfBounds(row));
        }
        if !(MIN_COLUMN..=MAX_COLUMN).contains(&column) {
            return Err(RpyxlError::ColumnOutOfBounds(column));
        }

        self.update_bounds(row, column);
        let cell = self
            .cells
            .entry((row, column))
            .or_insert_with(|| Cell::new(row, column, CellValue::Empty));
        if cell
            .number_format
            .as_deref()
            .map_or(true, |code| code == "General")
        {
            let temporal_format = match &value {
                CellValue::Date(_) => Some("yyyy-mm-dd"),
                CellValue::DateTime(text) if !text.contains('-') => Some("h:mm:ss.000"),
                CellValue::DateTime(_) => Some("yyyy-mm-dd h:mm:ss.000"),
                _ => None,
            };
            if let Some(code) = temporal_format {
                cell.number_format = Some(code.to_string());
            }
        }
        cell.value = value;
        Ok(cell)
    }

    pub fn get_cell_mut(&mut self, row: u32, column: u32) -> Option<&mut Cell> {
        self.cells.get_mut(&(row, column))
    }

    pub fn set_cell_number_format(
        &mut self,
        row: u32,
        column: u32,
        format: Option<String>,
    ) -> RpyxlResult<()> {
        let cell = self.get_or_create_cell(row, column)?;
        cell.number_format = format;
        Ok(())
    }

    pub fn set_cell_comment(
        &mut self,
        row: u32,
        column: u32,
        comment: Option<String>,
    ) -> RpyxlResult<()> {
        let cell = self.get_or_create_cell(row, column)?;
        cell.comment = comment;
        Ok(())
    }

    pub fn set_cell_hyperlink(
        &mut self,
        row: u32,
        column: u32,
        hyperlink: Option<String>,
    ) -> RpyxlResult<()> {
        let cell = self.get_or_create_cell(row, column)?;
        cell.hyperlink = hyperlink;
        Ok(())
    }

    pub fn insert_rows(&mut self, idx: u32, amount: u32) -> RpyxlResult<()> {
        self.edit_axis(idx, amount, true, true)
    }

    pub fn delete_rows(&mut self, idx: u32, amount: u32) -> RpyxlResult<()> {
        self.edit_axis(idx, amount, true, false)
    }

    pub fn insert_cols(&mut self, idx: u32, amount: u32) -> RpyxlResult<()> {
        self.edit_axis(idx, amount, false, true)
    }

    pub fn delete_cols(&mut self, idx: u32, amount: u32) -> RpyxlResult<()> {
        self.edit_axis(idx, amount, false, false)
    }

    fn edit_axis(&mut self, idx: u32, amount: u32, rows: bool, insert: bool) -> RpyxlResult<()> {
        let limit = if rows { MAX_ROW } else { MAX_COLUMN };
        if idx == 0 || idx > limit {
            return Err(if rows {
                RpyxlError::RowOutOfBounds(idx)
            } else {
                RpyxlError::ColumnOutOfBounds(idx)
            });
        }
        if amount == 0 {
            return Ok(());
        }
        // Use a wider endpoint so deletes crossing the sheet edge simply clamp.
        let end = (u64::from(idx) + u64::from(amount)).min(u64::from(limit) + 1);
        if insert {
            for &(row, col) in self.cells.keys() {
                let pos = if rows { row } else { col };
                if pos >= idx && u64::from(pos) + u64::from(amount) > u64::from(limit) {
                    return Err(RpyxlError::ValueError(
                        "Insertion would move cells beyond worksheet bounds".to_string(),
                    ));
                }
            }
        }
        let old_cells = std::mem::take(&mut self.cells);
        for ((mut row, mut col), mut cell) in old_cells {
            let pos = if rows { row } else { col };
            let target = if pos < idx {
                pos
            } else if insert {
                pos + amount // All destinations were validated before taking the map.
            } else if u64::from(pos) < end {
                continue;
            } else {
                pos - amount
            };
            if rows {
                row = target;
            } else {
                col = target;
            }
            cell.row = row;
            cell.column = col;
            self.cells.insert((row, col), cell);
        }
        self.recalculate_bounds();
        Ok(())
    }

    /// Move a rectangular range without translating formulas or updating references.
    /// Source holes clear destination cells, but never allocate a dense blank grid.
    pub fn move_range(&mut self, cell_range: &str, rows: i64, cols: i64) -> RpyxlResult<()> {
        let (min_row, min_col, max_row, max_col) = range_boundaries(cell_range)?;
        let shift = |pos: u32, delta: i64, limit: u32| -> RpyxlResult<u32> {
            let target = i64::from(pos)
                .checked_add(delta)
                .ok_or_else(|| RpyxlError::ValueError("Range offset overflow".to_string()))?;
            if target < 1 || target > i64::from(limit) {
                return Err(RpyxlError::ValueError(
                    "Moved range would be outside worksheet bounds".to_string(),
                ));
            }
            Ok(target as u32)
        };
        // Validate the full rectangle, including holes, before any mutation.
        let dest_min_row = shift(min_row, rows, MAX_ROW)?;
        let dest_max_row = shift(max_row, rows, MAX_ROW)?;
        let dest_min_col = shift(min_col, cols, MAX_COLUMN)?;
        let dest_max_col = shift(max_col, cols, MAX_COLUMN)?;
        if rows == 0 && cols == 0 {
            return Ok(());
        }
        let old_cells = std::mem::take(&mut self.cells);
        let mut moved = Vec::new();
        for ((row, col), mut cell) in old_cells {
            if (min_row..=max_row).contains(&row) && (min_col..=max_col).contains(&col) {
                let target_row = (i64::from(row) + rows) as u32;
                let target_col = (i64::from(col) + cols) as u32;
                cell.row = target_row;
                cell.column = target_col;
                moved.push(((target_row, target_col), cell));
            } else if !(dest_min_row..=dest_max_row).contains(&row)
                || !(dest_min_col..=dest_max_col).contains(&col)
            {
                self.cells.insert((row, col), cell);
            }
        }
        // Insert only after clearing both rectangles so overlapping moves are safe.
        self.cells.extend(moved);
        self.recalculate_bounds();
        Ok(())
    }

    fn recalculate_bounds(&mut self) {
        self.cached_min_row = None;
        self.cached_max_row = None;
        self.cached_min_col = None;
        self.cached_max_col = None;
        for &(row, col) in self.cells.keys() {
            self.cached_min_row = Some(self.cached_min_row.map_or(row, |m| m.min(row)));
            self.cached_max_row = Some(self.cached_max_row.map_or(row, |m| m.max(row)));
            self.cached_min_col = Some(self.cached_min_col.map_or(col, |m| m.min(col)));
            self.cached_max_col = Some(self.cached_max_col.map_or(col, |m| m.max(col)));
        }
    }

    fn update_bounds(&mut self, row: u32, column: u32) {
        self.cached_min_row = Some(self.cached_min_row.map_or(row, |m| m.min(row)));
        self.cached_max_row = Some(self.cached_max_row.map_or(row, |m| m.max(row)));
        self.cached_min_col = Some(self.cached_min_col.map_or(column, |m| m.min(column)));
        self.cached_max_col = Some(self.cached_max_col.map_or(column, |m| m.max(column)));
    }

    pub fn append(&mut self, values: Vec<CellValue>) -> RpyxlResult<u32> {
        if values.is_empty() {
            return Ok(self.max_row());
        }

        let target_row = if self.is_empty() {
            MIN_ROW
        } else {
            self.max_row() + 1
        };

        if target_row > MAX_ROW {
            return Err(RpyxlError::RowOutOfBounds(target_row));
        }

        for (idx, val) in values.into_iter().enumerate() {
            let col = (idx as u32) + 1;
            if col > MAX_COLUMN {
                return Err(RpyxlError::ColumnOutOfBounds(col));
            }
            self.set_cell_value(target_row, col, val)?;
        }

        Ok(target_row)
    }

    pub fn iter_rows_values(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> RpyxlResult<Vec<Vec<CellValue>>> {
        let r_start = min_row.unwrap_or_else(|| self.min_row());
        let r_end = max_row.unwrap_or_else(|| self.max_row());
        let c_start = min_col.unwrap_or_else(|| self.min_column());
        let c_end = max_col.unwrap_or_else(|| self.max_column());

        if r_start > r_end || c_start > c_end {
            return Ok(Vec::new());
        }

        let mut rows = Vec::with_capacity((r_end - r_start + 1) as usize);
        for r in r_start..=r_end {
            let mut row_vals = Vec::with_capacity((c_end - c_start + 1) as usize);
            for c in c_start..=c_end {
                row_vals.push(self.get_cell_value(r, c));
            }
            rows.push(row_vals);
        }
        Ok(rows)
    }

    pub fn iter_cols_values(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> RpyxlResult<Vec<Vec<CellValue>>> {
        let r_start = min_row.unwrap_or_else(|| self.min_row());
        let r_end = max_row.unwrap_or_else(|| self.max_row());
        let c_start = min_col.unwrap_or_else(|| self.min_column());
        let c_end = max_col.unwrap_or_else(|| self.max_column());

        if r_start > r_end || c_start > c_end {
            return Ok(Vec::new());
        }

        let mut cols = Vec::with_capacity((c_end - c_start + 1) as usize);
        for c in c_start..=c_end {
            let mut col_vals = Vec::with_capacity((r_end - r_start + 1) as usize);
            for r in r_start..=r_end {
                col_vals.push(self.get_cell_value(r, c));
            }
            cols.push(col_vals);
        }
        Ok(cols)
    }

    pub fn resolve_bounds(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> (u32, u32, u32, u32) {
        let r_start = min_row.unwrap_or_else(|| self.min_row());
        let r_end = max_row.unwrap_or_else(|| self.max_row());
        let c_start = min_col.unwrap_or_else(|| self.min_column());
        let c_end = max_col.unwrap_or_else(|| self.max_column());
        (r_start, r_end, c_start, c_end)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rich_cell(ws: &mut Worksheet, row: u32, col: u32) -> Cell {
        let cell = ws
            .set_cell_value(
                row,
                col,
                CellValue::formula_with_cached("=A1+1", CellValue::Int(7)),
            )
            .unwrap();
        cell.number_format = Some("0.00".to_string());
        cell.comment = Some("Keep me".to_string());
        cell.hyperlink = Some("https://example.com".to_string());
        cell.clone()
    }

    #[test]
    fn test_axis_edits_preserve_full_cells_and_sparse_bounds() {
        for rows in [true, false] {
            let mut ws = Worksheet::new("Sheet").unwrap();
            ws.set_cell_value(1, 1, CellValue::Int(1)).unwrap();
            let mut expected = rich_cell(&mut ws, 4, 4);
            let original = ws.clone();
            if rows {
                ws.insert_rows(2, 3).unwrap();
                expected.row = 7;
            } else {
                ws.insert_cols(2, 3).unwrap();
                expected.column = 7;
            }
            assert_eq!(ws.cell_count(), 2);
            assert_eq!(ws.get_cell(expected.row, expected.column), Some(&expected));
            assert!(ws.get_cell(4, 4).is_none());
            assert_eq!(ws.max_row(), if rows { 7 } else { 4 });
            assert_eq!(ws.max_column(), if rows { 4 } else { 7 });
            if rows {
                ws.delete_rows(2, 3).unwrap();
            } else {
                ws.delete_cols(2, 3).unwrap();
            }
            assert_eq!(ws, original);
        }
    }

    #[test]
    fn test_axis_delete_removes_band_and_clamps_at_sheet_end() {
        for rows in [true, false] {
            let mut ws = Worksheet::new("Sheet").unwrap();
            ws.set_cell_value(2, 2, CellValue::Int(2)).unwrap();
            let mut expected = rich_cell(&mut ws, 5, 5);
            if rows {
                ws.delete_rows(2, 2).unwrap();
                expected.row = 3;
            } else {
                ws.delete_cols(2, 2).unwrap();
                expected.column = 3;
            }
            assert_eq!(ws.cell_count(), 1);
            assert_eq!(ws.get_cell(expected.row, expected.column), Some(&expected));
            if rows {
                ws.delete_rows(3, u32::MAX).unwrap();
            } else {
                ws.delete_cols(3, u32::MAX).unwrap();
            }
            assert!(ws.is_empty());
            assert_eq!(ws.dimensions().unwrap(), "A1");
        }
    }

    #[test]
    fn test_axis_failures_are_atomic_and_zero_amount_is_noop() {
        let mut ws = Worksheet::new("Sheet").unwrap();
        rich_cell(&mut ws, MAX_ROW, MAX_COLUMN);
        let original = ws.clone();
        for rows in [true, false] {
            let limit = if rows { MAX_ROW } else { MAX_COLUMN };
            for idx in [0, limit + 1, u32::MAX] {
                assert!(ws.edit_axis(idx, 1, rows, true).is_err());
                assert_eq!(ws, original);
                assert!(ws.edit_axis(idx, 1, rows, false).is_err());
                assert_eq!(ws, original);
            }
            for amount in [1, u32::MAX] {
                assert!(ws.edit_axis(1, amount, rows, true).is_err());
                assert_eq!(ws, original);
            }
            ws.edit_axis(1, 0, rows, true).unwrap();
            ws.edit_axis(1, 0, rows, false).unwrap();
            assert_eq!(ws, original);
        }
    }

    #[test]
    fn test_move_range_overlap_in_both_directions_preserves_metadata() {
        for (rows, cols) in [(1, 1), (-1, -1), (1, 0), (0, -1)] {
            let mut ws = Worksheet::new("Sheet").unwrap();
            let mut first = rich_cell(&mut ws, 3, 3);
            let mut last = rich_cell(&mut ws, 4, 4);
            ws.set_cell_value(1, 1, CellValue::formula("=C3")).unwrap();
            ws.move_range("C3:D4", rows, cols).unwrap();
            first.row = (3 + rows) as u32;
            first.column = (3 + cols) as u32;
            last.row = (4 + rows) as u32;
            last.column = (4 + cols) as u32;
            assert_eq!(ws.get_cell(first.row, first.column), Some(&first));
            assert_eq!(ws.get_cell(last.row, last.column), Some(&last));
            assert_eq!(ws.get_cell_value(1, 1), CellValue::formula("=C3"));
            assert_eq!(ws.cell_count(), 3);
            for row in 3..=4 {
                for col in 3..=4 {
                    if (row, col) != (first.row, first.column)
                        && (row, col) != (last.row, last.column)
                    {
                        assert!(ws.get_cell(row, col).is_none());
                    }
                }
            }
        }
    }

    #[test]
    fn test_move_range_holes_overwrite_and_vacated_cells_clear() {
        let mut ws = Worksheet::new("Sheet").unwrap();
        let mut expected = rich_cell(&mut ws, 1, 1);
        rich_cell(&mut ws, 3, 3);
        rich_cell(&mut ws, 3, 4);
        rich_cell(&mut ws, 4, 4);
        ws.move_range("A1:B2", 2, 2).unwrap();
        expected.row = 3;
        expected.column = 3;
        assert_eq!(ws.get_cell(3, 3), Some(&expected));
        assert_eq!(ws.cell_count(), 1);
        assert_eq!(ws.dimensions().unwrap(), "C3");
        ws.move_range("A1:B2", 2, 2).unwrap();
        assert!(ws.is_empty());
        assert_eq!(ws.dimensions().unwrap(), "A1");
    }

    #[test]
    fn test_move_range_validates_full_rectangle_atomically() {
        let mut ws = Worksheet::new("Sheet").unwrap();
        rich_cell(&mut ws, 2, 2);
        let original = ws.clone();
        for (range, rows, cols) in [
            ("A0", 1, 1),
            ("bad", 0, 0),
            ("A1:B2", -1, 0),
            ("A1:B2", 0, -1),
            ("A1:XFD2", 0, 1),
            ("A1:B1048576", 1, 0),
            ("B2", i64::MAX, 0),
            ("B2", 0, i64::MIN),
        ] {
            assert!(ws.move_range(range, rows, cols).is_err());
            assert_eq!(ws, original);
        }
        ws.move_range("B2", 0, 0).unwrap();
        assert_eq!(ws, original);
    }

    #[test]
    fn test_structural_edits_on_empty_and_far_sparse_sheets() {
        let mut ws = Worksheet::new("Sheet").unwrap();
        ws.insert_rows(1, 1).unwrap();
        ws.insert_cols(1, 1).unwrap();
        ws.delete_rows(1, u32::MAX).unwrap();
        ws.delete_cols(1, u32::MAX).unwrap();
        ws.move_range("A1:XFD1048576", 0, 0).unwrap();
        assert!(ws.is_empty());
        rich_cell(&mut ws, MAX_ROW, MAX_COLUMN);
        ws.move_range("A1:XFD1048576", 0, 0).unwrap();
        assert_eq!(ws.cell_count(), 1);
        ws.move_range("B2:XFD1048576", -1, -1).unwrap();
        assert_eq!(ws.cell_count(), 1);
        assert!(ws.get_cell(MAX_ROW - 1, MAX_COLUMN - 1).is_some());
        ws.delete_rows(1, MAX_ROW - 2).unwrap();
        ws.delete_cols(1, MAX_COLUMN - 2).unwrap();
        assert_eq!(ws.dimensions().unwrap(), "A1");
        assert_eq!(ws.cell_count(), 1);
    }

    #[test]
    fn test_worksheet_creation() {
        let ws = Worksheet::new("Sheet1").unwrap();
        assert_eq!(ws.title, "Sheet1");
        assert_eq!(ws.min_row(), 1);
        assert_eq!(ws.max_row(), 1);
        assert_eq!(ws.min_column(), 1);
        assert_eq!(ws.max_column(), 1);
        assert_eq!(ws.dimensions().unwrap(), "A1");
        assert!(ws.is_empty());
    }

    #[test]
    fn test_cell_assignment_and_dimensions() {
        let mut ws = Worksheet::new("Data").unwrap();
        ws.set_cell_value(1, 1, CellValue::String("Header".to_string()))
            .unwrap();
        ws.set_cell_value(10, 5, CellValue::Int(100)).unwrap();

        assert_eq!(ws.min_row(), 1);
        assert_eq!(ws.max_row(), 10);
        assert_eq!(ws.min_column(), 1);
        assert_eq!(ws.max_column(), 5);
        assert_eq!(ws.dimensions().unwrap(), "A1:E10");

        assert_eq!(
            ws.get_cell_value(1, 1),
            CellValue::String("Header".to_string())
        );
        assert_eq!(ws.get_cell_value(10, 5), CellValue::Int(100));
        assert_eq!(ws.get_cell_value(2, 2), CellValue::Empty);
    }

    #[test]
    fn test_append() {
        let mut ws = Worksheet::new("Sheet").unwrap();
        ws.append(vec![
            CellValue::String("INV001".to_string()),
            CellValue::Float(1500.50),
        ])
        .unwrap();

        assert_eq!(ws.max_row(), 1);
        assert_eq!(ws.max_column(), 2);
        assert_eq!(
            ws.get_cell_value(1, 1),
            CellValue::String("INV001".to_string())
        );
        assert_eq!(ws.get_cell_value(1, 2), CellValue::Float(1500.50));

        ws.append(vec![
            CellValue::String("INV002".to_string()),
            CellValue::Float(2500.0),
        ])
        .unwrap();

        assert_eq!(ws.max_row(), 2);
        assert_eq!(ws.max_column(), 2);
        assert_eq!(
            ws.get_cell_value(2, 1),
            CellValue::String("INV002".to_string())
        );
    }

    #[test]
    fn test_iter_rows_values() {
        let mut ws = Worksheet::new("Sheet").unwrap();
        ws.append(vec![CellValue::Int(1), CellValue::Int(2)])
            .unwrap();
        ws.append(vec![CellValue::Int(3), CellValue::Int(4)])
            .unwrap();

        let rows = ws.iter_rows_values(None, None, None, None).unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0], vec![CellValue::Int(1), CellValue::Int(2)]);
        assert_eq!(rows[1], vec![CellValue::Int(3), CellValue::Int(4)]);
    }
}
