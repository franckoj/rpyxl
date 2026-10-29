use std::collections::BTreeMap;

use crate::cell::{Cell, CellValue};
use crate::coordinate::{
    tuple_to_coordinate, validate_sheet_title, MAX_COLUMN, MAX_ROW, MIN_COLUMN, MIN_ROW,
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
