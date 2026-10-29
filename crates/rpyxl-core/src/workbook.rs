use crate::coordinate::validate_sheet_title;
use crate::error::{RpyxlError, RpyxlResult};
use crate::worksheet::Worksheet;

#[derive(Debug, Clone, PartialEq)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
    active_index: usize,
    pub data_only: bool,
    pub read_only: bool,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    pub fn new() -> Self {
        let default_sheet = Worksheet::new("Sheet").expect("Default sheet title is valid");
        Self {
            sheets: vec![default_sheet],
            active_index: 0,
            data_only: false,
            read_only: false,
        }
    }

    pub fn with_options(data_only: bool, read_only: bool) -> Self {
        Self {
            sheets: Vec::new(),
            active_index: 0,
            data_only,
            read_only,
        }
    }

    pub fn is_data_only(&self) -> bool {
        self.data_only
    }

    pub fn is_read_only(&self) -> bool {
        self.read_only
    }

    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    pub fn sheetnames(&self) -> Vec<String> {
        self.sheets.iter().map(|s| s.title.clone()).collect()
    }

    pub fn active(&self) -> &Worksheet {
        &self.sheets[self.active_index]
    }

    pub fn active_mut(&mut self) -> &mut Worksheet {
        &mut self.sheets[self.active_index]
    }

    pub fn active_index(&self) -> usize {
        self.active_index
    }

    pub fn set_active_index(&mut self, index: usize) -> RpyxlResult<()> {
        if index >= self.sheets.len() {
            return Err(RpyxlError::SheetIndexOutOfBounds(index, self.sheets.len()));
        }
        self.active_index = index;
        Ok(())
    }

    pub fn set_active_by_title(&mut self, title: &str) -> RpyxlResult<()> {
        if let Some(pos) = self.sheets.iter().position(|s| s.title == title) {
            self.active_index = pos;
            Ok(())
        } else {
            Err(RpyxlError::SheetNotFound(title.to_string()))
        }
    }

    pub fn get_sheet(&self, title: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.title == title)
    }

    pub fn get_sheet_mut(&mut self, title: &str) -> Option<&mut Worksheet> {
        self.sheets.iter_mut().find(|s| s.title == title)
    }

    pub fn get_sheet_by_index(&self, index: usize) -> Option<&Worksheet> {
        self.sheets.get(index)
    }

    pub fn get_sheet_by_index_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.sheets.get_mut(index)
    }

    pub fn contains_sheet(&self, title: &str) -> bool {
        self.sheets.iter().any(|s| s.title == title)
    }

    pub fn create_sheet(
        &mut self,
        title: Option<&str>,
        index: Option<usize>,
    ) -> RpyxlResult<&mut Worksheet> {
        let sheet_title = match title {
            Some(t) => {
                validate_sheet_title(t)?;
                if self.sheets.iter().any(|s| s.title.eq_ignore_ascii_case(t)) {
                    return Err(RpyxlError::DuplicateSheetTitle(t.to_string()));
                }
                t.to_string()
            }
            None => self.generate_unique_sheet_name(),
        };

        let new_ws = Worksheet::new(sheet_title)?;
        let insert_idx = index.unwrap_or(self.sheets.len()).min(self.sheets.len());

        self.sheets.insert(insert_idx, new_ws);
        Ok(&mut self.sheets[insert_idx])
    }

    pub fn remove_sheet(&mut self, title: &str) -> RpyxlResult<()> {
        if self.sheets.len() <= 1 {
            return Err(RpyxlError::CannotRemoveLastSheet);
        }

        let pos = self
            .sheets
            .iter()
            .position(|s| s.title == title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.to_string()))?;

        self.sheets.remove(pos);

        if self.active_index >= self.sheets.len() {
            self.active_index = self.sheets.len() - 1;
        } else if pos < self.active_index {
            self.active_index -= 1;
        }

        Ok(())
    }

    pub fn rename_sheet(&mut self, old_title: &str, new_title: &str) -> RpyxlResult<()> {
        if old_title == new_title {
            return Ok(());
        }

        validate_sheet_title(new_title)?;

        if self
            .sheets
            .iter()
            .any(|s| s.title != old_title && s.title.eq_ignore_ascii_case(new_title))
        {
            return Err(RpyxlError::DuplicateSheetTitle(new_title.to_string()));
        }

        let ws = self
            .sheets
            .iter_mut()
            .find(|s| s.title == old_title)
            .ok_or_else(|| RpyxlError::SheetNotFound(old_title.to_string()))?;

        ws.set_title(new_title)?;
        Ok(())
    }

    fn generate_unique_sheet_name(&self) -> String {
        let mut idx = 1;
        loop {
            let candidate = format!("Sheet{}", idx);
            if !self
                .sheets
                .iter()
                .any(|s| s.title.eq_ignore_ascii_case(&candidate))
            {
                return candidate;
            }
            idx += 1;
        }
    }

    pub fn index_of(&self, title: &str) -> Option<usize> {
        self.sheets.iter().position(|s| s.title == title)
    }

    pub fn copy_worksheet(&mut self, from_title: &str) -> RpyxlResult<&mut Worksheet> {
        let source_ws = self
            .get_sheet(from_title)
            .ok_or_else(|| RpyxlError::SheetNotFound(from_title.to_string()))?
            .clone();

        let copy_title = self.generate_unique_copy_name(&source_ws.title);
        let mut new_ws = source_ws;
        new_ws.set_title(&copy_title)?;

        self.sheets.push(new_ws);
        Ok(self.sheets.last_mut().expect("just pushed"))
    }

    fn generate_unique_copy_name(&self, base_title: &str) -> String {
        let first_candidate = format!("{} Copy", base_title);
        if !self
            .sheets
            .iter()
            .any(|s| s.title.eq_ignore_ascii_case(&first_candidate))
        {
            return first_candidate;
        }

        let mut idx = 1;
        loop {
            let candidate = format!("{} Copy{}", base_title, idx);
            if !self
                .sheets
                .iter()
                .any(|s| s.title.eq_ignore_ascii_case(&candidate))
            {
                return candidate;
            }
            idx += 1;
        }
    }

    pub fn sheets(&self) -> &[Worksheet] {
        &self.sheets
    }

    pub fn sheets_mut(&mut self) -> &mut [Worksheet] {
        &mut self.sheets
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;

    #[test]
    fn test_workbook_default() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.sheetnames(), vec!["Sheet"]);
        assert_eq!(wb.active().title, "Sheet");
    }

    #[test]
    fn test_create_and_remove_sheets() {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("Invoices"), None).unwrap();
        assert_eq!(wb.sheetnames(), vec!["Sheet", "Invoices"]);

        wb.create_sheet(None, Some(0)).unwrap();
        assert_eq!(wb.sheetnames(), vec!["Sheet1", "Sheet", "Invoices"]);

        wb.remove_sheet("Sheet").unwrap();
        assert_eq!(wb.sheetnames(), vec!["Sheet1", "Invoices"]);

        wb.remove_sheet("Invoices").unwrap();
        assert_eq!(wb.sheetnames(), vec!["Sheet1"]);

        // Cannot remove last sheet
        assert!(wb.remove_sheet("Sheet1").is_err());
    }

    #[test]
    fn test_rename_sheet() {
        let mut wb = Workbook::new();
        wb.rename_sheet("Sheet", "Summary").unwrap();
        assert_eq!(wb.sheetnames(), vec!["Summary"]);
        assert_eq!(wb.active().title, "Summary");
    }

    #[test]
    fn test_duplicate_sheet_name() {
        let mut wb = Workbook::new();
        assert!(wb.create_sheet(Some("Sheet"), None).is_err());
        assert!(wb.create_sheet(Some("sheet"), None).is_err());
    }

    #[test]
    fn test_copy_worksheet_and_index() {
        let mut wb = Workbook::new();
        wb.active_mut()
            .set_cell_value(1, 1, CellValue::String("hello".to_string()))
            .unwrap();

        assert_eq!(wb.index_of("Sheet"), Some(0));
        assert_eq!(wb.index_of("NonExistent"), None);

        let copied = wb.copy_worksheet("Sheet").unwrap();
        assert_eq!(copied.title, "Sheet Copy");
        assert_eq!(
            copied.get_cell_value(1, 1),
            CellValue::String("hello".to_string())
        );

        let copied2 = wb.copy_worksheet("Sheet").unwrap();
        assert_eq!(copied2.title, "Sheet Copy1");

        assert_eq!(wb.sheetnames(), vec!["Sheet", "Sheet Copy", "Sheet Copy1"]);
        assert_eq!(wb.index_of("Sheet Copy"), Some(1));
        assert_eq!(wb.index_of("Sheet Copy1"), Some(2));
    }
}
