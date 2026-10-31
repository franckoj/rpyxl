use std::collections::HashMap;
use std::sync::Arc;

use crate::cell::RichStyle;
use crate::coordinate::validate_sheet_title;
use crate::error::{RpyxlError, RpyxlResult};
use crate::package::PackageTemplate;
use crate::worksheet::Worksheet;

#[derive(Debug, Clone)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
    active_index: usize,
    pub data_only: bool,
    pub read_only: bool,
    /// Whether numeric dates use the 1904-01-01 epoch instead of Excel's 1900 epoch.
    pub date_1904: bool,
    pub package_template: Option<Arc<PackageTemplate>>,
    pub source_styles_xml: Option<String>,
    source_styles: Vec<Arc<RichStyle>>,
    source_style_ids: HashMap<usize, usize>,
}

impl PartialEq for Workbook {
    fn eq(&self, other: &Self) -> bool {
        // source_style_ids is a pointer-keyed lookup cache, not workbook content.
        self.sheets == other.sheets
            && self.active_index == other.active_index
            && self.data_only == other.data_only
            && self.read_only == other.read_only
            && self.date_1904 == other.date_1904
            && self.package_template == other.package_template
            && self.source_styles_xml == other.source_styles_xml
            && self.source_styles == other.source_styles
    }
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
            date_1904: false,
            package_template: None,
            source_styles_xml: None,
            source_styles: Vec::new(),
            source_style_ids: HashMap::new(),
        }
    }

    pub fn with_options(data_only: bool, read_only: bool) -> Self {
        Self {
            sheets: Vec::new(),
            active_index: 0,
            data_only,
            read_only,
            date_1904: false,
            package_template: None,
            source_styles_xml: None,
            source_styles: Vec::new(),
            source_style_ids: HashMap::new(),
        }
    }

    pub fn set_source_styles(&mut self, styles: Vec<Arc<RichStyle>>) {
        self.source_style_ids = styles
            .iter()
            .enumerate()
            .map(|(index, style)| (Arc::as_ptr(style) as usize, index))
            .collect();
        self.source_styles = styles;
    }

    pub fn source_styles(&self) -> &[Arc<RichStyle>] {
        &self.source_styles
    }

    pub fn source_style_index(&self, style: &Arc<RichStyle>) -> Option<usize> {
        self.source_style_ids
            .get(&(Arc::as_ptr(style) as usize))
            .copied()
    }

    pub fn can_preserve_package(&self) -> bool {
        self.package_template.as_ref().is_some_and(|source| {
            self.sheetnames() == source.sheet_names
                && self.active_index == source.active_index
                && self.date_1904 == source.date_1904
                && self.sheets.len() == source.sheet_parts.len()
                && self.sheets.iter().all(|sheet| sheet.xml_template.is_some())
        })
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
    fn test_equality_ignores_source_style_pointer_cache() {
        let mut left = Workbook::new();
        let mut right = Workbook::new();
        for book in [&mut left, &mut right] {
            let style = Arc::new(RichStyle {
                font: Some("<font><b/></font>".into()),
                ..RichStyle::default()
            });
            book.set_source_styles(vec![Arc::clone(&style)]);
            book.source_styles_xml = Some("<styleSheet/>".into());
            book.active_mut()
                .set_cell_value(1, 1, CellValue::Int(7))
                .unwrap();
            book.active_mut().get_cell_mut(1, 1).unwrap().rich_style = Some(style);
        }
        assert_ne!(left.source_style_ids, right.source_style_ids);
        assert_eq!(left.source_style_index(&left.source_styles()[0]), Some(0));
        assert_eq!(right.source_style_index(&right.source_styles()[0]), Some(0));
        assert_eq!(left.source_style_index(&right.source_styles()[0]), None);
        assert_eq!(left, right);
        assert_eq!(left, left.clone());

        right.source_style_ids.clear();
        assert_eq!(left, right);
    }

    #[test]
    fn test_equality_includes_all_semantic_fields() {
        let mut book = Workbook::new();
        book.create_sheet(Some("Other"), None).unwrap();
        let changes: &[fn(&mut Workbook)] = &[
            |book| book.sheets[0].title = "Changed".into(),
            |book| book.active_index = 1,
            |book| book.data_only = true,
            |book| book.read_only = true,
            |book| book.date_1904 = true,
            |book| {
                book.package_template = Some(Arc::new(PackageTemplate {
                    archive: Arc::from([]),
                    workbook_path: "xl/workbook.xml".into(),
                    styles_path: "xl/styles.xml".into(),
                    styles_xml: "<styleSheet/>".into(),
                    sheet_parts: Vec::new(),
                    sheet_names: Vec::new(),
                    active_index: 0,
                    date_1904: false,
                }));
            },
            |book| book.source_styles_xml = Some("<styleSheet/>".into()),
            |book| book.set_source_styles(vec![Arc::new(RichStyle::default())]),
        ];
        for change in changes {
            let mut changed = book.clone();
            change(&mut changed);
            assert_ne!(book, changed);
        }

        book.set_source_styles(vec![Arc::new(RichStyle::default())]);
        let mut changed = book.clone();
        changed.set_source_styles(vec![Arc::new(RichStyle {
            alignment: Some("<alignment horizontal=\"right\"/>".into()),
            ..RichStyle::default()
        })]);
        assert_ne!(book, changed);
    }

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
