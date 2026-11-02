use std::sync::Arc;

/// Original worksheet framing, excluding the generated dimension and cell data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetTemplate {
    pub root_start: String,
    pub before_data: String,
    pub data_start: String,
    pub after_data: String,
}

/// A compressed source package whose untouched members can be copied verbatim.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PackageTemplate {
    pub archive: Arc<[u8]>,
    pub workbook_path: String,
    pub styles_path: String,
    pub styles_xml: String,
    pub sheet_parts: Vec<String>,
    pub sheet_names: Vec<String>,
    pub active_index: usize,
    pub date_1904: bool,
}
