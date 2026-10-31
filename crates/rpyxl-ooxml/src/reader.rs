use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::num::NonZeroU32;
use std::path::Path;
use std::sync::Arc;

use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::Reader;
use quick_xml::{NsReader, Writer};
use zip::ZipArchive;

use rpyxl_core::cell::{Cell, CellValue};
use rpyxl_core::coordinate::{coordinate_to_tuple, range_boundaries};
use rpyxl_core::error::{RpyxlError, RpyxlResult};
use rpyxl_core::package::WorksheetTemplate;
use rpyxl_core::workbook::Workbook;
use rpyxl_core::worksheet::Worksheet;

use crate::shared_strings::{decode_xml_text, SharedStringsTable};
use crate::styles::{utf8_declaration, Styles};

#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
    pub target_path: String,
    pub min_row: Option<u32>,
    pub max_row: Option<u32>,
    pub min_col: Option<u32>,
    pub max_col: Option<u32>,
}

pub struct XlsxPackage<R: Read + Seek> {
    pub archive: ZipArchive<R>,
    pub sheets: Vec<SheetInfo>,
    pub shared_strings: Arc<SharedStringsTable>,
    pub styles: Arc<Styles>,
    pub workbook_path: String,
    pub styles_path: Option<String>,
    pub active_sheet_index: usize,
    pub data_only: bool,
    pub read_only: bool,
}

impl XlsxPackage<BufReader<File>> {
    pub fn open<P: AsRef<Path>>(path: P, read_only: bool, data_only: bool) -> RpyxlResult<Self> {
        let file = File::open(path.as_ref())
            .map_err(|e| RpyxlError::IoError(format!("Failed to open file: {}", e)))?;
        let reader = BufReader::new(file);
        Self::from_reader(reader, read_only, data_only)
    }
}

pub const MAX_UNCOMPRESSED_FILE_SIZE: u64 = 512 * 1024 * 1024; // 512 MB per entry
pub const MAX_TOTAL_UNCOMPRESSED_SIZE: u64 = 1024 * 1024 * 1024; // 1 GB total
pub const MAX_COMPRESSION_RATIO: f64 = 100.0;
pub const MIN_SIZE_FOR_RATIO_CHECK: u64 = 5 * 1024 * 1024; // 5 MB uncompressed
pub const MAX_ZIP_ENTRIES: usize = 10_000;

pub fn validate_zip_archive_security<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> RpyxlResult<()> {
    let entry_count = archive.len();
    if entry_count > MAX_ZIP_ENTRIES {
        return Err(RpyxlError::ZipBombDetected(format!(
            "Archive contains {} entries, exceeding safety threshold of {}",
            entry_count, MAX_ZIP_ENTRIES
        )));
    }

    let mut total_uncompressed: u64 = 0;
    for i in 0..entry_count {
        let file = archive
            .by_index(i)
            .map_err(|e| RpyxlError::IoError(format!("Malformed ZIP entry #{}: {}", i, e)))?;
        let uncompressed = file.size();
        let compressed = file.compressed_size();

        if uncompressed > MAX_UNCOMPRESSED_FILE_SIZE {
            return Err(RpyxlError::ZipBombDetected(format!(
                "Entry '{}' uncompressed size ({} bytes) exceeds threshold of {} bytes",
                file.name(),
                uncompressed,
                MAX_UNCOMPRESSED_FILE_SIZE
            )));
        }

        total_uncompressed = total_uncompressed.saturating_add(uncompressed);
        if total_uncompressed > MAX_TOTAL_UNCOMPRESSED_SIZE {
            return Err(RpyxlError::ZipBombDetected(format!(
                "Total uncompressed archive size exceeds safety threshold of {} bytes",
                MAX_TOTAL_UNCOMPRESSED_SIZE
            )));
        }

        if uncompressed >= MIN_SIZE_FOR_RATIO_CHECK && compressed > 0 {
            let ratio = (uncompressed as f64) / (compressed as f64);
            if ratio > MAX_COMPRESSION_RATIO {
                return Err(RpyxlError::ZipBombDetected(format!(
                    "Entry '{}' compression ratio ({:.1}:1) exceeds maximum safety ratio of {:.1}:1",
                    file.name(),
                    ratio,
                    MAX_COMPRESSION_RATIO
                )));
            }
        }
    }
    Ok(())
}

impl<R: Read + Seek> XlsxPackage<R> {
    pub fn from_reader(reader: R, read_only: bool, data_only: bool) -> RpyxlResult<Self> {
        let mut archive = ZipArchive::new(reader)
            .map_err(|e| RpyxlError::IoError(format!("Invalid XLSX/ZIP archive: {}", e)))?;

        validate_zip_archive_security(&mut archive)?;

        // 1. Resolve workbook.xml path from _rels/.rels
        let workbook_path = resolve_workbook_path(&mut archive)?;

        // 2. Resolve workbook relationships (sheets, sharedStrings, styles)
        let rels_path = get_rels_path(&workbook_path);
        let rels_map = parse_relationships(&mut archive, &rels_path).unwrap_or_default();

        // 3. Parse xl/workbook.xml for sheets list & active sheet
        let (sheet_metas, active_idx, date_1904) =
            parse_workbook_xml(&mut archive, &workbook_path)?;

        // Resolve target paths and peek dimensions for each sheet
        let base_dir = get_base_dir(&workbook_path);
        let mut sheets = Vec::new();
        for meta in sheet_metas {
            let rel_target = rels_map
                .get(&meta.r_id)
                .filter(|rel| !rel.external)
                .map(|rel| rel.target.clone())
                .unwrap_or_else(|| format!("worksheets/sheet{}.xml", meta.sheet_id));
            let full_target = normalize_zip_path(&base_dir, &rel_target);
            let dim = peek_sheet_dimension(&mut archive, &full_target);

            sheets.push(SheetInfo {
                name: meta.name,
                sheet_id: meta.sheet_id,
                r_id: meta.r_id,
                target_path: full_target,
                min_row: dim.map(|d| d.0),
                max_row: dim.map(|d| d.2),
                min_col: dim.map(|d| d.1),
                max_col: dim.map(|d| d.3),
            });
        }

        // 4. Parse sharedStrings.xml if present
        let shared_strings = if let Some(ss_target) = find_rel_target(&rels_map, "sharedStrings") {
            let full_ss_path = normalize_zip_path(&base_dir, &ss_target);
            match archive.by_name(&full_ss_path) {
                Ok(file) => SharedStringsTable::from_reader(file)?,
                Err(_) => SharedStringsTable::default(),
            }
        } else {
            match archive.by_name("xl/sharedStrings.xml") {
                Ok(file) => SharedStringsTable::from_reader(file)?,
                Err(_) => SharedStringsTable::default(),
            }
        };

        // 5. Parse styles.xml if present
        let resolved_styles_path = find_rel_target(&rels_map, "styles")
            .map(|target| normalize_zip_path(&base_dir, &target))
            .unwrap_or_else(|| "xl/styles.xml".to_string());
        let (mut styles, styles_path) = match archive.by_name(&resolved_styles_path) {
            Ok(file) => (Styles::from_reader(file)?, Some(resolved_styles_path)),
            Err(_) => (Styles::default(), None),
        };

        styles.date_1904 = date_1904;

        Ok(Self {
            archive,
            sheets,
            shared_strings: Arc::new(shared_strings),
            styles: Arc::new(styles),
            workbook_path,
            styles_path,
            active_sheet_index: active_idx,
            data_only,
            read_only,
        })
    }

    pub fn get_sheet_info(&self, name: &str) -> Option<&SheetInfo> {
        self.sheets.iter().find(|s| s.name == name)
    }

    /// Read all sheets into an in-memory Workbook
    pub fn to_workbook(&mut self) -> RpyxlResult<Workbook> {
        let mut wb = Workbook::with_options(self.data_only, self.read_only);
        wb.date_1904 = self.styles.date_1904;
        wb.set_source_styles(self.styles.rich_styles().to_vec());
        wb.source_styles_xml = self.styles.source_xml().map(str::to_owned);

        for sheet_info in &self.sheets {
            let mut ws = Worksheet::new(&sheet_info.name)?;

            if let Ok(file) = self.archive.by_name(&sheet_info.target_path) {
                parse_worksheet_into(
                    file,
                    &mut ws,
                    &self.shared_strings,
                    &self.styles,
                    self.data_only,
                )?;
            }

            // Create sheet in workbook
            wb.create_sheet(Some(&sheet_info.name), None)?;
            if let Some(target_ws) = wb.get_sheet_mut(&sheet_info.name) {
                *target_ws = ws;
            }
        }

        // Set active sheet
        if self.active_sheet_index < wb.sheet_count() {
            let _ = wb.set_active_index(self.active_sheet_index);
        }

        Ok(wb)
    }

    /// Streaming iterator: reads row-by-row on the fly without loading all cells into RAM
    pub fn stream_rows(
        &mut self,
        sheet_name: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> RpyxlResult<Vec<Vec<(u32, u32, CellValue)>>> {
        let target_path = self
            .sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .map(|s| s.target_path.clone())
            .ok_or_else(|| RpyxlError::SheetNotFound(sheet_name.to_string()))?;

        let file = self
            .archive
            .by_name(&target_path)
            .map_err(|e| RpyxlError::IoError(format!("Cannot open sheet XML: {}", e)))?;

        parse_worksheet_streaming(
            file,
            &self.shared_strings,
            &self.styles,
            min_row,
            max_row,
            min_col,
            max_col,
            self.data_only,
        )
    }
}

fn peek_sheet_dimension<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    target_path: &str,
) -> Option<(u32, u32, u32, u32)> {
    let file = match archive.by_name(target_path) {
        Ok(f) => f,
        Err(_) => return None,
    };
    let mut xml_reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();
    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"dimension" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"ref" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                if let Ok((min_r, min_c, max_r, max_c)) = range_boundaries(val) {
                                    return Some((min_r, min_c, max_r, max_c));
                                }
                            }
                        }
                    }
                } else if e.local_name().as_ref() == b"sheetData" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

// Helpers for OOXML discovery

fn resolve_workbook_path<R: Read + Seek>(archive: &mut ZipArchive<R>) -> RpyxlResult<String> {
    if let Ok(file) = archive.by_name("_rels/.rels") {
        let mut reader = Reader::from_reader(BufReader::new(file));
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"Relationship" {
                        let mut is_office_doc = false;
                        let mut target = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"Type" {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    if val.ends_with("/officeDocument") {
                                        is_office_doc = true;
                                    }
                                }
                            } else if attr.key.local_name().as_ref() == b"Target" {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    target = Some(val.to_string());
                                }
                            }
                        }
                        if is_office_doc {
                            if let Some(t) = target {
                                return Ok(normalize_zip_path("", &t));
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    // Default fallback
    Ok("xl/workbook.xml".to_string())
}

fn get_rels_path(file_path: &str) -> String {
    let path = Path::new(file_path);
    let parent = path.parent().unwrap_or(Path::new(""));
    let filename = path.file_name().unwrap_or_default().to_string_lossy();
    let rels_filename = format!("{}.rels", filename);
    let rels_dir = parent.join("_rels");
    rels_dir
        .join(rels_filename)
        .to_string_lossy()
        .replace('\\', "/")
}

struct Relationship {
    target: String,
    kind: String,
    external: bool,
}

fn parse_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    rels_path: &str,
) -> RpyxlResult<HashMap<String, Relationship>> {
    let mut map = HashMap::new();
    let file = match archive.by_name(rels_path) {
        Ok(f) => f,
        Err(_) => return Ok(map),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = None;
                    let mut target = None;
                    let mut kind = String::new();
                    let mut external = false;
                    for attr in e.attributes().flatten() {
                        let value =
                            attr.decode_and_unescape_value(reader.decoder())
                                .map_err(|err| {
                                    RpyxlError::IoError(format!("Invalid relationship: {err}"))
                                })?;
                        match attr.key.local_name().as_ref() {
                            b"Id" => id = Some(value.into_owned()),
                            b"Target" => target = Some(value.into_owned()),
                            b"Type" => kind = value.into_owned(),
                            b"TargetMode" => external = value == "External",
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        map.insert(
                            id,
                            Relationship {
                                target,
                                kind,
                                external,
                            },
                        );
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(map)
}

struct RawSheetMeta {
    name: String,
    sheet_id: u32,
    r_id: String,
}

fn parse_workbook_xml<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    workbook_path: &str,
) -> RpyxlResult<(Vec<RawSheetMeta>, usize, bool)> {
    let file = archive
        .by_name(workbook_path)
        .map_err(|e| RpyxlError::IoError(format!("Cannot open {}: {}", workbook_path, e)))?;

    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut active_tab = 0;
    let mut date_1904 = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"sheet" => {
                    let mut name = String::new();
                    let mut sheet_id = 0;
                    let mut r_id = String::new();

                    for attr in e.attributes().flatten() {
                        let key = attr.key.local_name();
                        if key.as_ref() == b"name" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                name = v.to_string();
                            }
                        } else if key.as_ref() == b"sheetId" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                sheet_id = v.parse().unwrap_or(0);
                            }
                        } else if key.as_ref() == b"id" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                r_id = v.to_string();
                            }
                        }
                    }

                    if !name.is_empty() {
                        sheets.push(RawSheetMeta {
                            name,
                            sheet_id,
                            r_id,
                        });
                    }
                }
                b"workbookPr" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"date1904" {
                            date_1904 = matches!(attr.value.as_ref(), b"1" | b"true");
                        }
                    }
                }
                b"workbookView" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"activeTab" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                active_tab = v.parse().unwrap_or(0);
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((sheets, active_tab, date_1904))
}

fn find_rel_target(rels_map: &HashMap<String, Relationship>, keyword: &str) -> Option<String> {
    rels_map
        .values()
        .find(|rel| !rel.external && rel.kind.rsplit('/').next() == Some(keyword))
        .map(|rel| rel.target.clone())
}

fn get_base_dir(file_path: &str) -> String {
    let p = Path::new(file_path);
    p.parent()
        .map(|d| d.to_string_lossy().replace('\\', "/"))
        .unwrap_or_default()
}

fn normalize_zip_path(base_dir: &str, target: &str) -> String {
    let full = if target.starts_with('/') || base_dir.is_empty() {
        target.to_string()
    } else {
        format!("{}/{}", base_dir.trim_end_matches('/'), target)
    };
    let mut parts = Vec::new();
    for part in full.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}

// Worksheet XML parsing

const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

struct WorksheetFraming {
    writers: [Writer<Vec<u8>>; 4],
    phase: usize, // prolog/root, before sheetData, inside sheetData, after sheetData
    depth: usize,
    dimension_depth: Option<usize>,
    seen_root: bool,
    seen_decl: bool,
    closed_root: bool,
    seen_data: bool,
    safe: bool,
}

impl WorksheetFraming {
    fn new() -> Self {
        Self {
            writers: std::array::from_fn(|_| Writer::new(Vec::new())),
            phase: 0,
            depth: 0,
            dimension_depth: None,
            seen_root: false,
            seen_decl: false,
            closed_root: false,
            seen_data: false,
            safe: true,
        }
    }

    fn observe(
        &mut self,
        event: &Event<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) -> RpyxlResult<()> {
        let mut destination =
            (self.phase != 2 && self.dimension_depth.is_none()).then_some(self.phase);
        match event {
            Event::Decl(decl) => {
                self.safe &= !self.seen_root
                    && !self.seen_decl
                    && self.writers[0].get_ref().is_empty()
                    && decl.version().is_ok()
                    && utf8_declaration(decl);
                self.seen_decl = true;
            }
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                self.safe &= decoder.decode(e.as_ref()).is_ok();
                let mut default_ns = None;
                for attr in e.attributes() {
                    match attr {
                        Ok(attr) => match attr.decode_and_unescape_value(decoder) {
                            Ok(value) if attr.key.as_ref() == b"xmlns" => default_ns = Some(value),
                            Ok(_) => {}
                            Err(_) => self.safe = false,
                        },
                        Err(_) => self.safe = false,
                    }
                }
                if self.depth == 0 {
                    self.safe &= !self.seen_root
                        && !empty
                        && e.name().as_ref() == b"worksheet"
                        && default_ns.as_deref() == Some(SPREADSHEET_NS);
                    self.seen_root = true;
                    self.closed_root = empty;
                    destination = Some(0);
                    self.phase = 1;
                } else if self.depth == 1 && e.local_name().as_ref() == b"dimension" {
                    destination = None;
                    if !empty {
                        self.dimension_depth = Some(self.depth + 1);
                    }
                } else if self.depth == 1 && e.local_name().as_ref() == b"sheetData" {
                    self.safe &= !self.seen_data
                        && e.name().as_ref() == b"sheetData"
                        && default_ns.as_deref().is_none_or(|ns| ns == SPREADSHEET_NS);
                    self.seen_data = true;
                    // Expand <sheetData/> so regenerated cells can be inserted inside it.
                    self.writers[2]
                        .write_event(Event::Start(e.borrow()))
                        .map_err(|err| {
                            RpyxlError::IoError(format!("Cannot capture worksheet XML: {err}"))
                        })?;
                    self.phase = if empty { 3 } else { 2 };
                    destination = None;
                }
                if !empty {
                    self.depth += 1;
                }
            }
            Event::End(e) => {
                if self.depth == 0 {
                    self.safe = false;
                } else {
                    if self.dimension_depth == Some(self.depth) {
                        self.dimension_depth = None;
                        destination = None;
                    } else if self.phase == 2
                        && self.depth == 2
                        && e.local_name().as_ref() == b"sheetData"
                    {
                        self.phase = 3;
                        destination = None;
                    }
                    self.depth -= 1;
                    if self.depth == 0 {
                        self.closed_root = true;
                    }
                }
            }
            Event::Text(e) => match e.unescape() {
                Ok(text) if self.depth == 0 && !text.trim().is_empty() => self.safe = false,
                Ok(_) => {}
                Err(_) => self.safe = false,
            },
            Event::DocType(_) => self.safe = false,
            Event::CData(e) => self.safe &= self.depth != 0 && decoder.decode(e.as_ref()).is_ok(),
            Event::Comment(e) => {
                self.safe &= decoder.decode(e.as_ref()).is_ok()
                    && !e.as_ref().windows(2).any(|pair| pair == b"--")
                    && !e.as_ref().ends_with(b"-");
            }
            Event::Eof => destination = None,
            _ => {}
        }
        if let Some(index) = destination {
            self.writers[index]
                .write_event(event.borrow())
                .map_err(|err| {
                    RpyxlError::IoError(format!("Cannot capture worksheet XML: {err}"))
                })?;
        }
        Ok(())
    }

    fn finish(self) -> Option<WorksheetTemplate> {
        if !self.safe || !self.closed_root || self.depth != 0 || !self.seen_data || self.phase != 3
        {
            return None;
        }
        let [root, before, data, after] = self.writers;
        Some(WorksheetTemplate {
            root_start: String::from_utf8(root.into_inner()).ok()?,
            before_data: String::from_utf8(before.into_inner()).ok()?,
            data_start: String::from_utf8(data.into_inner()).ok()?,
            after_data: String::from_utf8(after.into_inner()).ok()?,
        })
    }
}

fn attach_cell_style(cell: &mut Cell, styles: &Styles, index: usize, explicit: bool) {
    cell.rich_style = styles.rich_style(index);
    cell.source_xf_explicit = explicit;
    cell.source_xf_id = u32::try_from(index)
        .ok()
        .and_then(|index| index.checked_add(1))
        .and_then(NonZeroU32::new);
    if let Some(code) = styles
        .number_format(index)
        .filter(|code| !code.eq_ignore_ascii_case("General"))
    {
        cell.number_format = Some(code.to_owned());
    }
}

fn clear_merged_values(ws: &mut Worksheet, ranges: &[(u32, u32, u32, u32)]) {
    for &(min_row, min_col, max_row, max_col) in ranges {
        if min_row > max_row || min_col > max_col {
            continue;
        }
        let stored: Vec<_> = ws
            .cells()
            .range((min_row, 0)..=(max_row, u32::MAX))
            .filter_map(|(&(row, col), _)| {
                (col >= min_col && col <= max_col && (row, col) != (min_row, min_col))
                    .then_some((row, col))
            })
            .collect();
        for (row, col) in stored {
            if let Some(cell) = ws.get_cell_mut(row, col) {
                cell.value = CellValue::Empty;
            }
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum WorksheetElement {
    Worksheet,
    SheetData,
    Row,
    Cell,
    Value,
    Formula,
    InlineString,
    InlineRun,
    InlineText,
    MergeCells,
    MergeCell,
    Other,
}

impl WorksheetElement {
    fn at_path(name: &[u8], spreadsheet: bool, ancestors: &[Self]) -> Self {
        if !spreadsheet {
            return Self::Other;
        }
        match (ancestors.last(), name) {
            (None, b"worksheet") => Self::Worksheet,
            (Some(Self::Worksheet), b"sheetData") => Self::SheetData,
            (Some(Self::Worksheet), b"mergeCells") => Self::MergeCells,
            (Some(Self::MergeCells), b"mergeCell") => Self::MergeCell,
            (Some(Self::SheetData), b"row") => Self::Row,
            (Some(Self::Row), b"c") => Self::Cell,
            (Some(Self::Cell), b"v") => Self::Value,
            (Some(Self::Cell), b"f") => Self::Formula,
            (Some(Self::Cell), b"is") => Self::InlineString,
            (Some(Self::InlineString), b"r") => Self::InlineRun,
            (Some(Self::InlineString | Self::InlineRun), b"t") => Self::InlineText,
            _ => Self::Other,
        }
    }
}

pub fn parse_worksheet_into<R: Read>(
    reader: R,
    ws: &mut Worksheet,
    shared_strings: &SharedStringsTable,
    styles: &Styles,
    data_only: bool,
) -> RpyxlResult<()> {
    let mut xml_reader = NsReader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut ancestors = Vec::new();
    let mut legacy_namespace = false;
    let mut current_coord = String::new();
    let mut current_type = String::new();
    let mut current_style = 0usize;
    let mut current_style_explicit = false;
    let mut formula_text = String::new();
    let mut val_text = String::new();
    let mut framing = WorksheetFraming::new();
    let mut merged_ranges = Vec::new();
    ws.xml_template = None;

    loop {
        let (namespace, event) = xml_reader
            .read_resolved_event_into(&mut buf)
            .map_err(|e| RpyxlError::IoError(format!("Error parsing worksheet XML: {e:?}")))?;
        // Consume the namespace borrow before borrowing the reader's decoder.
        let spreadsheet = match namespace {
            ResolveResult::Bound(ns) => ns.as_ref() == SPREADSHEET_NS.as_bytes(),
            ResolveResult::Unbound => {
                if ancestors.is_empty() && matches!(event, Event::Start(_) | Event::Empty(_)) {
                    legacy_namespace = true;
                }
                legacy_namespace
            }
            ResolveResult::Unknown(_) => false,
        };
        let decoder = xml_reader.decoder();
        framing.observe(&event, decoder)?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let empty = matches!(event, Event::Empty(_));
                let element =
                    WorksheetElement::at_path(e.local_name().as_ref(), spreadsheet, &ancestors);
                if ancestors.contains(&WorksheetElement::SheetData) {
                    ws.unsupported_cell_content |= matches!(
                        element,
                        WorksheetElement::Other | WorksheetElement::InlineRun
                    );
                    for attr in e.attributes() {
                        let supported = attr.is_ok_and(|attr| match element {
                            WorksheetElement::Cell => {
                                matches!(attr.key.as_ref(), b"r" | b"s" | b"t")
                            }
                            WorksheetElement::InlineText => attr.key.as_ref() == b"xml:space",
                            WorksheetElement::Row => true,
                            _ => false,
                        });
                        ws.unsupported_cell_content |= !supported;
                    }
                }
                match element {
                    WorksheetElement::Cell => {
                        current_coord.clear();
                        current_type.clear();
                        current_style = 0;
                        current_style_explicit = false;
                        formula_text.clear();
                        val_text.clear();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    if let Ok(v) = std::str::from_utf8(&attr.value) {
                                        current_coord = v.to_owned();
                                    }
                                }
                                b"t" => {
                                    if let Ok(v) = std::str::from_utf8(&attr.value) {
                                        current_type = v.to_owned();
                                    }
                                }
                                b"s" => {
                                    current_style_explicit = true;
                                    current_style = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(0);
                                }
                                _ => {}
                            }
                        }
                        if empty {
                            if let Ok((row, col)) = coordinate_to_tuple(&current_coord) {
                                let cell = ws.set_cell_value(row, col, CellValue::Empty)?;
                                attach_cell_style(
                                    cell,
                                    styles,
                                    current_style,
                                    current_style_explicit,
                                );
                            }
                            current_coord.clear();
                        }
                    }
                    WorksheetElement::MergeCell => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                if let Ok(value) = attr.decode_and_unescape_value(decoder) {
                                    if let Ok(range) = range_boundaries(&value) {
                                        merged_ranges.push(range);
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
                if !empty {
                    ancestors.push(element);
                }
            }
            Event::Text(ref e) => match ancestors.last() {
                Some(WorksheetElement::Value | WorksheetElement::InlineText) => {
                    val_text.push_str(&decode_xml_text(e, decoder)?);
                }
                Some(WorksheetElement::Formula) => {
                    formula_text.push_str(&decode_xml_text(e, decoder)?);
                }
                _ if ancestors.contains(&WorksheetElement::SheetData) => {
                    ws.unsupported_cell_content |= !decode_xml_text(e, decoder)?.trim().is_empty();
                }
                _ => {}
            },
            Event::End(_) => {
                if ancestors.pop() == Some(WorksheetElement::Cell) {
                    if let Ok((row, col)) = coordinate_to_tuple(&current_coord) {
                        let cell_value = resolve_cell_value(
                            &current_type,
                            current_style,
                            &val_text,
                            &formula_text,
                            shared_strings,
                            styles,
                            data_only,
                        );
                        let cell = ws.set_cell_value(row, col, cell_value)?;
                        attach_cell_style(cell, styles, current_style, current_style_explicit);
                    }
                    current_coord.clear();
                }
            }
            Event::Comment(_) | Event::PI(_) | Event::CData(_)
                if ancestors.contains(&WorksheetElement::SheetData) =>
            {
                ws.unsupported_cell_content = true;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    clear_merged_values(ws, &merged_ranges);
    ws.merged_ranges = merged_ranges;
    ws.xml_template = framing.finish();
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub fn parse_worksheet_streaming<R: Read>(
    reader: R,
    shared_strings: &SharedStringsTable,
    styles: &Styles,
    min_row: Option<u32>,
    max_row: Option<u32>,
    min_col: Option<u32>,
    max_col: Option<u32>,
    data_only: bool,
) -> RpyxlResult<Vec<Vec<(u32, u32, CellValue)>>> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut rows = Vec::new();
    let mut current_row_cells = Vec::new();

    let mut current_row_idx = 0u32;
    let mut current_coord = String::new();
    let mut current_type = String::new();
    let mut current_style = 0usize;

    let mut in_v = false;
    let mut in_f = false;
    let mut in_is_t = false;

    let mut formula_text = String::new();
    let mut val_text = String::new();

    let r_min = min_row.unwrap_or(1);
    let r_max = max_row.unwrap_or(1_048_576);
    let c_min = min_col.unwrap_or(1);
    let c_max = max_col.unwrap_or(16_384);

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"row" => {
                    current_row_cells.clear();
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"r" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                current_row_idx = v.parse().unwrap_or(0);
                            }
                        }
                    }
                }
                b"c" => {
                    current_coord.clear();
                    current_type.clear();
                    current_style = 0;
                    formula_text.clear();
                    val_text.clear();

                    for attr in e.attributes().flatten() {
                        let k = attr.key.local_name();
                        if k.as_ref() == b"r" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                current_coord = v.to_string();
                            }
                        } else if k.as_ref() == b"t" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                current_type = v.to_string();
                            }
                        } else if k.as_ref() == b"s" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                current_style = v.parse().unwrap_or(0);
                            }
                        }
                    }
                }
                b"v" => in_v = true,
                b"f" => in_f = true,
                b"t" => in_is_t = true,
                _ => {}
            },
            Ok(Event::Text(ref e)) => {
                if in_v {
                    let txt = decode_xml_text(e, xml_reader.decoder())?;
                    val_text.push_str(&txt);
                } else if in_f {
                    let txt = decode_xml_text(e, xml_reader.decoder())?;
                    formula_text.push_str(&txt);
                } else if in_is_t {
                    let txt = decode_xml_text(e, xml_reader.decoder())?;
                    val_text.push_str(&txt);
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => in_v = false,
                b"f" => in_f = false,
                b"t" => in_is_t = false,
                b"c" => {
                    if !current_coord.is_empty() {
                        if let Ok((row, col)) = coordinate_to_tuple(&current_coord) {
                            if row >= r_min && row <= r_max && col >= c_min && col <= c_max {
                                let cell_value = resolve_cell_value(
                                    &current_type,
                                    current_style,
                                    &val_text,
                                    &formula_text,
                                    shared_strings,
                                    styles,
                                    data_only,
                                );
                                current_row_cells.push((row, col, cell_value));
                            }
                        }
                    }
                }
                b"row" => {
                    if current_row_idx >= r_min && current_row_idx <= r_max {
                        if !current_row_cells.is_empty() {
                            rows.push(std::mem::take(&mut current_row_cells));
                        }
                    } else if current_row_idx > r_max {
                        break;
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(RpyxlError::IoError(format!(
                    "Error parsing streaming worksheet XML: {:?}",
                    e
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(rows)
}

fn parse_iso_temporal(value: &str) -> CellValue {
    if value.is_empty() {
        CellValue::Empty
    } else if NaiveDate::parse_from_str(value, "%Y-%m-%d").is_ok() {
        CellValue::Date(value.to_string())
    } else if NaiveTime::parse_from_str(value, "%H:%M:%S%.f").is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f").is_ok()
    {
        CellValue::DateTime(value.to_string())
    } else {
        // Do not label malformed or timezone-aware text as a naive temporal value.
        CellValue::String(value.to_string())
    }
}

fn resolve_cell_value(
    t_attr: &str,
    style_idx: usize,
    val_text: &str,
    formula_text: &str,
    shared_strings: &SharedStringsTable,
    styles: &Styles,
    data_only: bool,
) -> CellValue {
    let val_text = if matches!(t_attr, "inlineStr" | "str" | "e") {
        val_text
    } else {
        val_text.trim()
    };
    let raw_val = if t_attr == "s" {
        // Shared string index
        if let Ok(idx) = val_text.parse::<usize>() {
            shared_strings
                .get(idx)
                .map(|s| CellValue::String(s.to_string()))
                .unwrap_or(CellValue::Empty)
        } else {
            CellValue::Empty
        }
    } else if t_attr == "b" {
        CellValue::Bool(val_text == "1" || val_text.eq_ignore_ascii_case("true"))
    } else if t_attr == "inlineStr" || t_attr == "str" || t_attr == "e" {
        CellValue::String(val_text.to_string())
    } else if t_attr == "d" {
        parse_iso_temporal(val_text)
    } else if val_text.is_empty() {
        CellValue::Empty
    } else {
        // Number or Date
        if styles.is_date_format(style_idx) && !styles.is_duration_format(style_idx) {
            if let Ok(serial) = val_text.parse::<f64>() {
                if let Some(date_str) = styles.date_string_for_style(serial, style_idx) {
                    CellValue::DateTime(date_str)
                } else {
                    parse_numeric_or_string(val_text)
                }
            } else {
                parse_numeric_or_string(val_text)
            }
        } else {
            parse_numeric_or_string(val_text)
        }
    };

    if !formula_text.is_empty() {
        let fmt_formula = if formula_text.starts_with('=') {
            formula_text.to_string()
        } else {
            format!("={}", formula_text)
        };

        if data_only {
            raw_val
        } else {
            CellValue::Formula {
                formula: fmt_formula,
                cached_value: if raw_val.is_empty() {
                    None
                } else {
                    Some(Box::new(raw_val))
                },
            }
        }
    } else {
        raw_val
    }
}

pub struct StreamingRowReader<R: Read> {
    xml_reader: Reader<BufReader<R>>,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
    data_only: bool,
    buf: Vec<u8>,
    done: bool,
}

impl<R: Read> StreamingRowReader<R> {
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        reader: R,
        shared_strings: Arc<SharedStringsTable>,
        styles: Arc<Styles>,
        min_row: u32,
        max_row: u32,
        min_col: u32,
        max_col: u32,
        data_only: bool,
    ) -> Self {
        let mut xml_reader = Reader::from_reader(BufReader::new(reader));
        xml_reader.config_mut().trim_text(false);
        Self {
            xml_reader,
            shared_strings,
            styles,
            min_row,
            max_row,
            min_col,
            max_col,
            data_only,
            buf: Vec::new(),
            done: false,
        }
    }

    #[allow(clippy::type_complexity)]
    pub fn next_xml_row(&mut self) -> RpyxlResult<Option<(u32, Vec<(u32, CellValue)>)>> {
        if self.done {
            return Ok(None);
        }

        let mut current_row_idx = 0u32;
        let mut row_cells = Vec::new();
        let mut current_coord = String::new();
        let mut current_type = String::new();
        let mut current_style = 0usize;

        let mut in_v = false;
        let mut in_f = false;
        let mut in_is_t = false;

        let mut formula_text = String::new();
        let mut val_text = String::new();
        let mut inside_row = false;

        loop {
            match self.xml_reader.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"row" => {
                        inside_row = true;
                        current_row_idx = 0;
                        row_cells.clear();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"r" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_row_idx = v.parse().unwrap_or(0);
                                }
                            }
                        }
                    }
                    b"c" if inside_row => {
                        current_coord.clear();
                        current_type.clear();
                        current_style = 0;
                        formula_text.clear();
                        val_text.clear();

                        for attr in e.attributes().flatten() {
                            let k = attr.key.local_name();
                            if k.as_ref() == b"r" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_coord = v.to_string();
                                }
                            } else if k.as_ref() == b"t" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_type = v.to_string();
                                }
                            } else if k.as_ref() == b"s" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_style = v.parse().unwrap_or(0);
                                }
                            }
                        }
                    }
                    b"v" if inside_row => in_v = true,
                    b"f" if inside_row => in_f = true,
                    b"t" if inside_row => in_is_t = true,
                    _ => {}
                },
                Ok(Event::Text(ref e)) if inside_row => {
                    if in_v {
                        let txt = decode_xml_text(e, self.xml_reader.decoder())?;
                        val_text.push_str(&txt);
                    } else if in_f {
                        let txt = decode_xml_text(e, self.xml_reader.decoder())?;
                        formula_text.push_str(&txt);
                    } else if in_is_t {
                        let txt = decode_xml_text(e, self.xml_reader.decoder())?;
                        val_text.push_str(&txt);
                    }
                }
                Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                    b"v" => in_v = false,
                    b"f" => in_f = false,
                    b"t" => in_is_t = false,
                    b"c" if inside_row => {
                        if !current_coord.is_empty() {
                            if let Ok((row, col)) = coordinate_to_tuple(&current_coord) {
                                if current_row_idx == 0 {
                                    current_row_idx = row;
                                }
                                if col >= self.min_col && col <= self.max_col {
                                    let cell_value = resolve_cell_value(
                                        &current_type,
                                        current_style,
                                        &val_text,
                                        &formula_text,
                                        &self.shared_strings,
                                        &self.styles,
                                        self.data_only,
                                    );
                                    row_cells.push((col, cell_value));
                                }
                            }
                        }
                    }
                    b"row" if inside_row => {
                        inside_row = false;
                        if current_row_idx >= self.min_row && current_row_idx <= self.max_row {
                            self.buf.clear();
                            return Ok(Some((current_row_idx, row_cells)));
                        } else if current_row_idx > self.max_row {
                            self.done = true;
                            self.buf.clear();
                            return Ok(None);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => {
                    self.done = true;
                    break;
                }
                Err(e) => {
                    self.done = true;
                    return Err(RpyxlError::IoError(format!(
                        "Error in StreamingRowReader: {:?}",
                        e
                    )));
                }
                _ => {}
            }
            self.buf.clear();
        }

        Ok(None)
    }
}

pub fn read_single_cell<R: Read>(
    reader: R,
    target_row: u32,
    target_col: u32,
    shared_strings: &SharedStringsTable,
    styles: &Styles,
    data_only: bool,
) -> RpyxlResult<CellValue> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut current_row = 0u32;
    let mut current_coord = String::new();
    let mut current_type = String::new();
    let mut current_style = 0usize;
    let mut in_v = false;
    let mut in_f = false;
    let mut in_is_t = false;
    let mut formula_text = String::new();
    let mut val_text = String::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"row" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"r" {
                            if let Ok(v) = std::str::from_utf8(&attr.value) {
                                current_row = v.parse().unwrap_or(0);
                            }
                        }
                    }
                    if current_row > target_row {
                        return Ok(CellValue::Empty);
                    }
                }
                b"c" => {
                    if current_row == target_row || current_row == 0 {
                        current_coord.clear();
                        current_type.clear();
                        current_style = 0;
                        formula_text.clear();
                        val_text.clear();

                        for attr in e.attributes().flatten() {
                            let k = attr.key.local_name();
                            if k.as_ref() == b"r" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_coord = v.to_string();
                                }
                            } else if k.as_ref() == b"t" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_type = v.to_string();
                                }
                            } else if k.as_ref() == b"s" {
                                if let Ok(v) = std::str::from_utf8(&attr.value) {
                                    current_style = v.parse().unwrap_or(0);
                                }
                            }
                        }
                    }
                }
                b"v" if current_row == target_row || current_row == 0 => in_v = true,
                b"f" if current_row == target_row || current_row == 0 => in_f = true,
                b"t" if current_row == target_row || current_row == 0 => in_is_t = true,
                _ => {}
            },
            Ok(Event::Text(ref e)) => {
                if in_v {
                    let txt = decode_xml_text(e, xml_reader.decoder())?;
                    val_text.push_str(&txt);
                } else if in_f {
                    let txt = decode_xml_text(e, xml_reader.decoder())?;
                    formula_text.push_str(&txt);
                } else if in_is_t {
                    let txt = decode_xml_text(e, xml_reader.decoder())?;
                    val_text.push_str(&txt);
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => in_v = false,
                b"f" => in_f = false,
                b"t" => in_is_t = false,
                b"c" => {
                    if !current_coord.is_empty() {
                        if let Ok((r, c)) = coordinate_to_tuple(&current_coord) {
                            if r == target_row && c == target_col {
                                return Ok(resolve_cell_value(
                                    &current_type,
                                    current_style,
                                    &val_text,
                                    &formula_text,
                                    shared_strings,
                                    styles,
                                    data_only,
                                ));
                            }
                        }
                    }
                }
                b"row" if current_row == target_row => {
                    return Ok(CellValue::Empty);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(RpyxlError::IoError(format!(
                    "Error parsing sheet XML: {:?}",
                    e
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(CellValue::Empty)
}

fn parse_numeric_or_string(txt: &str) -> CellValue {
    if let Ok(i) = txt.parse::<i64>() {
        CellValue::Int(i)
    } else if let Ok(f) = txt.parse::<f64>() {
        CellValue::Float(f)
    } else {
        CellValue::String(txt.to_string())
    }
}

#[cfg(test)]
mod temporal_format_tests {
    use super::*;

    #[test]
    fn test_numeric_date_style_midnight_is_datetime() {
        let mut styles = Styles::from_reader(
            &br#"<styleSheet><numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs><xf numFmtId="164"/></cellXfs></styleSheet>"#[..],
        ).unwrap();
        let shared = SharedStringsTable::default();
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="0"><v>-1</v></c><c r="C1" s="0"><v>0</v></c><c r="D1" s="0"><v>0.5</v></c><c r="E1" s="0" t="d"><v>2023-01-01</v></c></row></sheetData></worksheet>"#;
        for date_1904 in [false, true] {
            styles.date_1904 = date_1904;
            let mut ws = Worksheet::new("Sheet").unwrap();
            parse_worksheet_into(&xml[..], &mut ws, &shared, &styles, false).unwrap();
            let expected = [
                CellValue::DateTime(
                    if date_1904 {
                        "1904-01-02 00:00:00"
                    } else {
                        "1900-01-01 00:00:00"
                    }
                    .into(),
                ),
                CellValue::DateTime(
                    if date_1904 {
                        "1903-12-31 00:00:00"
                    } else {
                        "1899-12-29 00:00:00"
                    }
                    .into(),
                ),
                CellValue::DateTime("00:00:00".into()),
                CellValue::DateTime("12:00:00".into()),
                CellValue::Date("2023-01-01".into()),
            ];
            let rows = parse_worksheet_streaming(
                &xml[..],
                &shared,
                &styles,
                None,
                None,
                None,
                None,
                false,
            )
            .unwrap();
            for (idx, value) in expected.iter().enumerate() {
                let col = idx as u32 + 1;
                assert_eq!(&ws.get_cell_value(1, col), value);
                assert_eq!(
                    &read_single_cell(&xml[..], 1, col, &shared, &styles, false).unwrap(),
                    value
                );
                assert_eq!(&rows[0][idx].2, value);
            }
            assert_eq!(
                ws.cells().get(&(1, 1)).unwrap().number_format.as_deref(),
                Some("yyyy-mm-dd")
            );
        }
    }

    #[test]
    fn test_iso_temporals_with_general_style() {
        let styles = Styles::from_reader(
            &br#"<styleSheet><cellXfs><xf numFmtId="0"/></cellXfs></styleSheet>"#[..],
        )
        .unwrap();
        let shared = SharedStringsTable::default();
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1" t="d"><v>2023-01-01</v></c><c r="B1" t="d"><v>2023-01-01T00:00:00</v></c><c r="C1" t="d"><v>12:34:56.123</v></c><c r="D1"><v>42</v></c></row></sheetData></worksheet>"#;
        let mut ws = Worksheet::new("Sheet").unwrap();
        parse_worksheet_into(&xml[..], &mut ws, &shared, &styles, false).unwrap();
        assert_eq!(
            ws.get_cell_value(1, 1),
            CellValue::Date("2023-01-01".into())
        );
        assert_eq!(
            ws.get_cell_value(1, 2),
            CellValue::DateTime("2023-01-01T00:00:00".into())
        );
        assert_eq!(
            ws.get_cell_value(1, 3),
            CellValue::DateTime("12:34:56.123".into())
        );
        for col in 1..=3 {
            assert_ne!(
                ws.cells().get(&(1, col)).unwrap().number_format.as_deref(),
                Some("General")
            );
            let value = read_single_cell(&xml[..], 1, col, &shared, &styles, false).unwrap();
            assert_eq!(value, ws.get_cell_value(1, col));
        }
        assert_eq!(
            ws.cells().get(&(1, 4)).unwrap().number_format.as_deref(),
            Some("General")
        );
        assert_eq!(
            parse_iso_temporal("2023-02-30"),
            CellValue::String("2023-02-30".into())
        );
        assert_eq!(
            parse_iso_temporal("25:00:00"),
            CellValue::String("25:00:00".into())
        );
    }
}
