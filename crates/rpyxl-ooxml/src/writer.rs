use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use rayon::prelude::*;
use rpyxl_core::cell::CellValue;
use rpyxl_core::coordinate::{tuple_to_coordinate, write_coordinate};
use rpyxl_core::error::{RpyxlError, RpyxlResult};
use rpyxl_core::workbook::Workbook;
use rpyxl_core::worksheet::Worksheet;
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;
use zip::ZipWriter;

use crate::styles::is_date_format_code;

pub fn save_workbook<P: AsRef<Path>>(path: P, wb: &Workbook) -> RpyxlResult<()> {
    let file = File::create(path.as_ref())
        .map_err(|e| RpyxlError::IoError(format!("Failed to create file: {}", e)))?;
    let writer = BufWriter::new(file);
    write_workbook(writer, wb)
}

pub fn write_workbook<W: Write + Seek>(writer: W, wb: &Workbook) -> RpyxlResult<()> {
    let mut formats = default_number_formats();
    for ws in wb.sheets() {
        for cell in ws.cells().values() {
            let code = effective_number_format(&cell.value, cell.number_format.as_deref());
            if !formats.iter().any(|format| format == code) {
                formats.push(code.to_string());
            }
        }
    }
    let mut zip = ZipWriter::new(writer);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    // 1. _rels/.rels
    zip.start_file("_rels/.rels", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    zip.write_all(RELS_ROOT.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 2. [Content_Types].xml
    zip.start_file("[Content_Types].xml", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    let content_types = generate_content_types(wb.sheet_count());
    zip.write_all(content_types.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 3. xl/_rels/workbook.xml.rels
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    let wb_rels = generate_workbook_rels(wb.sheet_count());
    zip.write_all(wb_rels.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 4. xl/workbook.xml
    zip.start_file("xl/workbook.xml", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    let wb_xml = generate_workbook_xml(wb);
    zip.write_all(wb_xml.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 5. xl/styles.xml
    zip.start_file("xl/styles.xml", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    zip.write_all(generate_styles_xml(&formats).as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 6. Worksheets (parallel serialization via Rayon if multi-sheet)
    let sheet_names = wb.sheetnames();
    let sheets: Vec<&Worksheet> = sheet_names
        .iter()
        .filter_map(|name| wb.get_sheet(name))
        .collect();

    let sheet_xmls: Vec<RpyxlResult<String>> = if sheets.len() > 1 {
        sheets
            .par_iter()
            .map(|ws| generate_worksheet_xml(ws, &formats, wb.date_1904))
            .collect()
    } else {
        sheets
            .iter()
            .map(|ws| generate_worksheet_xml(ws, &formats, wb.date_1904))
            .collect()
    };

    for (i, res) in sheet_xmls.into_iter().enumerate() {
        let sheet_idx = i + 1;
        let mut itoa_buf = itoa::Buffer::new();
        let mut sheet_path = String::with_capacity(32);
        sheet_path.push_str("xl/worksheets/sheet");
        sheet_path.push_str(itoa_buf.format(sheet_idx));
        sheet_path.push_str(".xml");

        zip.start_file(&sheet_path, options)
            .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
        let sheet_xml = res?;
        zip.write_all(sheet_xml.as_bytes())
            .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;
    }

    zip.finish()
        .map_err(|e| RpyxlError::IoError(format!("Failed to finalize ZIP archive: {}", e)))?;

    Ok(())
}

// -----------------------------------------------------------------------------
// Write-Only Streaming Implementation
// -----------------------------------------------------------------------------

pub struct WriteOnlySheet {
    pub title: String,
    pub row_count: u32,
    pub max_col: u32,
    writer: BufWriter<std::fs::File>,
    reader: std::fs::File,
    row_buf: String,
}

impl WriteOnlySheet {
    pub fn new(title: &str) -> RpyxlResult<Self> {
        let file = tempfile::tempfile()
            .map_err(|e| RpyxlError::IoError(format!("Failed to create temporary file: {}", e)))?;
        let read_handle = file
            .try_clone()
            .map_err(|e| RpyxlError::IoError(format!("Failed to clone temp file handle: {}", e)))?;
        Ok(Self {
            title: title.to_string(),
            row_count: 0,
            max_col: 0,
            writer: BufWriter::with_capacity(65536, file),
            reader: read_handle,
            row_buf: String::with_capacity(2048),
        })
    }

    pub fn title(&self) -> &str {
        &self.title
    }

    pub fn set_title(&mut self, title: &str) {
        self.title = title.to_string();
    }

    pub fn append_row(&mut self, row_values: &[CellValue]) -> RpyxlResult<()> {
        self.row_count += 1;
        let r = self.row_count;
        let col_count = row_values.len() as u32;
        if col_count > self.max_col {
            self.max_col = col_count;
        }

        self.row_buf.clear();
        self.row_buf.push_str("<row r=\"");
        let mut itoa_buf = itoa::Buffer::new();
        self.row_buf.push_str(itoa_buf.format(r));
        self.row_buf.push_str("\">");
        for (idx, val) in row_values.iter().enumerate() {
            let c = (idx + 1) as u32;
            write_cell(r, c, val, &mut self.row_buf);
        }
        self.row_buf.push_str("</row>\n");
        self.writer
            .write_all(self.row_buf.as_bytes())
            .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;
        Ok(())
    }
}

pub fn save_write_only_workbook<P: AsRef<Path>>(
    path: P,
    sheets: &mut [&mut WriteOnlySheet],
) -> RpyxlResult<()> {
    let file = File::create(path.as_ref())
        .map_err(|e| RpyxlError::IoError(format!("Failed to create file: {}", e)))?;
    let writer = BufWriter::new(file);
    write_write_only_workbook(writer, sheets)
}

pub fn save_write_only_workbook_slice<P: AsRef<Path>>(
    path: P,
    sheets: &mut [WriteOnlySheet],
) -> RpyxlResult<()> {
    let mut refs: Vec<&mut WriteOnlySheet> = sheets.iter_mut().collect();
    save_write_only_workbook(path, &mut refs)
}

pub fn write_write_only_workbook<W: Write + Seek>(
    writer: W,
    sheets: &mut [&mut WriteOnlySheet],
) -> RpyxlResult<()> {
    if sheets.is_empty() {
        return Err(RpyxlError::ValueError(
            "Cannot save empty write-only workbook".to_string(),
        ));
    }

    let mut zip = ZipWriter::new(writer);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    // 1. _rels/.rels
    zip.start_file("_rels/.rels", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    zip.write_all(RELS_ROOT.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 2. [Content_Types].xml
    zip.start_file("[Content_Types].xml", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    let content_types = generate_content_types(sheets.len());
    zip.write_all(content_types.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 3. xl/_rels/workbook.xml.rels
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    let wb_rels = generate_workbook_rels(sheets.len());
    zip.write_all(wb_rels.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 4. xl/workbook.xml
    zip.start_file("xl/workbook.xml", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    let mut wb_xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
        xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n\
        <bookViews><workbookView activeTab=\"0\"/></bookViews>\n\
        <sheets>\n",
    );
    for (i, sheet) in sheets.iter().enumerate() {
        let sheet_id = i + 1;
        let r_id = sheet_id + 1;
        wb_xml.push_str(&format!(
            "  <sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>\n",
            sheet.title, sheet_id, r_id
        ));
    }
    wb_xml.push_str("</sheets>\n</workbook>");
    zip.write_all(wb_xml.as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 5. xl/styles.xml
    zip.start_file("xl/styles.xml", options)
        .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;
    zip.write_all(generate_styles_xml(&default_number_formats()).as_bytes())
        .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

    // 6. Write each sheet streaming data
    for (i, sheet) in sheets.iter_mut().enumerate() {
        sheet
            .writer
            .flush()
            .map_err(|e| RpyxlError::IoError(format!("Flush error: {}", e)))?;
        sheet
            .reader
            .seek(std::io::SeekFrom::Start(0))
            .map_err(|e| RpyxlError::IoError(format!("Seek error: {}", e)))?;

        let sheet_idx = i + 1;
        let sheet_path = format!("xl/worksheets/sheet{}.xml", sheet_idx);
        zip.start_file(&sheet_path, options)
            .map_err(|e| RpyxlError::IoError(format!("Zip error: {}", e)))?;

        let max_r = if sheet.row_count == 0 {
            1
        } else {
            sheet.row_count
        };
        let max_c = if sheet.max_col == 0 { 1 } else { sheet.max_col };
        let end_coord = tuple_to_coordinate(max_r, max_c).unwrap_or_else(|_| "A1".to_string());

        let header = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
            <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n\
            <dimension ref=\"A1:{}\"/>\n\
            <sheetData>\n",
            end_coord
        );
        zip.write_all(header.as_bytes())
            .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;

        std::io::copy(&mut sheet.reader, &mut zip)
            .map_err(|e| RpyxlError::IoError(format!("Streaming copy error: {}", e)))?;

        zip.write_all(b"</sheetData>\n</worksheet>")
            .map_err(|e| RpyxlError::IoError(format!("Write error: {}", e)))?;
    }

    zip.finish()
        .map_err(|e| RpyxlError::IoError(format!("Failed to finalize ZIP archive: {}", e)))?;

    Ok(())
}

// -----------------------------------------------------------------------------
// XML Helpers and Generators
// -----------------------------------------------------------------------------

const RELS_ROOT: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  \
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n\
</Relationships>";

fn default_number_formats() -> Vec<String> {
    [
        "General",
        "yyyy-mm-dd h:mm:ss.000",
        "yyyy-mm-dd",
        "h:mm:ss.000",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect()
}

/// Returns the default format for a value, including standalone ISO times and formula caches.
pub fn default_number_format(value: &CellValue) -> &'static str {
    match value {
        CellValue::Date(_) => "yyyy-mm-dd",
        CellValue::DateTime(s) if NaiveTime::parse_from_str(s, "%H:%M:%S%.f").is_ok() => {
            "h:mm:ss.000"
        }
        CellValue::DateTime(_) => "yyyy-mm-dd h:mm:ss.000",
        CellValue::Formula {
            cached_value: Some(value),
            ..
        } => default_number_format(value),
        _ => "General",
    }
}

/// Keeps an existing temporal format, but never writes a temporal serial with a
/// non-temporal format that would cause it to be loaded as an ordinary number.
pub fn effective_number_format<'a>(value: &CellValue, number_format: Option<&'a str>) -> &'a str {
    let default = default_number_format(value);
    number_format
        .filter(|code| default == "General" || is_date_format_code(code))
        .unwrap_or(default)
}

fn generate_styles_xml(formats: &[String]) -> String {
    let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    xml.push_str(&format!("<numFmts count=\"{}\">", formats.len() - 1));
    for (idx, code) in formats.iter().enumerate().skip(1) {
        xml.push_str(&format!("<numFmt numFmtId=\"{}\" formatCode=\"", 163 + idx));
        // Attribute whitespace must use character references to survive XML normalization.
        let mut escaped = String::new();
        write_xml_escaped(code, &mut escaped);
        xml.push_str(&escaped.replace('\n', "&#10;").replace('\t', "&#9;"));
        xml.push_str("\"/>");
    }
    xml.push_str("</numFmts><fonts count=\"1\"><font><name val=\"Calibri\"/><family val=\"2\"/><sz val=\"11\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
    xml.push_str(&format!("<cellXfs count=\"{}\">", formats.len()));
    for idx in 0..formats.len() {
        let id = if idx == 0 { 0 } else { 163 + idx };
        xml.push_str(&format!("<xf numFmtId=\"{id}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"));
    }
    xml.push_str("</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");
    xml
}

fn generate_content_types(sheet_count: usize) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
        <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  \
        <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n  \
        <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  \
        <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n  \
        <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n",
    );
    for i in 1..=sheet_count {
        s.push_str(&format!(
            "  <Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
            i
        ));
    }
    s.push_str("</Types>");
    s
}

fn generate_workbook_rels(sheet_count: usize) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
        <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  \
        <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n",
    );
    for i in 1..=sheet_count {
        let r_id = i + 1;
        s.push_str(&format!(
            "  <Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>\n",
            r_id, i
        ));
    }
    s.push_str("</Relationships>");
    s
}

fn generate_workbook_xml(wb: &Workbook) -> String {
    let active_idx = wb.active_index();
    let mut s = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
        xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n  \
        <workbookPr date1904=\"{}\"/>\n  \
        <bookViews><workbookView activeTab=\"{}\"/></bookViews>\n  \
        <sheets>\n",
        if wb.date_1904 { "1" } else { "0" },
        active_idx
    );
    for (i, name) in wb.sheetnames().iter().enumerate() {
        let sheet_id = i + 1;
        let r_id = sheet_id + 1;
        s.push_str(&format!(
            "    <sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>\n",
            name, sheet_id, r_id
        ));
    }
    s.push_str("  </sheets>\n</workbook>");
    s
}

fn generate_worksheet_xml(
    ws: &Worksheet,
    formats: &[String],
    date_1904: bool,
) -> RpyxlResult<String> {
    let min_r = ws.min_row();
    let max_r = ws.max_row();
    let min_c = ws.min_column();
    let max_c = ws.max_column();

    let cell_count = ws.cell_count();
    let estimated_size = 512 + cell_count * 45;
    let mut s = String::with_capacity(estimated_size);

    s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n  <dimension ref=\"");
    write_coordinate(min_r, min_c, &mut s)?;
    s.push(':');
    write_coordinate(max_r, max_c, &mut s)?;
    s.push_str("\"/>\n  <sheetData>\n");

    let mut itoa_buf = itoa::Buffer::new();
    let mut current_row = 0;
    for (&(r, c), cell) in ws.cells() {
        if r != current_row {
            if current_row != 0 {
                s.push_str("</row>\n");
            }
            current_row = r;
            s.push_str("    <row r=\"");
            s.push_str(itoa_buf.format(r));
            s.push_str("\">");
        }
        let code = effective_number_format(&cell.value, cell.number_format.as_deref());
        let style = formats
            .iter()
            .position(|format| format == code)
            .unwrap_or(0);
        write_cell_styled(r, c, &cell.value, &mut s, style, date_1904);
    }
    if current_row != 0 {
        s.push_str("</row>\n");
    }

    s.push_str("  </sheetData>\n</worksheet>");
    Ok(s)
}

#[inline]
fn write_cell_open(row: u32, col: u32, style: usize, out: &mut String) {
    out.push_str("<c r=\"");
    let _ = write_coordinate(row, col, out);
    if style != 0 {
        out.push_str("\" s=\"");
        out.push_str(itoa::Buffer::new().format(style));
    }
}

fn write_cell(row: u32, col: u32, val: &CellValue, out: &mut String) {
    let style = match default_number_format(val) {
        "yyyy-mm-dd h:mm:ss.000" => 1,
        "yyyy-mm-dd" => 2,
        "h:mm:ss.000" => 3,
        _ => 0,
    };
    write_cell_styled(row, col, val, out, style, false);
}

fn write_cell_styled(
    row: u32,
    col: u32,
    val: &CellValue,
    out: &mut String,
    style: usize,
    date_1904: bool,
) {
    match val {
        CellValue::Empty => {
            if style != 0 {
                write_cell_open(row, col, style, out);
                out.push_str("\"/>");
            }
        }
        CellValue::String(s) => {
            write_cell_open(row, col, style, out);
            out.push_str("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            write_xml_escaped(s, out);
            out.push_str("</t></is></c>");
        }
        CellValue::Int(i) => {
            write_cell_open(row, col, style, out);
            out.push_str("\" t=\"n\"><v>");
            let mut itoa_buf = itoa::Buffer::new();
            out.push_str(itoa_buf.format(*i));
            out.push_str("</v></c>");
        }
        CellValue::Float(f) => {
            write_cell_open(row, col, style, out);
            out.push_str("\" t=\"n\"><v>");
            let mut ryu_buf = ryu::Buffer::new();
            out.push_str(ryu_buf.format(*f));
            out.push_str("</v></c>");
        }
        CellValue::Bool(b) => {
            write_cell_open(row, col, style, out);
            out.push_str("\" t=\"b\"><v>");
            out.push(if *b { '1' } else { '0' });
            out.push_str("</v></c>");
        }
        CellValue::Date(s) | CellValue::DateTime(s) => {
            if let Some(serial) = temporal_to_excel_serial(s, date_1904) {
                write_cell_open(row, col, style, out);
                out.push_str("\" t=\"n\"><v>");
                let mut ryu_buf = ryu::Buffer::new();
                out.push_str(ryu_buf.format(serial));
                out.push_str("</v></c>");
            } else {
                write_cell_open(row, col, style, out);
                out.push_str("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                write_xml_escaped(s, out);
                out.push_str("</t></is></c>");
            }
        }
        CellValue::Formula {
            formula,
            cached_value,
        } => {
            let formula_clean = formula.trim_start_matches('=');
            write_cell_open(row, col, style, out);
            match cached_value.as_deref() {
                Some(CellValue::String(_)) => out.push_str("\" t=\"str"),
                Some(CellValue::Bool(_)) => out.push_str("\" t=\"b"),
                _ => {}
            }
            out.push_str("\"><f>");
            write_xml_escaped(formula_clean, out);
            out.push_str("</f>");
            if let Some(cached) = cached_value {
                match cached.as_ref() {
                    CellValue::Int(i) => {
                        out.push_str("<v>");
                        let mut itoa_buf = itoa::Buffer::new();
                        out.push_str(itoa_buf.format(*i));
                        out.push_str("</v>");
                    }
                    CellValue::Float(f) => {
                        out.push_str("<v>");
                        let mut ryu_buf = ryu::Buffer::new();
                        out.push_str(ryu_buf.format(*f));
                        out.push_str("</v>");
                    }
                    CellValue::String(s) => {
                        out.push_str("<v>");
                        write_xml_escaped(s, out);
                        out.push_str("</v>");
                    }
                    CellValue::Bool(b) => {
                        out.push_str("<v>");
                        out.push(if *b { '1' } else { '0' });
                        out.push_str("</v>");
                    }
                    CellValue::Date(s) | CellValue::DateTime(s) => {
                        if let Some(serial) = temporal_to_excel_serial(s, date_1904) {
                            out.push_str("<v>");
                            out.push_str(ryu::Buffer::new().format(serial));
                            out.push_str("</v>");
                        } else {
                            out.push_str("<v/>");
                        }
                    }
                    _ => out.push_str("<v/>"),
                }
            } else {
                out.push_str("<v/>");
            }
            out.push_str("</c>");
        }
    }
}

#[inline]
fn write_xml_escaped(s: &str, out: &mut String) {
    let bytes = s.as_bytes();
    let mut last_idx = 0;

    for (idx, &b) in bytes.iter().enumerate() {
        let esc = match b {
            b'&' => "&amp;",
            b'<' => "&lt;",
            b'>' => "&gt;",
            b'"' => "&quot;",
            b'\'' => "&apos;",
            b'\r' => "&#13;",
            _ => continue,
        };
        if idx > last_idx {
            if let Ok(sub) = std::str::from_utf8(&bytes[last_idx..idx]) {
                out.push_str(sub);
            }
        }
        out.push_str(esc);
        last_idx = idx + 1;
    }

    if last_idx == 0 {
        out.push_str(s);
    } else if last_idx < bytes.len() {
        if let Ok(sub) = std::str::from_utf8(&bytes[last_idx..]) {
            out.push_str(sub);
        }
    }
}

#[cfg(test)]
fn date_to_excel_serial(s: &str) -> Option<f64> {
    temporal_to_excel_serial(s, false)
}

fn temporal_to_excel_serial(s: &str, date_1904: bool) -> Option<f64> {
    let fraction = |time: NaiveTime| {
        // Leap seconds and timezone-aware values have no Excel serial equivalent.
        if time.nanosecond() >= 1_000_000_000 {
            return None;
        }
        Some((time.num_seconds_from_midnight() as f64 + time.nanosecond() as f64 / 1e9) / 86400.0)
    };
    if let Ok(time) = NaiveTime::parse_from_str(s, "%H:%M:%S%.f") {
        return fraction(time);
    }
    let (date, frac) = if s.len() == 10 {
        (NaiveDate::parse_from_str(s, "%Y-%m-%d").ok()?, 0.0)
    } else {
        let dt = NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S%.f")
            .or_else(|_| NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S%.f"))
            .ok()?;
        (dt.date(), fraction(dt.time())?)
    };
    let base = if date_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1)?
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30)?
    };
    let mut days = (date - base).num_days();
    if !date_1904 && days > 0 && days <= 60 {
        days -= 1;
    }
    Some(days as f64 + frac)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::XlsxPackage;

    #[test]
    fn test_text_serialization_preserves_carriage_returns_and_whitespace() {
        let mut output = String::new();
        write_cell(
            1,
            1,
            &CellValue::String("  héllo <&>\r\nline2\r  ".to_string()),
            &mut output,
        );
        assert!(output.contains("<t xml:space=\"preserve\">"));
        assert!(output.contains("  héllo &lt;&amp;&gt;&#13;\nline2&#13;  </t>"));
        assert!(!output.contains('\r'));
    }

    #[test]
    fn test_date_to_excel_serial() {
        assert_eq!(date_to_excel_serial("2023-01-01"), Some(44927.0));
        assert_eq!(date_to_excel_serial("2023-01-01 12:00:00"), Some(44927.5));
        assert_eq!(date_to_excel_serial("12:00:00"), Some(0.5));
        assert_eq!(date_to_excel_serial("1900-01-01"), Some(1.0));
        assert_eq!(date_to_excel_serial("1900-02-28"), Some(59.0));
        assert_eq!(date_to_excel_serial("1900-03-01"), Some(61.0));
        assert_eq!(temporal_to_excel_serial("1904-01-02", true), Some(1.0));
        for (input, expected) in [("12:00:00.1", 0.1), ("12:00:00.123456", 0.123456)] {
            let seconds = (date_to_excel_serial(input).unwrap() - 0.5) * 86400.0;
            assert!((seconds - expected).abs() < 1e-8);
        }
        assert_eq!(date_to_excel_serial("25:00:00"), None);
        assert_eq!(date_to_excel_serial("2023-01-01T12:00:00+01:00"), None);
    }

    #[test]
    fn test_temporal_non_temporal_format_fallback() {
        let values = [
            (CellValue::Date("2023-01-01".into()), "yyyy-mm-dd"),
            (
                CellValue::DateTime("2023-01-01T00:00:00".into()),
                "yyyy-mm-dd h:mm:ss.000",
            ),
            (CellValue::DateTime("12:34:56.123".into()), "h:mm:ss.000"),
            (
                CellValue::formula_with_cached("=A1", CellValue::Date("2023-01-01".into())),
                "yyyy-mm-dd",
            ),
        ];
        for (value, expected) in values {
            for format in [None, Some("General"), Some("general"), Some("0.00")] {
                assert_eq!(effective_number_format(&value, format), expected);
            }
            assert_eq!(
                effective_number_format(&value, Some("yyyy/mm/dd hh:mm:ss")),
                "yyyy/mm/dd hh:mm:ss"
            );
        }
        assert_eq!(
            effective_number_format(&CellValue::Int(1), Some("General")),
            "General"
        );
        assert_eq!(
            effective_number_format(&CellValue::Int(1), Some("0.00")),
            "0.00"
        );
    }

    #[test]
    fn test_general_formatted_temporals_roundtrip() {
        let mut wb = Workbook::new();
        let values = [
            CellValue::Date("2023-01-01".into()),
            CellValue::DateTime("2023-01-01 00:00:00".into()),
            CellValue::DateTime("12:34:56.123".into()),
        ];
        for (idx, value) in values.iter().enumerate() {
            let cell = wb
                .active_mut()
                .set_cell_value(1, idx as u32 + 1, value.clone())
                .unwrap();
            cell.number_format = Some("General".into());
        }
        let mut buffer = std::io::Cursor::new(Vec::new());
        write_workbook(&mut buffer, &wb).unwrap();
        buffer.set_position(0);
        let loaded = XlsxPackage::from_reader(buffer, false, false)
            .unwrap()
            .to_workbook()
            .unwrap();
        let expected = [
            CellValue::DateTime("2023-01-01 00:00:00".into()),
            CellValue::DateTime("2023-01-01 00:00:00".into()),
            CellValue::DateTime("12:34:56.123".into()),
        ];
        for (idx, value) in expected.iter().enumerate() {
            assert_eq!(&loaded.active().get_cell_value(1, idx as u32 + 1), value);
        }
    }

    #[test]
    fn test_temporal_formats_and_epochs_roundtrip() {
        for date_1904 in [false, true] {
            let mut wb = Workbook::new();
            wb.date_1904 = date_1904;
            let ws = wb.active_mut();
            let values = [
                CellValue::Date("2023-01-01".into()),
                CellValue::DateTime("2023-01-01T12:34:56.123456".into()),
                CellValue::DateTime("12:34:56.123456".into()),
                CellValue::DateTime("2023-01-01T00:00:00".into()),
                CellValue::Float(1234.5),
                CellValue::Empty,
                CellValue::formula_with_cached(
                    "=B1",
                    CellValue::DateTime("2023-01-01 12:34:56.123".into()),
                ),
                CellValue::Float(1.5),
            ];
            for (idx, value) in values.into_iter().enumerate() {
                ws.set_cell_value(1, idx as u32 + 1, value).unwrap();
            }
            let custom = "#,##0.00 \"USD & <tax>\"";
            ws.get_cell_mut(1, 5).unwrap().number_format = Some(custom.into());
            ws.get_cell_mut(1, 6).unwrap().number_format = Some("0.0000".into());
            ws.get_cell_mut(1, 8).unwrap().number_format = Some("[h]:mm:ss".into());
            let mut buffer = std::io::Cursor::new(Vec::new());
            write_workbook(&mut buffer, &wb).unwrap();
            buffer.set_position(0);
            let mut package = XlsxPackage::from_reader(buffer, false, false).unwrap();
            assert_eq!(package.styles.date_1904, date_1904);
            let loaded = package.to_workbook().unwrap();
            assert_eq!(loaded.date_1904, date_1904);
            let ws = loaded.active();
            assert_eq!(
                ws.get_cell_value(1, 1),
                CellValue::DateTime("2023-01-01 00:00:00".into())
            );
            assert_eq!(
                ws.get_cell_value(1, 2),
                CellValue::DateTime("2023-01-01 12:34:56.123".into())
            );
            assert_eq!(
                ws.get_cell_value(1, 3),
                CellValue::DateTime("12:34:56.123".into())
            );
            assert_eq!(
                ws.get_cell_value(1, 4),
                CellValue::DateTime("2023-01-01 00:00:00".into())
            );
            assert_eq!(
                ws.cells().get(&(1, 5)).unwrap().number_format.as_deref(),
                Some(custom)
            );
            assert_eq!(
                ws.cells().get(&(1, 6)).unwrap().number_format.as_deref(),
                Some("0.0000")
            );
            assert_eq!(
                ws.get_cell_value(1, 7),
                CellValue::formula_with_cached(
                    "=B1",
                    CellValue::DateTime("2023-01-01 12:34:56.123".into())
                )
            );
            assert_eq!(ws.get_cell_value(1, 8), CellValue::Float(1.5));
            // Exercise the same epoch-aware conversion through a value-only reader.
            assert_eq!(
                package
                    .stream_rows("Sheet", Some(1), Some(1), Some(3), Some(3))
                    .unwrap()[0][0]
                    .2,
                CellValue::DateTime("12:34:56.123".into())
            );
        }
    }

    #[test]
    fn test_standard_save_and_read() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("standard_test.xlsx");

        let mut wb = Workbook::new();
        {
            let ws = wb.active_mut();
            ws.set_cell_value(1, 1, CellValue::Int(100)).unwrap();
            ws.set_cell_value(1, 2, CellValue::String("RPyXL Writer".to_string()))
                .unwrap();
            ws.set_cell_value(2, 1, CellValue::Float(std::f64::consts::PI))
                .unwrap();
            ws.set_cell_value(2, 2, CellValue::Bool(true)).unwrap();
            ws.set_cell_value(
                3,
                1,
                CellValue::Formula {
                    formula: "=SUM(A1:A2)".to_string(),
                    cached_value: Some(Box::new(CellValue::Int(103))),
                },
            )
            .unwrap();
        }

        save_workbook(&path, &wb).unwrap();
        assert!(path.exists());

        // Read it back with rpyxl-ooxml reader
        let mut pkg = XlsxPackage::open(&path, false, false).unwrap();
        let target_wb = pkg.to_workbook().unwrap();
        assert_eq!(target_wb.sheetnames(), &["Sheet"]);

        let ws = target_wb.get_sheet("Sheet").unwrap();
        assert_eq!(
            ws.get_cell(1, 1).map(|c| &c.value),
            Some(&CellValue::Int(100))
        );
        assert_eq!(
            ws.get_cell(1, 2).map(|c| &c.value),
            Some(&CellValue::String("RPyXL Writer".to_string()))
        );
        assert_eq!(
            ws.get_cell(2, 2).map(|c| &c.value),
            Some(&CellValue::Bool(true))
        );
    }

    #[test]
    fn test_write_only_save_and_read() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("write_only_test.xlsx");

        let mut sheet = WriteOnlySheet::new("StreamingSheet").unwrap();
        sheet
            .append_row(&[CellValue::Int(1), CellValue::String("Row1".to_string())])
            .unwrap();
        sheet
            .append_row(&[CellValue::Int(2), CellValue::String("Row2".to_string())])
            .unwrap();

        let mut sheets = [&mut sheet];
        save_write_only_workbook(&path, &mut sheets).unwrap();
        assert!(path.exists());

        // Read back with XlsxPackage
        let mut pkg = XlsxPackage::open(&path, false, false).unwrap();
        let target_wb = pkg.to_workbook().unwrap();
        assert_eq!(target_wb.sheetnames(), &["StreamingSheet"]);

        let ws = target_wb.get_sheet("StreamingSheet").unwrap();
        assert_eq!(
            ws.get_cell(1, 1).map(|c| &c.value),
            Some(&CellValue::Int(1))
        );
        assert_eq!(
            ws.get_cell(1, 2).map(|c| &c.value),
            Some(&CellValue::String("Row1".to_string()))
        );
        assert_eq!(
            ws.get_cell(2, 1).map(|c| &c.value),
            Some(&CellValue::Int(2))
        );
        assert_eq!(
            ws.get_cell(2, 2).map(|c| &c.value),
            Some(&CellValue::String("Row2".to_string()))
        );
    }
}
