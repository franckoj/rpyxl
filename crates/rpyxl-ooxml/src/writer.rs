use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

use chrono::NaiveDate;
use rayon::prelude::*;
use rpyxl_core::cell::CellValue;
use rpyxl_core::coordinate::{tuple_to_coordinate, write_coordinate};
use rpyxl_core::error::{RpyxlError, RpyxlResult};
use rpyxl_core::workbook::Workbook;
use rpyxl_core::worksheet::Worksheet;
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;
use zip::ZipWriter;

pub fn save_workbook<P: AsRef<Path>>(path: P, wb: &Workbook) -> RpyxlResult<()> {
    let file = File::create(path.as_ref())
        .map_err(|e| RpyxlError::IoError(format!("Failed to create file: {}", e)))?;
    let writer = BufWriter::new(file);
    write_workbook(writer, wb)
}

pub fn write_workbook<W: Write + Seek>(writer: W, wb: &Workbook) -> RpyxlResult<()> {
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
    zip.write_all(STYLES_XML.as_bytes())
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
            .map(|ws| generate_worksheet_xml(ws))
            .collect()
    } else {
        sheets.iter().map(|ws| generate_worksheet_xml(ws)).collect()
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
    zip.write_all(STYLES_XML.as_bytes())
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

const STYLES_XML: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n  \
<numFmts count=\"1\">\n    \
<numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/>\n  \
</numFmts>\n  \
<fonts count=\"1\">\n    \
<font><name val=\"Calibri\"/><family val=\"2\"/><sz val=\"11\"/></font>\n  \
</fonts>\n  \
<fills count=\"2\">\n    \
<fill><patternFill patternType=\"none\"/></fill>\n    \
<fill><patternFill patternType=\"gray125\"/></fill>\n  \
</fills>\n  \
<borders count=\"1\">\n    \
<border><left/><right/><top/><bottom/><diagonal/></border>\n  \
</borders>\n  \
<cellStyleXfs count=\"1\">\n    \
<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n  \
</cellStyleXfs>\n  \
<cellXfs count=\"2\">\n    \
<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n    \
<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>\n  \
</cellXfs>\n  \
<cellStyles count=\"1\">\n    \
<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n  \
</cellStyles>\n\
</styleSheet>";

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
        <bookViews><workbookView activeTab=\"{}\"/></bookViews>\n  \
        <sheets>\n",
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

fn generate_worksheet_xml(ws: &Worksheet) -> RpyxlResult<String> {
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
        write_cell(r, c, &cell.value, &mut s);
    }
    if current_row != 0 {
        s.push_str("</row>\n");
    }

    s.push_str("  </sheetData>\n</worksheet>");
    Ok(s)
}

#[inline]
fn write_cell_open(row: u32, col: u32, out: &mut String) {
    out.push_str("<c r=\"");
    let _ = write_coordinate(row, col, out);
}

fn write_cell(row: u32, col: u32, val: &CellValue, out: &mut String) {
    match val {
        CellValue::Empty => {}
        CellValue::String(s) => {
            write_cell_open(row, col, out);
            out.push_str("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            write_xml_escaped(s, out);
            out.push_str("</t></is></c>");
        }
        CellValue::Int(i) => {
            write_cell_open(row, col, out);
            out.push_str("\" t=\"n\"><v>");
            let mut itoa_buf = itoa::Buffer::new();
            out.push_str(itoa_buf.format(*i));
            out.push_str("</v></c>");
        }
        CellValue::Float(f) => {
            write_cell_open(row, col, out);
            out.push_str("\" t=\"n\"><v>");
            let mut ryu_buf = ryu::Buffer::new();
            out.push_str(ryu_buf.format(*f));
            out.push_str("</v></c>");
        }
        CellValue::Bool(b) => {
            write_cell_open(row, col, out);
            out.push_str("\" t=\"b\"><v>");
            out.push(if *b { '1' } else { '0' });
            out.push_str("</v></c>");
        }
        CellValue::Date(s) | CellValue::DateTime(s) => {
            if let Some(serial) = date_to_excel_serial(s) {
                write_cell_open(row, col, out);
                out.push_str("\" s=\"1\"><v>");
                let mut ryu_buf = ryu::Buffer::new();
                out.push_str(ryu_buf.format(serial));
                out.push_str("</v></c>");
            } else {
                write_cell_open(row, col, out);
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
            write_cell_open(row, col, out);
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

fn date_to_excel_serial(s: &str) -> Option<f64> {
    if s.len() < 10 {
        return None;
    }
    let year: i32 = s.get(0..4)?.parse().ok()?;
    let month: u32 = s.get(5..7)?.parse().ok()?;
    let day: u32 = s.get(8..10)?.parse().ok()?;

    let (hour, min, sec, micro) =
        if s.len() >= 19 && (s.as_bytes()[10] == b' ' || s.as_bytes()[10] == b'T') {
            let h: u32 = s.get(11..13)?.parse().ok()?;
            let m: u32 = s.get(14..16)?.parse().ok()?;
            let sec_str = &s[17..];
            if let Some((s_part, us_part)) = sec_str.split_once('.') {
                let s_val: u32 = s_part.parse().ok()?;
                let us_val: u32 = us_part.parse().ok()?;
                (h, m, s_val, us_val)
            } else {
                let s_val: u32 = sec_str.get(..2)?.parse().ok()?;
                (h, m, s_val, 0)
            }
        } else {
            (0, 0, 0, 0)
        };

    let date = NaiveDate::from_ymd_opt(year, month, day)?;
    let base = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let days = (date - base).num_days();

    let frac = (hour as f64 * 3600.0 + min as f64 * 60.0 + sec as f64 + micro as f64 / 1_000_000.0)
        / 86400.0;
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
