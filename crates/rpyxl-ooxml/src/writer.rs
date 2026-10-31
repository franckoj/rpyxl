use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::File;
use std::io::{BufWriter, Cursor, Read, Seek, SeekFrom, Write};
use std::path::Path;
use std::sync::Arc;

use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::NsReader;
use rayon::prelude::*;
use rpyxl_core::cell::{Cell, CellValue, RichStyle};
use rpyxl_core::coordinate::{tuple_to_coordinate, write_coordinate};
use rpyxl_core::error::{RpyxlError, RpyxlResult};
use rpyxl_core::package::PackageTemplate;
use rpyxl_core::workbook::Workbook;
use rpyxl_core::worksheet::Worksheet;
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;
use zip::{ZipArchive, ZipWriter};

use crate::styles::is_date_format_code;

pub fn save_workbook<P: AsRef<Path>>(path: P, wb: &Workbook) -> RpyxlResult<()> {
    save_staged(path.as_ref(), wb, false)
}

/// Rebuild the package explicitly, retaining the original style IDs when available.
/// Package-level metadata and worksheet framing are not copied in this mode.
pub fn save_workbook_generated<P: AsRef<Path>>(path: P, wb: &Workbook) -> RpyxlResult<()> {
    save_staged(path.as_ref(), wb, true)
}

fn save_staged(path: &Path, wb: &Workbook, generated: bool) -> RpyxlResult<()> {
    // Never truncate a destination (which may also be the source) until the entire
    // archive has serialized successfully and all source ZIP handles are closed.
    let mut staged = tempfile::tempfile().map_err(write_error)?;
    if generated {
        write_workbook_generated(&mut staged, wb)?;
    } else {
        write_workbook(&mut staged, wb)?;
    }
    staged.seek(SeekFrom::Start(0)).map_err(write_error)?;
    let mut destination = BufWriter::new(File::create(path).map_err(write_error)?);
    std::io::copy(&mut staged, &mut destination).map_err(write_error)?;
    destination.flush().map_err(write_error)
}

fn write_error(error: impl std::fmt::Display) -> RpyxlError {
    RpyxlError::IoError(format!("Cannot write XLSX package: {error}"))
}

fn template_error(message: impl std::fmt::Display) -> RpyxlError {
    RpyxlError::ValueError(format!("Cannot preserve source package: {message}"))
}

fn validate_cell_content(wb: &Workbook) -> RpyxlResult<()> {
    if let Some(sheet) = wb
        .sheets()
        .iter()
        .find(|sheet| sheet.unsupported_cell_content)
    {
        return Err(RpyxlError::ValueError(format!(
            "Cannot safely save worksheet {:?}: unsupported cell content (such as shared/array formulas, rich inline text, or cell metadata) would be lost",
            sheet.title
        )));
    }
    Ok(())
}

pub fn write_workbook<W: Write + Seek>(writer: W, wb: &Workbook) -> RpyxlResult<()> {
    validate_cell_content(wb)?;
    if wb.package_template.is_some() && !wb.can_preserve_package() {
        return Err(template_error(
            "workbook structure or worksheet templates changed; use save_workbook_generated / write_workbook_generated explicitly",
        ));
    }
    let styles = StyleTables::for_workbook(wb)?;
    if let Some(source) = &wb.package_template {
        write_preserved_package(writer, wb, source, &styles)
    } else {
        write_workbook_with_styles(writer, wb, &styles)
    }
}

/// Bypass source-package copying, but never renumber an available original style base.
pub fn write_workbook_generated<W: Write + Seek>(writer: W, wb: &Workbook) -> RpyxlResult<()> {
    validate_cell_content(wb)?;
    let styles = StyleTables::for_workbook(wb)?;
    write_workbook_with_styles(writer, wb, &styles)
}

fn write_workbook_with_styles<W: Write + Seek>(
    writer: W,
    wb: &Workbook,
    styles: &StyleTables,
) -> RpyxlResult<()> {
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
    zip.write_all(styles.to_xml().as_bytes())
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
            .map(|ws| generate_worksheet_xml(ws, styles, wb.date_1904))
            .collect()
    } else {
        sheets
            .iter()
            .map(|ws| generate_worksheet_xml(ws, styles, wb.date_1904))
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

fn read_package_part<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> RpyxlResult<String> {
    let mut xml = String::new();
    archive
        .by_name(path)
        .map_err(template_error)?
        .read_to_string(&mut xml)
        .map_err(template_error)?;
    Ok(xml)
}

fn resolve_part(base: &str, target: &str) -> RpyxlResult<String> {
    let joined = if target.starts_with('/') {
        target.to_owned()
    } else {
        format!(
            "{}{target}",
            base.rsplit_once('/')
                .map_or("", |(parent, _)| &base[..parent.len() + 1])
        )
    };
    let mut parts = Vec::new();
    for part in joined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts
                    .pop()
                    .ok_or_else(|| template_error("relationship escapes package root"))?;
            }
            part if !part.contains(['\\', ':', '#', '?']) => parts.push(part),
            _ => return Err(template_error("unsupported relationship target")),
        }
    }
    Ok(parts.join("/"))
}

fn package_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> RpyxlResult<HashMap<String, XmlElement>> {
    let root = parse_xml(&read_package_part(archive, path)?)?;
    if root.local_name() != "Relationships"
        || root.namespace != "http://schemas.openxmlformats.org/package/2006/relationships"
    {
        return Err(template_error("invalid package relationships"));
    }
    let mut relationships = HashMap::new();
    for child in root.children {
        let id = child
            .attributes
            .get("Id")
            .ok_or_else(|| template_error("relationship has no Id"))?
            .clone();
        if child.local_name() != "Relationship" || relationships.insert(id, child).is_some() {
            return Err(template_error("invalid or duplicate relationship"));
        }
    }
    Ok(relationships)
}

fn relationship_target(relationship: &XmlElement, base: &str, kind: &str) -> RpyxlResult<String> {
    if relationship
        .attributes
        .get("TargetMode")
        .is_some_and(|mode| mode != "Internal")
        || !relationship
            .attributes
            .get("Type")
            .is_some_and(|value| value.rsplit('/').next() == Some(kind))
    {
        return Err(template_error(format!("invalid {kind} relationship")));
    }
    resolve_part(
        base,
        relationship
            .attributes
            .get("Target")
            .ok_or_else(|| template_error("relationship has no Target"))?,
    )
}

fn validate_package_bindings<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    source: &PackageTemplate,
) -> RpyxlResult<()> {
    crate::reader::validate_zip_archive_security(archive)?;
    let mut names = HashSet::new();
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index).map_err(template_error)?;
        if !names.insert(entry.name().to_owned()) {
            return Err(template_error("duplicate ZIP member names"));
        }
    }
    let mut replaced = HashSet::new();
    for part in std::iter::once(&source.styles_path).chain(&source.sheet_parts) {
        if !names.contains(part) || !replaced.insert(part) || part == &source.workbook_path {
            return Err(template_error(
                "missing, duplicate, or overlapping source part paths",
            ));
        }
    }
    let root_rels = package_relationships(archive, "_rels/.rels")?;
    let office = root_rels
        .values()
        .filter(|rel| {
            rel.attributes
                .get("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"))
        })
        .collect::<Vec<_>>();
    if office.len() != 1
        || relationship_target(office[0], "", "officeDocument")? != source.workbook_path
    {
        return Err(template_error(
            "workbook path does not match source relationships",
        ));
    }
    let (parent, filename) = source
        .workbook_path
        .rsplit_once('/')
        .unwrap_or(("", &source.workbook_path));
    let rels_path = if parent.is_empty() {
        format!("_rels/{filename}.rels")
    } else {
        format!("{parent}/_rels/{filename}.rels")
    };
    let rels = package_relationships(archive, &rels_path)?;
    let style_rels = rels
        .values()
        .filter(|rel| {
            rel.attributes
                .get("Type")
                .is_some_and(|kind| kind.ends_with("/styles"))
        })
        .collect::<Vec<_>>();
    if style_rels.len() != 1
        || relationship_target(style_rels[0], &source.workbook_path, "styles")?
            != source.styles_path
    {
        return Err(template_error(
            "styles path does not match source relationships",
        ));
    }
    let workbook = parse_xml(&read_package_part(archive, &source.workbook_path)?)?;
    if !workbook.is_main("workbook") {
        return Err(template_error("invalid source workbook"));
    }
    let epoch = workbook
        .children
        .iter()
        .find(|node| node.is_main("workbookPr"))
        .and_then(|node| node.attributes.get("date1904"))
        .is_some_and(|value| value == "1" || value == "true");
    let active = workbook
        .children
        .iter()
        .find(|node| node.is_main("bookViews"))
        .and_then(|node| {
            node.children
                .iter()
                .find(|node| node.is_main("workbookView"))
        })
        .map_or(Ok(0), |node| node.index("activeTab", 0))?;
    let sheets = workbook
        .children
        .iter()
        .find(|node| node.is_main("sheets"))
        .ok_or_else(|| template_error("missing workbook sheets"))?;
    if epoch != source.date_1904
        || active != source.active_index
        || sheets.children.len() != source.sheet_parts.len()
        || source.sheet_names.len() != source.sheet_parts.len()
    {
        return Err(template_error(
            "source workbook structure does not match template",
        ));
    }
    for ((sheet, name), part) in sheets
        .children
        .iter()
        .zip(&source.sheet_names)
        .zip(&source.sheet_parts)
    {
        let id = sheet
            .expanded_attributes
            .iter()
            .find(|(key, _)| {
                key == "http://schemas.openxmlformats.org/officeDocument/2006/relationships:id"
            })
            .map(|(_, value)| value)
            .ok_or_else(|| template_error("missing worksheet relationship"))?;
        let rel = rels
            .get(id)
            .ok_or_else(|| template_error("unbound worksheet relationship"))?;
        if !sheet.is_main("sheet")
            || sheet.attributes.get("name") != Some(name)
            || relationship_target(rel, &source.workbook_path, "worksheet")? != *part
        {
            return Err(template_error("worksheet bindings do not match template"));
        }
    }
    // Read only the style part, not themes/images/shared strings or other payloads.
    let file = archive
        .by_name(&source.styles_path)
        .map_err(template_error)?;
    let original_styles = crate::styles::Styles::from_reader(file)?;
    if original_styles.source_xml() != Some(source.styles_xml.as_str()) {
        return Err(template_error(
            "stylesheet template does not match source archive",
        ));
    }
    Ok(())
}

struct RowFrame {
    start: String,
    end: String,
}

fn source_sheet_rows(
    xml: &str,
    ws: &Worksheet,
    style_count: usize,
    dxf_count: usize,
) -> RpyxlResult<BTreeMap<u32, RowFrame>> {
    let template = ws
        .xml_template
        .as_ref()
        .ok_or_else(|| template_error("missing worksheet framing"))?;
    let framed = format!(
        "{}{}{}{}</sheetData>{}",
        template.root_start, template.before_data, template.data_start, "", template.after_data
    );
    let frame = parse_xml(&framed)?;
    if !frame.is_main("worksheet") || frame.name != "worksheet" {
        return Err(template_error("unsafe worksheet root"));
    }
    let data = frame
        .children
        .iter()
        .filter(|node| node.is_main("sheetData"))
        .collect::<Vec<_>>();
    if data.len() != 1
        || data[0].name != "sheetData"
        || !data[0].children.is_empty()
        || data[0].has_text()
    {
        return Err(template_error("invalid worksheet data framing"));
    }
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut metadata = quick_xml::Writer::new(Vec::new());
    let mut depth = 0;
    let mut data_depth = None;
    let mut rows = BTreeMap::new();
    loop {
        let event = reader.read_event().map_err(template_error)?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = matches!(event, Event::Empty(_));
                if depth == 1 && element.local_name().as_ref() == b"sheetData" {
                    metadata
                        .write_event(Event::Start(element.borrow()))
                        .map_err(write_error)?;
                    if empty {
                        metadata
                            .write_event(Event::End(element.to_end()))
                            .map_err(write_error)?;
                    } else {
                        data_depth = Some(depth + 1);
                    }
                } else if let Some(data_depth) = data_depth {
                    if (depth == data_depth && element.local_name().as_ref() == b"row")
                        || (depth == data_depth + 1 && element.local_name().as_ref() == b"c")
                    {
                        for attribute in element.attributes() {
                            let attribute = attribute.map_err(template_error)?;
                            if attribute.key.as_ref() == b"s" {
                                let index = attribute
                                    .decode_and_unescape_value(reader.decoder())
                                    .map_err(template_error)?
                                    .parse::<usize>()
                                    .map_err(template_error)?;
                                if index >= style_count {
                                    return Err(template_error(
                                        "source cell/row references an unavailable XF",
                                    ));
                                }
                            }
                        }
                    }
                    if depth == data_depth && element.local_name().as_ref() == b"row" {
                        let mut row = None;
                        for attribute in element.attributes() {
                            let attribute = attribute.map_err(template_error)?;
                            if attribute.key.as_ref() == b"r" {
                                row = Some(
                                    attribute
                                        .decode_and_unescape_value(reader.decoder())
                                        .map_err(template_error)?
                                        .parse::<u32>()
                                        .map_err(template_error)?,
                                );
                            }
                            if attribute.key.as_ref() == b"xmlns"
                                && attribute.value.as_ref() != SPREADSHEET_NS
                            {
                                return Err(template_error(
                                    "unsafe default namespace on source row",
                                ));
                            }
                        }
                        let row = row
                            .filter(|row| *row > 0)
                            .ok_or_else(|| template_error("source row has no valid row index"))?;
                        let mut writer = quick_xml::Writer::new(Vec::new());
                        writer
                            .write_event(Event::Start(element.borrow()))
                            .map_err(write_error)?;
                        let frame = RowFrame {
                            start: String::from_utf8(writer.into_inner())
                                .map_err(template_error)?,
                            end: format!("</{}>", String::from_utf8_lossy(element.name().as_ref())),
                        };
                        if rows.insert(row, frame).is_some() {
                            return Err(template_error("duplicate source row index"));
                        }
                    }
                } else {
                    metadata.write_event(event.borrow()).map_err(write_error)?;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                if depth == 0 {
                    return Err(template_error("unexpected worksheet closing tag"));
                }
                if data_depth == Some(depth) {
                    data_depth = None;
                }
                if data_depth.is_none() {
                    metadata.write_event(event.borrow()).map_err(write_error)?;
                }
                depth -= 1;
            }
            Event::Eof => break,
            Event::DocType(_) => return Err(template_error("DOCTYPE in source worksheet")),
            _ if data_depth.is_none() => {
                metadata.write_event(event.borrow()).map_err(write_error)?
            }
            _ => {}
        }
    }
    let original = parse_xml(&String::from_utf8(metadata.into_inner()).map_err(template_error)?)?;
    validate_sheet_style_references(&original, style_count, dxf_count)?;
    let metadata_keys = |root: &XmlElement| {
        root.children
            .iter()
            .filter(|node| !node.is_main("dimension"))
            .map(XmlElement::key)
            .collect::<Vec<_>>()
    };
    if original.namespace != frame.namespace
        || original.expanded_attributes != frame.expanded_attributes
        || metadata_keys(&original) != metadata_keys(&frame)
    {
        return Err(template_error(
            "worksheet framing does not match source archive",
        ));
    }
    Ok(rows)
}

fn validate_sheet_style_references(
    node: &XmlElement,
    style_count: usize,
    dxf_count: usize,
) -> RpyxlResult<()> {
    for (element, attribute, count) in [
        ("col", "style", style_count),
        ("cfRule", "dxfId", dxf_count),
    ] {
        if node.is_main(element)
            && node.attributes.contains_key(attribute)
            && node.index(attribute, 0)? >= count
        {
            return Err(template_error(format!(
                "source {element} references an unavailable {attribute}"
            )));
        }
    }
    for child in &node.children {
        validate_sheet_style_references(child, style_count, dxf_count)?;
    }
    Ok(())
}

fn generate_preserved_worksheet(
    ws: &Worksheet,
    styles: &StyleTables,
    date_1904: bool,
    rows: &BTreeMap<u32, RowFrame>,
) -> RpyxlResult<String> {
    let template = ws
        .xml_template
        .as_ref()
        .ok_or_else(|| template_error("missing worksheet framing"))?;
    let mut dimension = String::from("<dimension ref=\"");
    write_coordinate(ws.min_row(), ws.min_column(), &mut dimension)?;
    dimension.push(':');
    write_coordinate(ws.max_row(), ws.max_column(), &mut dimension)?;
    dimension.push_str("\"/>");
    let mut xml = String::with_capacity(
        template.root_start.len()
            + template.before_data.len()
            + template.after_data.len()
            + ws.cell_count() * 45
            + 128,
    );
    xml.push_str(&template.root_start);
    // sheetPr precedes dimension in the worksheet schema.
    let prefix = parse_xml(&format!(
        "{}{}{}{}</sheetData>{}",
        template.root_start, template.before_data, template.data_start, "", template.after_data
    ))?;
    let insert = prefix
        .children
        .iter()
        .find(|node| node.is_main("sheetPr"))
        .map_or(0, |node| node.end - template.root_start.len());
    xml.push_str(&template.before_data[..insert]);
    xml.push_str(&dimension);
    xml.push_str(&template.before_data[insert..]);
    xml.push_str(&template.data_start);
    let row_ids = rows
        .keys()
        .copied()
        .chain(ws.cells().keys().map(|&(row, _)| row))
        .collect::<std::collections::BTreeSet<_>>();
    for row in row_ids {
        if let Some(frame) = rows.get(&row) {
            xml.push_str(&frame.start);
        } else {
            xml.push_str(&format!("<row r=\"{row}\">"));
        }
        for (&(r, c), cell) in ws.cells().range((row, 0)..=(row, u32::MAX)) {
            write_workbook_cell(r, c, cell, &mut xml, styles.cell_style(cell), date_1904);
        }
        xml.push_str(rows.get(&row).map_or("</row>", |frame| frame.end.as_str()));
    }
    xml.push_str("</sheetData>");
    xml.push_str(&template.after_data);
    Ok(xml)
}

fn write_preserved_package<W: Write + Seek>(
    writer: W,
    wb: &Workbook,
    source: &PackageTemplate,
    styles: &StyleTables,
) -> RpyxlResult<()> {
    if wb
        .source_styles_xml
        .as_deref()
        .is_some_and(|xml| xml != source.styles_xml)
    {
        return Err(template_error(
            "independent style base does not match the source package",
        ));
    }
    if !parse_xml(&source.styles_xml)?.is_main("styleSheet") {
        return Err(template_error(
            "legacy stylesheet requires explicit generated mode",
        ));
    }
    let mut archive =
        ZipArchive::new(Cursor::new(source.archive.as_ref())).map_err(template_error)?;
    validate_package_bindings(&mut archive, source)?;
    let mut replacements = HashMap::new();
    replacements.insert(source.styles_path.as_str(), styles.to_xml());
    for (sheet, part) in wb.sheets().iter().zip(&source.sheet_parts) {
        let original = read_package_part(&mut archive, part)?;
        let native = styles
            .native
            .as_ref()
            .ok_or_else(|| template_error("missing native style base"))?;
        let rows = source_sheet_rows(
            &original,
            sheet,
            native.source_xf_count,
            native.source_dxf_count,
        )?;
        replacements.insert(
            part.as_str(),
            generate_preserved_worksheet(sheet, styles, wb.date_1904, &rows)?,
        );
    }
    let mut zip = ZipWriter::new(writer);
    zip.set_raw_comment(archive.comment().to_vec().into_boxed_slice());
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let file = archive.by_index_raw(index).map_err(template_error)?;
        if let Some(xml) = replacements.get(file.name()) {
            zip.start_file(file.name(), options).map_err(write_error)?;
            zip.write_all(xml.as_bytes()).map_err(write_error)?;
        } else {
            zip.raw_copy_file(file).map_err(write_error)?;
        }
    }
    zip.finish().map_err(write_error)?;
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
    StyleTables::new(formats.to_vec()).to_xml()
}

const SPREADSHEET_NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn is_xml_character(c: char) -> bool {
    matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
}

fn is_xml_qname(name: &[u8]) -> bool {
    let Ok(name) = std::str::from_utf8(name) else {
        return false;
    };
    let start = |c: char| matches!(c, 'A'..='Z' | '_' | 'a'..='z' | '\u{c0}'..='\u{d6}' | '\u{d8}'..='\u{f6}' | '\u{f8}'..='\u{2ff}' | '\u{370}'..='\u{37d}' | '\u{37f}'..='\u{1fff}' | '\u{200c}'..='\u{200d}' | '\u{2070}'..='\u{218f}' | '\u{2c00}'..='\u{2fef}' | '\u{3001}'..='\u{d7ff}' | '\u{f900}'..='\u{fdcf}' | '\u{fdf0}'..='\u{fffd}' | '\u{10000}'..='\u{effff}');
    let mut parts = name.split(':');
    let valid_part = |part: &str| {
        let mut chars = part.chars();
        chars.next().is_some_and(start)
            && chars.all(|c| start(c) || matches!(c, '-' | '.' | '0'..='9' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}'))
    };
    valid_part(parts.next().unwrap_or_default())
        && parts.next().map_or(true, valid_part)
        && parts.next().is_none()
}

// Bound recursive tree operations and the allocations made for template elements.
const MAX_XML_DEPTH: usize = 128;
const MAX_XML_NODES: usize = 1_000_000;

fn check_xml_limits(
    depth: usize,
    nodes: &mut usize,
    max_depth: usize,
    max_nodes: usize,
) -> RpyxlResult<()> {
    if depth > max_depth {
        return Err(template_error(format!(
            "XML depth exceeds limit of {max_depth}"
        )));
    }
    if *nodes >= max_nodes {
        return Err(template_error(format!(
            "XML element count exceeds limit of {max_nodes}"
        )));
    }
    *nodes += 1;
    Ok(())
}

fn validate_style_fragment(fragment: &str, expected: &str) -> RpyxlResult<()> {
    let invalid = || RpyxlError::ValueError(format!("Invalid rich-style {expected} fragment"));
    if !fragment.chars().all(is_xml_character) {
        return Err(invalid());
    }
    // Validate in the same default namespace as styles.xml: an omitted namespace
    // inherits SpreadsheetML, whereas an explicit xmlns="" must not do so.
    let wrapped = format!(
        "<rpyxl-style-fragment xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{fragment}</rpyxl-style-fragment>"
    );
    let mut reader = NsReader::from_str(&wrapped);
    reader.config_mut().check_comments = true;
    reader.read_resolved_event().map_err(|_| invalid())?;

    #[derive(Clone, Copy, PartialEq)]
    enum Context {
        Standard,
        ExtensionList,
        ExtensionPayload,
    }
    let mut contexts = vec![Context::Standard];
    let mut nodes = 1; // Include the synthetic wrapper in the budget.
    let mut seen_root = false;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(|_| invalid())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if matches!(namespace, ResolveResult::Unknown(_))
                    || !is_xml_qname(element.name().as_ref())
                {
                    return Err(invalid());
                }
                let spreadsheet =
                    matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == SPREADSHEET_NS);
                check_xml_limits(contexts.len() + 1, &mut nodes, MAX_XML_DEPTH, MAX_XML_NODES)?;
                let parent = *contexts.last().ok_or_else(invalid)?;
                if parent != Context::ExtensionPayload && !spreadsheet {
                    return Err(invalid());
                }
                let local_name = element.local_name();
                if contexts.len() == 1 {
                    if seen_root || local_name.as_ref() != expected.as_bytes() {
                        return Err(invalid());
                    }
                    seen_root = true;
                }
                let context = match parent {
                    Context::ExtensionPayload => Context::ExtensionPayload,
                    Context::ExtensionList => {
                        if local_name.as_ref() != b"ext" {
                            return Err(invalid());
                        }
                        Context::ExtensionPayload
                    }
                    Context::Standard if local_name.as_ref() == b"extLst" => Context::ExtensionList,
                    Context::Standard => Context::Standard,
                };
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|_| invalid())?;
                    if !is_xml_qname(attribute.key.as_ref())
                        || attribute.value.contains(&b'<')
                        || (attribute.key.as_ref() != b"xmlns"
                            && !attribute.key.as_ref().starts_with(b"xmlns:")
                            && matches!(
                                reader.resolve_attribute(attribute.key).0,
                                ResolveResult::Unknown(_)
                            ))
                    {
                        return Err(invalid());
                    }
                    let value = attribute
                        .decode_and_unescape_value(decoder)
                        .map_err(|_| invalid())?;
                    if !value.chars().all(is_xml_character) {
                        return Err(invalid());
                    }
                }
                if matches!(event, Event::Start(_)) {
                    contexts.push(context);
                }
            }
            Event::End(_) => {
                contexts.pop().ok_or_else(invalid)?;
                // Only the synthetic closing tag may close the wrapper.
                if contexts.is_empty() && reader.buffer_position() != wrapped.len() as u64 {
                    return Err(invalid());
                }
            }
            Event::Text(text) => {
                if text.as_ref().windows(3).any(|bytes| bytes == b"]]>") {
                    return Err(invalid());
                }
                let text = text.unescape().map_err(|_| invalid())?;
                if !text.chars().all(is_xml_character)
                    || (contexts.last() != Some(&Context::ExtensionPayload)
                        && !text.chars().all(|c| matches!(c, ' ' | '\t' | '\r' | '\n')))
                {
                    return Err(invalid());
                }
            }
            Event::CData(text) if contexts.last() == Some(&Context::ExtensionPayload) => {
                let text = decoder.decode(text.as_ref()).map_err(|_| invalid())?;
                if !text.chars().all(is_xml_character) {
                    return Err(invalid());
                }
            }
            Event::Eof if seen_root && contexts.is_empty() => return Ok(()),
            // Prologs, DTDs and processing instructions are never allowed, even
            // inside extension payloads. Comments are ordinary XML content.
            Event::Comment(_) => {}
            _ => return Err(invalid()),
        }
    }
}

// Byte ranges let us append to native tables without rebuilding unrelated XML.
#[derive(Clone)]
struct XmlElement {
    start: usize,
    open_end: usize,
    close_start: usize,
    end: usize,
    name: String,
    namespace: String,
    attributes: HashMap<String, String>,
    expanded_attributes: Vec<(String, String)>,
    children: Vec<XmlElement>,
    text: Vec<(usize, String)>,
}

impl XmlElement {
    fn has_text(&self) -> bool {
        self.text.iter().any(|(_, text)| !text.trim().is_empty())
    }

    fn push_text(&mut self, text: &str) {
        let position = self.children.len();
        if let Some((last_position, last_text)) = self.text.last_mut() {
            if *last_position == position {
                last_text.push_str(text);
                return;
            }
        }
        self.text.push((position, text.to_owned()));
    }
    fn local_name(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(&self.name)
    }

    fn is_main(&self, name: &str) -> bool {
        self.namespace.as_bytes() == SPREADSHEET_NS && self.local_name() == name
    }

    fn index(&self, name: &str, default: usize) -> RpyxlResult<usize> {
        self.attributes.get(name).map_or(Ok(default), |value| {
            value
                .parse()
                .map_err(|_| template_error(format!("invalid {name} index")))
        })
    }

    fn max_format_id(&self) -> RpyxlResult<usize> {
        let own = if self.namespace.as_bytes() == SPREADSHEET_NS {
            self.index("numFmtId", 163)?
        } else {
            163
        };
        if own > u32::MAX as usize {
            return Err(template_error("invalid number format ID"));
        }
        self.children
            .iter()
            .try_fold(own.max(163), |maximum, child| {
                Ok(maximum.max(child.max_format_id()?))
            })
    }

    fn key(&self) -> String {
        let mut key = String::new();
        self.write_key(&mut key);
        key
    }

    fn write_key(&self, key: &mut String) {
        fn field(key: &mut String, value: &str) {
            key.push_str(itoa::Buffer::new().format(value.len()));
            key.push(':');
            key.push_str(value);
        }

        // Length-prefixed fields are unambiguous without escaping descendant keys.
        // Namespace declarations and insignificant whitespace are not style identity.
        key.push('(');
        field(key, &self.namespace);
        field(key, self.local_name());
        key.push('[');
        for (name, value) in &self.expanded_attributes {
            field(key, name);
            field(key, value);
        }
        key.push(']');
        key.push('[');
        for (position, text) in &self.text {
            if self.namespace.as_bytes() != SPREADSHEET_NS || !text.trim().is_empty() {
                key.push_str(itoa::Buffer::new().format(*position));
                key.push(':');
                field(key, text);
            }
        }
        key.push(']');
        key.push('[');
        for child in &self.children {
            child.write_key(key);
        }
        key.push(']');
        key.push(')');
    }
}

fn parse_xml(xml: &str) -> RpyxlResult<XmlElement> {
    parse_xml_with_limits(xml, MAX_XML_DEPTH, MAX_XML_NODES)
}

fn parse_xml_with_limits(xml: &str, max_depth: usize, max_nodes: usize) -> RpyxlResult<XmlElement> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_comments = true;
    let mut stack: Vec<XmlElement> = Vec::new();
    let mut root = None;
    let mut declaration = false;
    let mut nodes = 0;
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(template_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                check_xml_limits(stack.len() + 1, &mut nodes, max_depth, max_nodes)?;
                if !is_xml_qname(element.name().as_ref()) {
                    return Err(template_error("invalid XML element name"));
                }
                let namespace = match namespace {
                    ResolveResult::Bound(ns) => {
                        String::from_utf8(ns.as_ref().to_vec()).map_err(template_error)?
                    }
                    ResolveResult::Unbound => String::new(),
                    ResolveResult::Unknown(_) => {
                        return Err(template_error("unbound XML namespace"))
                    }
                };
                let end = reader.buffer_position() as usize;
                let mut attributes = HashMap::new();
                let mut expanded_attributes = Vec::new();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(template_error)?;
                    if !is_xml_qname(attribute.key.as_ref()) || attribute.value.contains(&b'<') {
                        return Err(template_error("invalid XML attribute"));
                    }
                    let key = std::str::from_utf8(attribute.key.as_ref())
                        .map_err(template_error)?
                        .to_owned();
                    let value = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(template_error)?
                        .into_owned();
                    if !value.chars().all(is_xml_character) {
                        return Err(template_error("invalid XML character"));
                    }
                    if key != "xmlns" && !key.starts_with("xmlns:") {
                        let (ns, local) = reader.resolve_attribute(attribute.key);
                        let ns = match ns {
                            ResolveResult::Bound(ns) => {
                                String::from_utf8(ns.as_ref().to_vec()).map_err(template_error)?
                            }
                            ResolveResult::Unbound => String::new(),
                            ResolveResult::Unknown(_) => {
                                return Err(template_error("unbound attribute namespace"))
                            }
                        };
                        expanded_attributes.push((
                            format!("{ns}:{}", String::from_utf8_lossy(local.as_ref())),
                            value.clone(),
                        ));
                    }
                    attributes.insert(key, value);
                }
                expanded_attributes.sort();
                let node = XmlElement {
                    start,
                    open_end: end,
                    close_start: end,
                    end,
                    name: std::str::from_utf8(element.name().as_ref())
                        .map_err(template_error)?
                        .to_owned(),
                    namespace,
                    attributes,
                    expanded_attributes,
                    children: Vec::new(),
                    text: Vec::new(),
                };
                if matches!(event, Event::Start(_)) {
                    stack.push(node);
                } else if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err(template_error("multiple XML roots"));
                }
            }
            Event::End(_) => {
                let mut node = stack
                    .pop()
                    .ok_or_else(|| template_error("unexpected closing tag"))?;
                node.close_start = start;
                node.end = reader.buffer_position() as usize;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err(template_error("multiple XML roots"));
                }
            }
            Event::Text(text) => {
                let value = text.unescape().map_err(template_error)?;
                if !value.chars().all(is_xml_character) {
                    return Err(template_error("invalid XML text"));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.push_text(&value);
                } else if !value.trim().is_empty() {
                    return Err(template_error("text outside XML root"));
                }
            }
            Event::CData(text) => {
                let value = reader
                    .decoder()
                    .decode(text.as_ref())
                    .map_err(template_error)?;
                if !value.chars().all(is_xml_character) {
                    return Err(template_error("invalid XML CDATA"));
                }
                stack
                    .last_mut()
                    .ok_or_else(|| template_error("CDATA outside XML root"))?
                    .push_text(&value);
            }
            Event::Decl(decl) if !declaration && root.is_none() && stack.is_empty() => {
                if !crate::styles::utf8_declaration(&decl) {
                    return Err(template_error("source XML must be normalized to UTF-8"));
                }
                declaration = true;
            }
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof if stack.is_empty() => {
                return root.ok_or_else(|| template_error("missing XML root"))
            }
            _ => return Err(template_error("incomplete or unsafe XML template")),
        }
    }
}

fn xml_attribute(value: &str) -> String {
    let mut escaped = String::new();
    write_xml_escaped(value, &mut escaped);
    escaped.replace('\n', "&#10;").replace('\t', "&#9;")
}

fn rewrite_open(xml: &str, node: &XmlElement, updates: &[(&str, String)]) -> RpyxlResult<String> {
    let mut reader = quick_xml::Reader::from_str(&xml[node.start..node.open_end]);
    let event = reader.read_event().map_err(template_error)?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Err(template_error("invalid opening tag")),
    };
    let mut start = BytesStart::new(&node.name);
    for attribute in element.attributes() {
        let attribute = attribute.map_err(template_error)?;
        if !updates
            .iter()
            .any(|(key, _)| attribute.key.as_ref() == key.as_bytes())
        {
            start.push_attribute(attribute);
        }
    }
    for (key, value) in updates {
        start.push_attribute((*key, xml_attribute(value).as_str()));
    }
    let mut writer = quick_xml::Writer::new(Vec::new());
    writer
        .write_event(Event::Start(start))
        .map_err(write_error)?;
    String::from_utf8(writer.into_inner()).map_err(template_error)
}

struct NativeGroup {
    node: Option<XmlElement>,
    original_count: usize,
    additions: Vec<String>,
    ids: HashMap<String, usize>,
}

impl NativeGroup {
    fn from_root(
        root: &XmlElement,
        group: &str,
        record: &str,
        required: bool,
    ) -> RpyxlResult<Self> {
        let mut matching = root.children.iter().filter(|node| node.is_main(group));
        let node = matching.next().cloned();
        if matching.next().is_some()
            || (required && node.as_ref().is_none_or(|node| node.children.is_empty()))
        {
            return Err(template_error(format!(
                "missing or duplicate {group} style table"
            )));
        }
        let mut ids = HashMap::new();
        if let Some(node) = &node {
            for (index, child) in node.children.iter().enumerate() {
                if !child.is_main(record) {
                    return Err(template_error(format!("unsupported record in {group}")));
                }
                ids.entry(child.key()).or_insert(index);
            }
            if node.index("count", node.children.len())? != node.children.len() {
                return Err(template_error(format!("mismatched {group} count")));
            }
        }
        Ok(Self {
            original_count: node.as_ref().map_or(0, |node| node.children.len()),
            node,
            additions: Vec::new(),
            ids,
        })
    }

    fn intern(&mut self, fragment: &str, record: &str) -> RpyxlResult<usize> {
        let wrapped = format!(
            "<wrapper xmlns=\"{}\">{fragment}</wrapper>",
            String::from_utf8_lossy(SPREADSHEET_NS)
        );
        let parsed = parse_xml(&wrapped)?;
        let node = parsed
            .children
            .first()
            .ok_or_else(|| template_error("missing style fragment"))?;
        if parsed.children.len() != 1 || parsed.has_text() || !node.is_main(record) {
            return Err(template_error("invalid native style fragment"));
        }
        let key = node.key();
        if let Some(&index) = self.ids.get(&key) {
            return Ok(index);
        }
        validate_style_fragment(fragment, record)?;
        // Appended native records must be independent of a prefixed stylesheet's
        // default namespace, including namespace overrides on the original group.
        let mut fragment_xml = rewrite_open(
            &wrapped,
            node,
            &[(
                "xmlns",
                String::from_utf8_lossy(SPREADSHEET_NS).into_owned(),
            )],
        )?;
        if node.close_start != node.end {
            fragment_xml.push_str(&wrapped[node.open_end..node.close_start]);
        }
        fragment_xml.push_str(&format!("</{}>", node.name));
        let index = self.original_count + self.additions.len();
        self.ids.insert(key, index);
        self.additions.push(fragment_xml);
        Ok(index)
    }

    fn replacement(&self, xml: &str, group: &str) -> RpyxlResult<String> {
        let count = self.original_count + self.additions.len();
        let mut result = if let Some(node) = &self.node {
            let mut result = rewrite_open(xml, node, &[("count", count.to_string())])?;
            if node.close_start != node.end {
                result.push_str(&xml[node.open_end..node.close_start]);
            }
            result
        } else {
            format!(
                "<{group} xmlns=\"{}\" count=\"{count}\">",
                String::from_utf8_lossy(SPREADSHEET_NS)
            )
        };
        for addition in &self.additions {
            result.push_str(addition);
        }
        result.push_str(&format!(
            "</{}>",
            self.node.as_ref().map_or(group, |node| &node.name)
        ));
        Ok(result)
    }
}

fn native_style_fragment(fragment: &str, record: &str) -> RpyxlResult<String> {
    validate_style_fragment(fragment, record)?;
    let wrapped = format!(
        "<wrapper xmlns=\"{}\">{fragment}</wrapper>",
        String::from_utf8_lossy(SPREADSHEET_NS)
    );
    let parsed = parse_xml(&wrapped)?;
    let node = &parsed.children[0];
    let mut output = rewrite_open(
        &wrapped,
        node,
        &[(
            "xmlns",
            String::from_utf8_lossy(SPREADSHEET_NS).into_owned(),
        )],
    )?;
    if node.close_start != node.end {
        output.push_str(&wrapped[node.open_end..node.close_start]);
    }
    output.push_str(&format!("</{}>", node.name));
    Ok(output)
}

fn clone_native_xf(
    xml: &str,
    node: &XmlElement,
    updates: &[(&str, String)],
    children: &[(&str, String)],
) -> RpyxlResult<String> {
    let rank = |name: &str| match name {
        "alignment" => 0,
        "protection" => 1,
        "extLst" => 2,
        _ => 3,
    };
    let body_end = if node.close_start == node.end {
        node.open_end
    } else {
        node.close_start
    };
    let mut replacements = Vec::new();
    for (name, fragment) in children {
        let existing = node
            .children
            .iter()
            .filter(|child| child.is_main(name))
            .collect::<Vec<_>>();
        if existing.is_empty() {
            let position = node
                .children
                .iter()
                .find(|child| {
                    child.namespace.as_bytes() == SPREADSHEET_NS
                        && rank(child.local_name()) > rank(name)
                })
                .map_or(body_end, |child| child.start);
            replacements.push((position, position, fragment.as_str()));
        } else {
            for (index, child) in existing.into_iter().enumerate() {
                replacements.push((
                    child.start,
                    child.end,
                    if index == 0 { fragment.as_str() } else { "" },
                ));
            }
        }
    }
    replacements.sort_by_key(|(start, end, _)| (*start, *end));
    let mut output = rewrite_open(xml, node, updates)?;
    let mut offset = node.open_end;
    for (start, end, replacement) in replacements {
        output.push_str(&xml[offset..start]);
        output.push_str(replacement);
        offset = end;
    }
    output.push_str(&xml[offset..body_end]);
    output.push_str(&format!("</{}>", node.name));
    Ok(output)
}

fn normalize_legacy_styles(xml: &str, root: &XmlElement) -> RpyxlResult<String> {
    // Namespace-less fixtures are reader-compatible legacy input, not a valid
    // passthrough stylesheet. Normalize only this explicit generated-save case.
    let mut normalized = xml[..root.start].to_owned();
    normalized.push_str(&rewrite_open(
        xml,
        root,
        &[(
            "xmlns",
            String::from_utf8_lossy(SPREADSHEET_NS).into_owned(),
        )],
    )?);
    if root.close_start != root.end {
        normalized.push_str(&xml[root.open_end..root.close_start]);
    }
    normalized.push_str(&format!("</{}>", root.name));
    normalized.push_str(&xml[root.end..]);
    let root = parse_xml(&normalized)?;
    let defaults = StyleTables::new(default_number_formats()).to_xml();
    let default_root = parse_xml(&defaults)?;
    let order = [
        "numFmts",
        "fonts",
        "fills",
        "borders",
        "cellStyleXfs",
        "cellXfs",
        "cellStyles",
        "dxfs",
        "tableStyles",
        "colors",
        "extLst",
    ];
    let mut replacements = Vec::new();
    for group in ["fonts", "fills", "borders", "cellStyleXfs"] {
        let default = default_root
            .children
            .iter()
            .find(|node| node.is_main(group))
            .unwrap();
        if let Some(existing) = root.children.iter().find(|node| node.is_main(group)) {
            if existing.children.is_empty() {
                let mut replacement = rewrite_open(
                    &normalized,
                    existing,
                    &[("count", default.children.len().to_string())],
                )?;
                replacement.push_str(&defaults[default.open_end..default.close_start]);
                replacement.push_str(&format!("</{}>", existing.name));
                replacements.push((existing.start, existing.end, replacement));
            }
        } else {
            let rank = order.iter().position(|name| *name == group).unwrap();
            let position = root
                .children
                .iter()
                .find(|node| {
                    node.namespace.as_bytes() == SPREADSHEET_NS
                        && order
                            .iter()
                            .position(|name| *name == node.local_name())
                            .is_some_and(|index| index > rank)
                })
                .map_or(root.close_start, |node| node.start);
            replacements.push((
                position,
                position,
                defaults[default.start..default.end].to_owned(),
            ));
        }
    }
    replacements.sort_by_key(|(start, end, _)| (*start, *end));
    let mut output = String::new();
    let mut offset = 0;
    for (start, end, replacement) in replacements {
        output.push_str(&normalized[offset..start]);
        output.push_str(&replacement);
        offset = end;
    }
    output.push_str(&normalized[offset..]);
    Ok(output)
}

struct NativeStyleTables {
    xml: String,
    source_xf_count: usize,
    source_dxf_count: usize,
    // A replacement Arc may be shared by cells with different original XF
    // attributes/extensions. Keep provenance in the lookup, not just the Arc.
    cells: HashMap<(Option<usize>, usize), HashMap<String, usize>>,
}

impl NativeStyleTables {
    fn for_workbook(wb: &Workbook, styles_xml: &str) -> RpyxlResult<Self> {
        let parsed = crate::styles::Styles::from_reader(styles_xml.as_bytes())?;
        let original_root = parse_xml(styles_xml)?;
        let normalized;
        let xml = if original_root.namespace.is_empty() && original_root.name == "styleSheet" {
            normalized = normalize_legacy_styles(styles_xml, &original_root)?;
            normalized.as_str()
        } else {
            styles_xml
        };
        let root = parse_xml(xml)?;
        if !root.is_main("styleSheet") {
            return Err(template_error("invalid source stylesheet namespace/root"));
        }
        let mut fonts = NativeGroup::from_root(&root, "fonts", "font", true)?;
        let mut fills = NativeGroup::from_root(&root, "fills", "fill", true)?;
        let mut borders = NativeGroup::from_root(&root, "borders", "border", true)?;
        let mut formats = NativeGroup::from_root(&root, "numFmts", "numFmt", false)?;
        let mut xfs = NativeGroup::from_root(&root, "cellXfs", "xf", true)?;
        let bases = NativeGroup::from_root(&root, "cellStyleXfs", "xf", false)?;
        let dxfs = NativeGroup::from_root(&root, "dxfs", "dxf", false)?;
        if wb.source_styles().len() != xfs.original_count
            || wb.source_styles() != parsed.rich_styles()
        {
            return Err(template_error(
                "source style identities do not match the original XF table",
            ));
        }
        let mut format_ids = HashMap::new();
        let mut used_formats = HashSet::new();
        if let Some(group) = &formats.node {
            for node in &group.children {
                let id = node.index("numFmtId", usize::MAX)?;
                let code = node
                    .attributes
                    .get("formatCode")
                    .ok_or_else(|| template_error("missing formatCode"))?;
                if id == usize::MAX || id > u32::MAX as usize || !used_formats.insert(id) {
                    return Err(template_error("invalid or duplicate number format ID"));
                }
                format_ids.entry(code.clone()).or_insert(id);
            }
        }
        for id in 0..164 {
            if let Some(code) = crate::styles::builtin_number_format(id as u32) {
                format_ids.entry(code.to_owned()).or_insert(id);
            }
        }
        for group in [&xfs, &bases] {
            if let Some(group) = &group.node {
                for xf in &group.children {
                    for (attribute, size) in [
                        ("fontId", fonts.original_count),
                        ("fillId", fills.original_count),
                        ("borderId", borders.original_count),
                    ] {
                        if xf.index(attribute, 0)? >= size {
                            return Err(template_error(format!("out-of-range {attribute}")));
                        }
                    }
                    let id = xf.index("numFmtId", 0)?;
                    if id >= 164 && !used_formats.contains(&id) {
                        return Err(template_error("XF references an unavailable number format"));
                    }
                    if xf.attributes.contains_key("xfId")
                        && xf.index("xfId", 0)? >= bases.original_count
                    {
                        return Err(template_error("XF references an unavailable base style"));
                    }
                }
            }
        }
        // Differential formats can also own IDs. Do not allocate over those IDs
        // even though their records and references stay in the untouched dxfs group.
        let mut next_format = root.max_format_id()? + 1;
        let mut source_style_values = HashMap::new();
        for (index, style) in wb.source_styles().iter().enumerate() {
            source_style_values.entry(style.as_ref()).or_insert(index);
        }
        let mut cells: HashMap<(Option<usize>, usize), HashMap<String, usize>> = HashMap::new();
        for sheet in wb.sheets() {
            for cell in sheet.cells().values() {
                let pointer = cell
                    .rich_style
                    .as_ref()
                    .map_or(0, |style| Arc::as_ptr(style) as usize);
                let provenance = cell.source_xf_index();
                if provenance.is_some_and(|index| index >= xfs.original_count) {
                    return Err(template_error(
                        "cell provenance references an unavailable XF",
                    ));
                }
                let cache_key = (provenance, pointer);
                let code = effective_number_format(&cell.value, cell.number_format.as_deref());
                if !code.chars().all(is_xml_character) {
                    return Err(template_error("invalid character in number format"));
                }
                if cells
                    .get(&cache_key)
                    .is_some_and(|formats| formats.contains_key(code))
                {
                    continue;
                }
                let original = provenance.or_else(|| {
                    cell.rich_style.as_ref().and_then(|style| {
                        wb.source_style_index(style)
                            .or_else(|| source_style_values.get(style.as_ref()).copied())
                    })
                });
                let default_style = RichStyle::default();
                let style = cell.rich_style.as_deref().unwrap_or(&default_style);
                let index = if let Some(index) = original.filter(|&index| {
                    parsed.number_format(index).unwrap_or("General") == code
                        && parsed.rich_styles()[index].as_ref() == style
                }) {
                    index
                } else {
                    let format_id = if let Some(&id) = format_ids.get(code) {
                        id
                    } else {
                        if next_format > u32::MAX as usize {
                            return Err(template_error("number format IDs exhausted"));
                        }
                        let id = next_format;
                        next_format += 1;
                        format_ids.insert(code.to_owned(), id);
                        formats.additions.push(format!(
                            "<numFmt xmlns=\"{}\" numFmtId=\"{id}\" formatCode=\"{}\"/>",
                            String::from_utf8_lossy(SPREADSHEET_NS),
                            xml_attribute(code)
                        ));
                        id
                    };
                    let fragment = if let Some(index) = original {
                        let node = &xfs.node.as_ref().unwrap().children[index];
                        let previous = parsed.rich_styles()[index].as_ref();
                        let mut updates = Vec::new();
                        if parsed.number_format(index).unwrap_or("General") != code {
                            updates.push(("numFmtId", format_id.to_string()));
                            updates.push(("applyNumberFormat", "1".into()));
                        }
                        for (current, old, group, record, attribute, apply) in [
                            (
                                &style.font,
                                &previous.font,
                                &mut fonts,
                                "font",
                                "fontId",
                                "applyFont",
                            ),
                            (
                                &style.fill,
                                &previous.fill,
                                &mut fills,
                                "fill",
                                "fillId",
                                "applyFill",
                            ),
                            (
                                &style.border,
                                &previous.border,
                                &mut borders,
                                "border",
                                "borderId",
                                "applyBorder",
                            ),
                        ] {
                            if current != old {
                                let id = current
                                    .as_deref()
                                    .map(|fragment| group.intern(fragment, record))
                                    .transpose()?
                                    .unwrap_or(0);
                                updates.push((attribute, id.to_string()));
                                updates.push((apply, "1".into()));
                            }
                        }
                        let mut children = Vec::new();
                        for (current, old, record, apply) in [
                            (
                                &style.alignment,
                                &previous.alignment,
                                "alignment",
                                "applyAlignment",
                            ),
                            (
                                &style.protection,
                                &previous.protection,
                                "protection",
                                "applyProtection",
                            ),
                        ] {
                            if current != old {
                                // Empty records explicitly reset inherited defaults;
                                // omitting them would reactivate the base XF value.
                                let empty = format!("<{record}/>");
                                children.push((
                                    record,
                                    native_style_fragment(
                                        current.as_deref().unwrap_or(&empty),
                                        record,
                                    )?,
                                ));
                                updates.push((apply, "1".into()));
                            }
                        }
                        let fragment = clone_native_xf(xml, node, &updates, &children)?;
                        // Original prefixes may be declared on the stylesheet or group.
                        // Parse/deduplicate in that exact namespace context below.
                        fragment
                    } else {
                        let style = cell.rich_style.as_deref();
                        let font = style
                            .and_then(|style| style.font.as_deref())
                            .map(|s| fonts.intern(s, "font"))
                            .transpose()?
                            .unwrap_or(0);
                        let fill = style
                            .and_then(|style| style.fill.as_deref())
                            .map(|s| fills.intern(s, "fill"))
                            .transpose()?
                            .unwrap_or(0);
                        let border = style
                            .and_then(|style| style.border.as_deref())
                            .map(|s| borders.intern(s, "border"))
                            .transpose()?
                            .unwrap_or(0);
                        let mut fragment = format!("<xf xmlns=\"{}\" numFmtId=\"{format_id}\" fontId=\"{font}\" fillId=\"{fill}\" borderId=\"{border}\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"", String::from_utf8_lossy(SPREADSHEET_NS));
                        let alignment = style.and_then(|style| style.alignment.as_deref());
                        let protection = style.and_then(|style| style.protection.as_deref());
                        if alignment.is_some() {
                            fragment.push_str(" applyAlignment=\"1\"");
                        }
                        if protection.is_some() {
                            fragment.push_str(" applyProtection=\"1\"");
                        }
                        fragment.push('>');
                        for (fragment_xml, name) in
                            [(alignment, "alignment"), (protection, "protection")]
                        {
                            if let Some(fragment_xml) = fragment_xml {
                                validate_style_fragment(fragment_xml, name)?;
                                fragment.push_str(fragment_xml);
                            }
                        }
                        fragment.push_str("</xf>");
                        fragment
                    };
                    let group = xfs.node.as_ref().unwrap();
                    let context = format!(
                        "{}{}{fragment}</{}></{}>",
                        &xml[root.start..root.open_end],
                        rewrite_open(xml, group, &[])?,
                        group.name,
                        root.name
                    );
                    let context = parse_xml(&context)?;
                    let key = context.children[0].children[0].key();
                    if let Some(&index) = xfs.ids.get(&key) {
                        index
                    } else {
                        let index = xfs.original_count + xfs.additions.len();
                        xfs.ids.insert(key, index);
                        xfs.additions.push(fragment);
                        index
                    }
                };
                cells
                    .entry(cache_key)
                    .or_default()
                    .insert(code.to_owned(), index);
            }
        }
        let mut replacements = Vec::new();
        for (group, name) in [
            (&formats, "numFmts"),
            (&fonts, "fonts"),
            (&fills, "fills"),
            (&borders, "borders"),
            (&xfs, "cellXfs"),
        ] {
            if !group.additions.is_empty() {
                let (start, end) = group.node.as_ref().map_or(
                    (
                        fonts.node.as_ref().unwrap().start,
                        fonts.node.as_ref().unwrap().start,
                    ),
                    |node| (node.start, node.end),
                );
                replacements.push((start, end, group.replacement(xml, name)?));
            }
        }
        replacements.sort_by_key(|(start, end, _)| (*start, *end));
        let mut output = String::with_capacity(xml.len() + 1024);
        let mut offset = 0;
        for (start, end, replacement) in replacements {
            output.push_str(&xml[offset..start]);
            output.push_str(&replacement);
            offset = end;
        }
        output.push_str(&xml[offset..]);
        Ok(Self {
            xml: output,
            source_xf_count: xfs.original_count,
            source_dxf_count: dxfs.original_count,
            cells,
        })
    }

    fn cell_style(&self, cell: &Cell) -> usize {
        let pointer = cell
            .rich_style
            .as_ref()
            .map_or(0, |style| Arc::as_ptr(style) as usize);
        self.cells[&(cell.source_xf_index(), pointer)]
            [effective_number_format(&cell.value, cell.number_format.as_deref())]
    }
}

#[derive(Default)]
struct ComponentTable {
    values: Vec<String>,
    ids: HashMap<String, usize>,
}

impl ComponentTable {
    fn intern(&mut self, fragment: &str, root: &str) -> RpyxlResult<usize> {
        if let Some(&id) = self.ids.get(fragment) {
            return Ok(id);
        }
        validate_style_fragment(fragment, root)?;
        let id = self.values.len();
        self.values.push(fragment.to_owned());
        self.ids.insert(fragment.to_owned(), id);
        Ok(id)
    }

    fn write_xml(&self, root: &str, xml: &mut String) {
        xml.push_str(&format!("<{root} count=\"{}\">", self.values.len()));
        for fragment in &self.values {
            xml.push_str(fragment);
        }
        xml.push_str(&format!("</{root}>"));
    }
}

#[derive(Clone, Default, PartialEq, Eq, Hash)]
struct RichComponents {
    font: usize,
    fill: usize,
    border: usize,
    alignment: Option<usize>,
    protection: Option<usize>,
}

struct StyleTables {
    native: Option<NativeStyleTables>,
    formats: Vec<String>,
    format_ids: HashMap<String, usize>,
    fonts: ComponentTable,
    fills: ComponentTable,
    borders: ComponentTable,
    alignments: ComponentTable,
    protections: ComponentTable,
    rich: Vec<RichComponents>,
    // Pointer lookup avoids hashing shared XML strings for every cell.
    rich_pointers: HashMap<usize, usize>,
    xfs: Vec<(usize, usize)>,
    xf_ids: HashMap<(usize, usize), usize>,
}

impl StyleTables {
    fn new(formats: Vec<String>) -> Self {
        let mut tables = Self {
            native: None,
            format_ids: formats
                .iter()
                .enumerate()
                .map(|(i, s)| (s.clone(), i))
                .collect(),
            xfs: (0..formats.len()).map(|i| (i, 0)).collect(),
            xf_ids: (0..formats.len()).map(|i| ((i, 0), i)).collect(),
            formats,
            fonts: ComponentTable::default(),
            fills: ComponentTable::default(),
            borders: ComponentTable::default(),
            alignments: ComponentTable::default(),
            protections: ComponentTable::default(),
            rich: vec![RichComponents::default()],
            rich_pointers: HashMap::new(),
        };
        tables
            .fonts
            .intern(
                "<font><name val=\"Calibri\"/><family val=\"2\"/><sz val=\"11\"/></font>",
                "font",
            )
            .expect("valid default font");
        tables
            .fills
            .intern("<fill><patternFill patternType=\"none\"/></fill>", "fill")
            .expect("valid default fill");
        tables
            .fills
            .intern(
                "<fill><patternFill patternType=\"gray125\"/></fill>",
                "fill",
            )
            .expect("valid default fill");
        tables
            .borders
            .intern(
                "<border><left/><right/><top/><bottom/><diagonal/></border>",
                "border",
            )
            .expect("valid default border");
        tables
    }

    fn for_workbook(wb: &Workbook) -> RpyxlResult<Self> {
        let mut tables = Self::new(default_number_formats());
        if let Some(xml) = wb.source_styles_xml.as_deref().or_else(|| {
            wb.package_template
                .as_ref()
                .map(|source| source.styles_xml.as_str())
        }) {
            tables.native = Some(NativeStyleTables::for_workbook(wb, xml)?);
            return Ok(tables);
        }
        if !wb.source_styles().is_empty() {
            return Err(template_error(
                "source styles exist but the original style base is unavailable",
            ));
        }
        let mut rich_ids: HashMap<Arc<RichStyle>, usize> = HashMap::new();
        // Keep explicit rich defaults separate from an unstyled cell, but merge
        // rich styles whose optional components resolve to the same table IDs.
        let mut component_ids: HashMap<RichComponents, usize> = HashMap::new();
        for ws in wb.sheets() {
            for cell in ws.cells().values() {
                let rich_id = if let Some(style) = &cell.rich_style {
                    let pointer = Arc::as_ptr(style) as usize;
                    if let Some(&id) = tables.rich_pointers.get(&pointer) {
                        id
                    } else {
                        let id = if let Some(&id) = rich_ids.get(style) {
                            id
                        } else {
                            let components = RichComponents {
                                font: style
                                    .font
                                    .as_deref()
                                    .map(|s| tables.fonts.intern(s, "font"))
                                    .transpose()?
                                    .unwrap_or(0),
                                fill: style
                                    .fill
                                    .as_deref()
                                    .map(|s| tables.fills.intern(s, "fill"))
                                    .transpose()?
                                    .unwrap_or(0),
                                border: style
                                    .border
                                    .as_deref()
                                    .map(|s| tables.borders.intern(s, "border"))
                                    .transpose()?
                                    .unwrap_or(0),
                                alignment: style
                                    .alignment
                                    .as_deref()
                                    .map(|s| tables.alignments.intern(s, "alignment"))
                                    .transpose()?,
                                protection: style
                                    .protection
                                    .as_deref()
                                    .map(|s| tables.protections.intern(s, "protection"))
                                    .transpose()?,
                            };
                            let id =
                                *component_ids.entry(components.clone()).or_insert_with(|| {
                                    let id = tables.rich.len();
                                    tables.rich.push(components);
                                    id
                                });
                            rich_ids.insert(Arc::clone(style), id);
                            id
                        };
                        tables.rich_pointers.insert(pointer, id);
                        id
                    }
                } else {
                    0
                };
                let code = effective_number_format(&cell.value, cell.number_format.as_deref());
                let format_id = if let Some(&id) = tables.format_ids.get(code) {
                    id
                } else {
                    let id = tables.formats.len();
                    tables.formats.push(code.to_owned());
                    tables.format_ids.insert(code.to_owned(), id);
                    id
                };
                let next_xf = tables.xfs.len();
                tables
                    .xf_ids
                    .entry((format_id, rich_id))
                    .or_insert_with(|| {
                        tables.xfs.push((format_id, rich_id));
                        next_xf
                    });
            }
        }
        Ok(tables)
    }

    fn cell_style(&self, cell: &Cell) -> usize {
        if let Some(native) = &self.native {
            return native.cell_style(cell);
        }
        let code = effective_number_format(&cell.value, cell.number_format.as_deref());
        let format_id = self.format_ids[code];
        let rich_id = cell
            .rich_style
            .as_ref()
            .map(|style| self.rich_pointers[&(Arc::as_ptr(style) as usize)])
            .unwrap_or(0);
        self.xf_ids[&(format_id, rich_id)]
    }

    fn to_xml(&self) -> String {
        if let Some(native) = &self.native {
            return native.xml.clone();
        }
        let formats = &self.formats;
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
        xml.push_str("</numFmts>");
        self.fonts.write_xml("fonts", &mut xml);
        self.fills.write_xml("fills", &mut xml);
        self.borders.write_xml("borders", &mut xml);
        xml.push_str("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
        xml.push_str(&format!("<cellXfs count=\"{}\">", self.xfs.len()));
        for &(format, rich) in &self.xfs {
            let id = if format == 0 { 0 } else { 163 + format };
            let components = &self.rich[rich];
            xml.push_str(&format!("<xf numFmtId=\"{id}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\" applyNumberFormat=\"1\"", components.font, components.fill, components.border));
            if rich != 0 {
                xml.push_str(" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"");
            }
            if components.alignment.is_some() {
                xml.push_str(" applyAlignment=\"1\"");
            }
            if components.protection.is_some() {
                xml.push_str(" applyProtection=\"1\"");
            }
            if components.alignment.is_none() && components.protection.is_none() {
                xml.push_str("/>");
            } else {
                xml.push('>');
                if let Some(id) = components.alignment {
                    xml.push_str(&self.alignments.values[id]);
                }
                if let Some(id) = components.protection {
                    xml.push_str(&self.protections.values[id]);
                }
                xml.push_str("</xf>");
            }
        }
        xml.push_str("</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");
        xml
    }
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
    styles: &StyleTables,
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
        let style = styles.cell_style(cell);
        write_workbook_cell(r, c, cell, &mut s, style, date_1904);
    }
    if current_row != 0 {
        s.push_str("</row>\n");
    }

    s.push_str("  </sheetData>\n</worksheet>");
    Ok(s)
}

fn write_workbook_cell(
    row: u32,
    col: u32,
    cell: &Cell,
    out: &mut String,
    style: usize,
    date_1904: bool,
) {
    write_cell_styled_explicit(
        row,
        col,
        &cell.value,
        out,
        style,
        date_1904,
        cell.source_xf_explicit,
    );
}

#[inline]
fn write_cell_open(row: u32, col: u32, style: usize, explicit: bool, out: &mut String) {
    out.push_str("<c r=\"");
    let _ = write_coordinate(row, col, out);
    if style != 0 || explicit {
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
    write_cell_styled_explicit(row, col, val, out, style, date_1904, false);
}

fn write_cell_styled_explicit(
    row: u32,
    col: u32,
    val: &CellValue,
    out: &mut String,
    style: usize,
    date_1904: bool,
    explicit: bool,
) {
    match val {
        CellValue::Empty => {
            if style != 0 || explicit {
                write_cell_open(row, col, style, explicit, out);
                out.push_str("\"/>");
            }
        }
        CellValue::String(s) => {
            write_cell_open(row, col, style, explicit, out);
            out.push_str("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            write_xml_escaped(s, out);
            out.push_str("</t></is></c>");
        }
        CellValue::Int(i) => {
            write_cell_open(row, col, style, explicit, out);
            out.push_str("\" t=\"n\"><v>");
            let mut itoa_buf = itoa::Buffer::new();
            out.push_str(itoa_buf.format(*i));
            out.push_str("</v></c>");
        }
        CellValue::Float(f) => {
            write_cell_open(row, col, style, explicit, out);
            out.push_str("\" t=\"n\"><v>");
            let mut ryu_buf = ryu::Buffer::new();
            out.push_str(ryu_buf.format(*f));
            out.push_str("</v></c>");
        }
        CellValue::Bool(b) => {
            write_cell_open(row, col, style, explicit, out);
            out.push_str("\" t=\"b\"><v>");
            out.push(if *b { '1' } else { '0' });
            out.push_str("</v></c>");
        }
        CellValue::Date(s) | CellValue::DateTime(s) => {
            if let Some(serial) = temporal_to_excel_serial(s, date_1904) {
                write_cell_open(row, col, style, explicit, out);
                out.push_str("\" t=\"n\"><v>");
                let mut ryu_buf = ryu::Buffer::new();
                out.push_str(ryu_buf.format(serial));
                out.push_str("</v></c>");
            } else {
                write_cell_open(row, col, style, explicit, out);
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
            write_cell_open(row, col, style, explicit, out);
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
    fn test_xml_keys_ignore_prefixes_attribute_order_and_main_whitespace() {
        let first = parse_xml(&format!(
            r#"<font xmlns="{}" xmlns:a="urn:attributes" a:custom="yes" flag="1">
                <name val="Calibri"/>
            </font>"#,
            String::from_utf8_lossy(SPREADSHEET_NS)
        ))
        .unwrap();
        let second = parse_xml(&format!(
            r#"<s:font xmlns:s="{}" xmlns:b="urn:attributes" flag="1" b:custom="yes"><s:name val="Calibri"/></s:font>"#,
            String::from_utf8_lossy(SPREADSHEET_NS)
        ))
        .unwrap();
        assert_eq!(first.key(), second.key());
    }

    #[test]
    fn test_xml_keys_distinguish_fields_namespaces_text_positions_and_child_order() {
        for (first, second) in [
            (r#"<x a="a:b" b="c"/>"#, r#"<x a="a" b="b:c"/>"#),
            (r#"<x a="[]():"/>"#, r#"<x a="[]()"/>"#),
            (r#"<x xmlns="urn:one"/>"#, r#"<x xmlns="urn:two"/>"#),
            ("<x> <y/></x>", "<x><y/> </x>"),
            ("<x> </x>", "<x/>"),
            ("<x><a/><b/></x>", "<x><b/><a/></x>"),
        ] {
            assert_ne!(
                parse_xml(first).unwrap().key(),
                parse_xml(second).unwrap().key(),
                "{first} and {second}"
            );
        }
    }

    #[test]
    fn test_xml_keys_grow_linearly_for_nested_extension_payloads() {
        let mut lengths = Vec::new();
        for depth in [8, 16, 32, 64] {
            let xml = format!(
                r#"<root xmlns="urn:extension">{}<leaf a="&quot;\[]:()">&quot;\</leaf>{}</root>"#,
                "<node>".repeat(depth),
                "</node>".repeat(depth)
            );
            let key = parse_xml(&xml).unwrap().key();
            assert!(key.len() < xml.len() * 8);
            assert_eq!(key.matches('\\').count(), 2);
            lengths.push(key.len());
        }
        for lengths in lengths.windows(3) {
            assert_eq!(lengths[2] - lengths[1], 2 * (lengths[1] - lengths[0]));
        }
    }

    #[test]
    fn test_xml_parser_preserves_namespaced_element_byte_ranges() {
        let xml = r#"<s:root xmlns:s="urn:root"><s:child/></s:root>"#;
        let root = parse_xml(xml).unwrap();
        assert_eq!(root.namespace, "urn:root");
        assert_eq!(&xml[root.start..root.end], xml);
        assert_eq!(
            &xml[root.start..root.open_end],
            r#"<s:root xmlns:s="urn:root">"#
        );
        assert_eq!(&xml[root.close_start..root.end], "</s:root>");
        let child = &root.children[0];
        assert_eq!(child.namespace, "urn:root");
        assert_eq!(&xml[child.start..child.end], "<s:child/>");
        assert_eq!(child.close_start, child.end);
        assert!(parse_xml("<unbound:root/>").is_err());
        assert!(parse_xml(r#"<root unbound:attribute="value"/>"#).is_err());
    }

    #[test]
    fn test_xml_parser_limits_depth_for_start_and_empty_elements() {
        for leaf in ["<x/>", "<x></x>"] {
            for depth in [MAX_XML_DEPTH, MAX_XML_DEPTH + 1] {
                let xml = format!(
                    "{}{leaf}{}",
                    "<x>".repeat(depth - 1),
                    "</x>".repeat(depth - 1)
                );
                let result = parse_xml(&xml);
                if depth == MAX_XML_DEPTH {
                    let root = result.unwrap();
                    assert!(!root.key().is_empty());
                } else {
                    assert!(result.err().unwrap().to_string().contains("XML depth"));
                }
            }
        }
    }

    #[test]
    fn test_xml_parser_limits_total_elements_including_empty_siblings() {
        let xml = "<root><a/><b></b><c/></root>";
        assert!(parse_xml_with_limits(xml, MAX_XML_DEPTH, 4).is_ok());
        let error = parse_xml_with_limits(xml, MAX_XML_DEPTH, 3).err().unwrap();
        assert!(error.to_string().contains("XML element count"));
        assert!(parse_xml_with_limits("<root/>", MAX_XML_DEPTH, 0).is_err());
    }

    #[test]
    fn test_style_fragment_limits_extension_payload_depth() {
        for depth in [MAX_XML_DEPTH, MAX_XML_DEPTH + 1] {
            // The wrapper, font, extLst and ext account for four levels.
            let fragment = format!(
                r#"<font xmlns:x="urn:extension"><extLst><ext uri="private">{}{}</ext></extLst></font>"#,
                "<x:node>".repeat(depth - 4),
                "</x:node>".repeat(depth - 4)
            );
            let result = validate_style_fragment(&fragment, "font");
            if depth == MAX_XML_DEPTH {
                result.unwrap();
            } else {
                assert!(result.unwrap_err().to_string().contains("XML depth"));
            }
        }
    }

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
