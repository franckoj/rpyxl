//! Opt-in, conservative rust_xlsxwriter backend for new editable workbooks.
//!
//! Cells are visited in the core store's row-major order and copied into standard
//! in-memory worksheets. This creates a second cell store; no memory win is claimed.
//! Constant-memory spilling is deferred because upstream uses panicking temporary
//! file I/O. The completed ZIP is staged with fallible I/O before touching output.
//! Loaded packages and features that cannot be represented faithfully are errors,
//! not an invitation to fall back to the default writer or discard metadata.

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufReader, BufWriter, Seek, SeekFrom, Write};

use quick_xml::events::{BytesText, Event};
use std::path::Path;
use std::sync::Arc;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use rpyxl_core::cell::{Cell, CellValue, RichStyle};
use rpyxl_core::coordinate::{validate_sheet_title, MAX_COLUMN, MAX_ROW};
use rpyxl_core::error::{RpyxlError, RpyxlResult};
use rpyxl_core::workbook::Workbook;
use rust_xlsxwriter::{
    Color, FontScheme, Format, FormatAlign, FormatBorder, FormatPattern, FormatScript,
    FormatUnderline, Formula, Workbook as Backend, Worksheet,
};

use crate::writer::{
    effective_number_format, parse_xml, temporal_to_excel_serial, validate_style_fragment,
    XmlElement,
};

/// Save a new editable workbook, without truncating `path` until serialization succeeds.
/// Final destination I/O failures (for example disk full during copy) are not atomic.
pub fn save_workbook_rust_xlsxwriter<P: AsRef<Path>>(path: P, wb: &Workbook) -> RpyxlResult<()> {
    let mut staged = serialize(wb)?;
    let mut destination = BufWriter::new(File::create(path).map_err(io_error)?);
    std::io::copy(&mut staged, &mut destination).map_err(io_error)?;
    destination.flush().map_err(io_error)
}

/// Write at the caller's current stream position; do not close or truncate it.
/// Validation/serialization errors leave both its contents and position untouched.
/// As with the default writer, destination I/O failures can leave partial output.
pub fn write_workbook_rust_xlsxwriter<W: Write + Seek>(
    mut writer: W,
    wb: &Workbook,
) -> RpyxlResult<()> {
    let mut staged = serialize(wb)?;
    std::io::copy(&mut staged, &mut writer).map_err(io_error)?;
    writer.flush().map_err(io_error)
}

fn unsupported(message: impl std::fmt::Display) -> RpyxlError {
    RpyxlError::ValueError(format!("rust_xlsxwriter backend: {message}"))
}

fn io_error(error: impl std::fmt::Display) -> RpyxlError {
    RpyxlError::IoError(format!("rust_xlsxwriter backend: {error}"))
}

fn base_format() -> Format {
    // Match the native writer's unthemed Calibri 11 default, not the crate's
    // implicit theme color and minor font scheme.
    Format::new()
        .set_font_color(Color::Automatic)
        .set_font_scheme(FontScheme::None)
}

#[derive(Default)]
struct Formats {
    // Workbook cells retain their Arcs throughout serialization, so these
    // pointer identities cannot be recycled while this cache is in use.
    pointers: HashMap<Option<usize>, usize>,
    by_content: HashMap<Option<Arc<RichStyle>>, usize>,
    styles: Vec<CompiledStyle>,
}

struct CompiledStyle {
    base: Format,
    variants: HashMap<String, Format>,
}

impl Formats {
    fn intern(&mut self, style: &Option<Arc<RichStyle>>) -> RpyxlResult<usize> {
        let pointer = style.as_ref().map(|s| Arc::as_ptr(s) as usize);
        if let Some(&index) = self.pointers.get(&pointer) {
            return Ok(index);
        }
        // Hash/compare component XML only on the first encounter of an Arc.
        // Distinct Arcs with equal content share the same validated compilation.
        let index = match self.by_content.entry(style.clone()) {
            std::collections::hash_map::Entry::Occupied(entry) => *entry.get(),
            std::collections::hash_map::Entry::Vacant(entry) => {
                let base = compile_style(style.as_deref())?;
                let index = self.styles.len();
                self.styles.push(CompiledStyle {
                    base,
                    variants: HashMap::new(),
                });
                entry.insert(index);
                index
            }
        };
        self.pointers.insert(pointer, index);
        Ok(index)
    }

    fn get(&self, style: &Option<Arc<RichStyle>>, code: &str) -> &Format {
        let pointer = style.as_ref().map(|s| Arc::as_ptr(s) as usize);
        &self.styles[self.pointers[&pointer]].variants[code]
    }
}

fn preflight(wb: &Workbook) -> RpyxlResult<Formats> {
    if wb.read_only || wb.data_only || wb.date_1904 {
        return Err(unsupported(
            "requires an editable normal workbook using the 1900 epoch",
        ));
    }
    if wb.package_template.is_some()
        || wb.source_styles_xml.is_some()
        || !wb.source_styles().is_empty()
    {
        return Err(unsupported(
            "loaded workbook/package/style provenance is unsupported",
        ));
    }
    if wb.sheet_count() == 0 || wb.active_index() >= wb.sheet_count() {
        return Err(unsupported("workbook must have a valid active worksheet"));
    }
    let mut names = HashSet::new();
    let mut formats = Formats::default();
    let mut format_count = 1; // Include the backend's default XF.
    for sheet in wb.sheets() {
        validate_sheet_title(&sheet.title)?;
        validate_text(&sheet.title)?;
        if sheet.title.contains(['\r', '\n', '\t']) {
            return Err(unsupported(
                "worksheet names containing control whitespace are unsupported",
            ));
        }
        // The backend reserves History (case insensitive) and counts Unicode
        // scalar values for the 31-character limit, just like the core.
        if sheet.title.eq_ignore_ascii_case("History") || !names.insert(sheet.title.to_lowercase())
        {
            return Err(unsupported("duplicate or reserved worksheet name"));
        }
        if sheet.xml_template.is_some()
            || sheet.structural_passthrough_enabled()
            || sheet.source_rows_changed()
            || sheet.unsupported_cell_content
            || !sheet.merged_ranges.is_empty()
        {
            return Err(unsupported(format!(
                "worksheet {:?} has source, merged, or unsupported content",
                sheet.title
            )));
        }
        for (&(row, col), cell) in sheet.cells() {
            if !(1..=MAX_ROW).contains(&row)
                || !(1..=MAX_COLUMN).contains(&col)
                || cell.row != row
                || cell.column != col
            {
                return Err(unsupported("invalid or inconsistent cell coordinates"));
            }
            if cell.source_xf_id.is_some() {
                return Err(unsupported("cell has loaded XF provenance"));
            }
            if cell.comment.is_some() || cell.hyperlink.is_some() {
                return Err(unsupported("comments and hyperlinks are unsupported"));
            }
            validate_value(&cell.value)?;
            if let Some(code) = cell.number_format.as_deref() {
                validate_attribute_text(code)?;
            }
            let code = effective_number_format(&cell.value, cell.number_format.as_deref());
            validate_attribute_text(code)?;
            if code.is_empty() {
                return Err(unsupported("empty number format is unsupported"));
            }
            // Explicit XF-zero flags can also mark newly assigned styles or
            // blank defaults. Actual loaded provenance is rejected above.
            let index = formats.intern(&cell.rich_style)?;
            let compiled = &mut formats.styles[index];
            if !compiled.variants.contains_key(code) {
                format_count += 1;
                if format_count > 64_000 {
                    return Err(unsupported("more than 64000 compiled style combinations"));
                }
                let base = compiled.base.clone();
                compiled.variants.insert(
                    code.to_owned(),
                    if code == "General" {
                        base
                    } else {
                        base.set_num_format(code)
                    },
                );
            }
        }
    }
    Ok(formats)
}

fn serialize(wb: &Workbook) -> RpyxlResult<File> {
    let formats = preflight(wb)?;
    let mut needs_xml_fixup = false;
    let mut backend = Backend::new();
    backend
        .set_default_format(&base_format(), 20, 64)
        .map_err(io_error)?;
    for (index, sheet) in wb.sheets().iter().enumerate() {
        // Do not use constant/low-memory sheets here: their row spilling uses
        // expect() on real temporary-file I/O and can abort the Python process.
        let target = backend.add_worksheet();
        target.set_name(&sheet.title).map_err(io_error)?;
        target.set_selected(index == wb.active_index());
        target.set_active(index == wb.active_index());
        // An empty Formula::result uses the worksheet default. Override the
        // crate's numeric zero so missing caches remain missing on data_only.
        target.set_formula_result_default("");
        for cell in sheet.cells().values() {
            let code = effective_number_format(&cell.value, cell.number_format.as_deref());
            let format = formats.get(&cell.rich_style, code);
            needs_xml_fixup |= match &cell.value {
                CellValue::String(s) => s.contains('\r'),
                CellValue::Formula { cached_value, .. } => {
                    cached_value.is_none()
                        || matches!(cached_value.as_deref(), Some(CellValue::String(s)) if s.contains('\r'))
                }
                _ => false,
            };
            write_cell(target, cell, format)?;
        }
    }
    let mut staged = tempfile::tempfile().map_err(io_error)?;
    backend.save_to_writer(&mut staged).map_err(io_error)?;
    staged.seek(SeekFrom::Start(0)).map_err(io_error)?;
    if needs_xml_fixup {
        return normalize_cell_xml(staged);
    }
    Ok(staged)
}

// Our reader intentionally performs XML decoding, not Excel's _xHHHH_
// decoding. Convert only the crate's generated CR escapes in cell text/cache
// nodes to equivalent XML references. Literal escape tokens are rejected in
// preflight, so this cannot turn user text into an accidental control character.
// An uncached formula also needs its t="str" removed: otherwise our reader
// treats <v></v> as an empty string cache instead of None. Buffer at most one
// cell, never a worksheet or the workbook. Skip the pass for CR-free books
// without uncached formulas.
fn normalize_cell_xml(staged: File) -> RpyxlResult<File> {
    let mut archive = ZipArchive::new(staged).map_err(io_error)?;
    let mut result = tempfile::tempfile().map_err(io_error)?;
    {
        let mut output = ZipWriter::new(&mut result);
        for index in 0..archive.len() {
            let part = archive.by_index(index).map_err(io_error)?;
            let shared_strings = part.name() == "xl/sharedStrings.xml";
            let worksheet =
                part.name().starts_with("xl/worksheets/") && part.name().ends_with(".xml");
            if !shared_strings && !worksheet {
                output.raw_copy_file(part).map_err(io_error)?;
                continue;
            }
            output
                .start_file(
                    part.name(),
                    SimpleFileOptions::default().compression_method(part.compression()),
                )
                .map_err(io_error)?;
            if shared_strings {
                normalize_shared_strings(BufReader::new(part), &mut output)?;
                continue;
            }
            let mut reader = quick_xml::Reader::from_reader(BufReader::new(part));
            let mut writer = quick_xml::Writer::new(&mut output);
            let mut buffer = Vec::new();
            let mut cell = Vec::new();
            loop {
                let event = reader.read_event_into(&mut buffer).map_err(io_error)?;
                if matches!(event, Event::Eof) {
                    break;
                }
                let start_cell = matches!(&event, Event::Start(e) if e.name().as_ref() == b"c");
                if start_cell || !cell.is_empty() {
                    let end_cell = matches!(&event, Event::End(e) if e.name().as_ref() == b"c");
                    cell.push(event.into_owned());
                    if end_cell {
                        write_normalized_cell(&mut writer, &mut cell)?;
                    }
                } else {
                    writer.write_event(event).map_err(io_error)?;
                }
                buffer.clear();
            }
        }
        output.finish().map_err(io_error)?;
    }
    result.seek(SeekFrom::Start(0)).map_err(io_error)?;
    Ok(result)
}

fn normalize_shared_strings<R: std::io::BufRead, W: Write>(input: R, output: W) -> RpyxlResult<()> {
    let mut reader = quick_xml::Reader::from_reader(input);
    let mut writer = quick_xml::Writer::new(output);
    let mut buffer = Vec::new();
    let mut in_text = false;
    loop {
        match reader.read_event_into(&mut buffer).map_err(io_error)? {
            Event::Eof => break,
            Event::Start(mut start) => {
                in_text = start.name().as_ref() == b"t";
                if in_text
                    && !start
                        .attributes()
                        .any(|a| a.is_ok_and(|a| a.key.as_ref() == b"xml:space"))
                {
                    start.push_attribute(("xml:space", "preserve"));
                }
                writer.write_event(Event::Start(start)).map_err(io_error)?;
            }
            Event::Text(text) if in_text => {
                let raw = std::str::from_utf8(text.as_ref()).map_err(io_error)?;
                let replaced = raw.replace("_x000D_", "&#13;");
                writer
                    .write_event(Event::Text(BytesText::from_escaped(&replaced)))
                    .map_err(io_error)?;
            }
            Event::End(end) => {
                in_text = false;
                writer.write_event(Event::End(end)).map_err(io_error)?;
            }
            other => writer.write_event(other).map_err(io_error)?,
        }
        buffer.clear();
    }
    Ok(())
}

fn write_normalized_cell<W: Write>(
    writer: &mut quick_xml::Writer<W>,
    cell: &mut Vec<Event<'static>>,
) -> RpyxlResult<()> {
    let mut has_formula = false;
    let mut has_cache = false;
    let mut in_value = false;
    for event in cell.iter() {
        match event {
            Event::Start(e) if e.name().as_ref() == b"f" => has_formula = true,
            Event::Start(e) if e.name().as_ref() == b"v" => in_value = true,
            Event::End(e) if e.name().as_ref() == b"v" => in_value = false,
            Event::Text(e) if in_value => has_cache |= !e.as_ref().is_empty(),
            _ => {}
        }
    }
    let mut cell_text = false;
    for event in cell.drain(..) {
        match event {
            Event::Start(mut start) => {
                cell_text = matches!(start.name().as_ref(), b"t" | b"v");
                if start.name().as_ref() == b"c" && has_formula && !has_cache {
                    let attributes = start
                        .attributes()
                        .filter_map(|a| match a {
                            Ok(a) if a.key.as_ref() != b"t" => {
                                Some(Ok((a.key.as_ref().to_vec(), a.value.into_owned())))
                            }
                            Ok(_) => None,
                            Err(e) => Some(Err(e)),
                        })
                        .collect::<Result<Vec<_>, _>>()
                        .map_err(io_error)?;
                    start.clear_attributes();
                    for (key, value) in &attributes {
                        start.push_attribute((key.as_slice(), value.as_slice()));
                    }
                }
                if start.name().as_ref() == b"t"
                    && !start
                        .attributes()
                        .any(|a| a.is_ok_and(|a| a.key.as_ref() == b"xml:space"))
                {
                    start.push_attribute(("xml:space", "preserve"));
                }
                writer.write_event(Event::Start(start)).map_err(io_error)?;
            }
            Event::Text(text) if cell_text => {
                let raw = std::str::from_utf8(text.as_ref()).map_err(io_error)?;
                let replaced = raw.replace("_x000D_", "&#13;");
                writer
                    .write_event(Event::Text(BytesText::from_escaped(&replaced)))
                    .map_err(io_error)?;
            }
            Event::End(end) => {
                cell_text = false;
                writer.write_event(Event::End(end)).map_err(io_error)?;
            }
            other => writer.write_event(other).map_err(io_error)?,
        }
    }
    Ok(())
}

fn has_excel_escape(text: &str) -> bool {
    text.as_bytes()
        .windows(7)
        .any(|s| s.starts_with(b"_x") && s[2..6].iter().all(u8::is_ascii_hexdigit) && s[6] == b'_')
}

fn validate_text(text: &str) -> RpyxlResult<()> {
    if !text.chars().all(|c| matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
        return Err(unsupported("text contains an invalid XML character"));
    }
    Ok(())
}

fn validate_attribute_text(text: &str) -> RpyxlResult<()> {
    validate_text(text)?;
    // Upstream's attribute escaper leaves CR and tabs literal. A conforming XML
    // parser normalizes them to spaces, even if the input fragment used &#N;.
    // Conservatively exclude LF as well from these style attributes.
    if text.contains(['\r', '\n', '\t']) {
        return Err(unsupported(
            "font names and number formats containing control whitespace are unsupported",
        ));
    }
    Ok(())
}

fn validate_value(value: &CellValue) -> RpyxlResult<()> {
    match value {
        CellValue::Empty | CellValue::Bool(_) => Ok(()),
        CellValue::Int(n) => {
            // i128 avoids the saturating f64-to-i64 cast hiding i64::MAX loss.
            if (*n as f64) as i128 != i128::from(*n)
                || (*n as f64).to_string().parse::<i64>().ok() != Some(*n)
            {
                Err(unsupported("integer cannot round-trip exactly through the backend's f64/decimal representation"))
            } else {
                Ok(())
            }
        }
        CellValue::Float(n) if !n.is_finite() => Err(unsupported("non-finite float")),
        CellValue::Float(_) => Ok(()),
        CellValue::String(s) => {
            validate_text(s)?;
            if s.is_empty() || has_excel_escape(s) {
                return Err(unsupported("empty strings and literal Excel escape tokens cannot round-trip through this backend and the native reader"));
            }
            if s.encode_utf16().count() > 32_767 {
                return Err(unsupported("string exceeds Excel's 32767-character limit"));
            }
            Ok(())
        }
        CellValue::Date(s) | CellValue::DateTime(s) => {
            temporal_to_excel_serial(s, false)
                .ok_or_else(|| unsupported("invalid or timezone-aware temporal value"))?;
            Ok(())
        }
        CellValue::Formula {
            formula,
            cached_value,
        } => {
            validate_formula(formula)?;
            if let Some(cached) = cached_value.as_deref() {
                match cached {
                    CellValue::Empty | CellValue::Formula { .. } => {
                        return Err(unsupported("empty/nested formula cache is unsupported"));
                    }
                    CellValue::String(s)
                        if s.is_empty()
                            || s.parse::<f64>().is_ok()
                            || matches!(
                                s.as_str(),
                                "TRUE"
                                    | "FALSE"
                                    | "#DIV/0!"
                                    | "#N/A"
                                    | "#NAME?"
                                    | "#NULL!"
                                    | "#NUM!"
                                    | "#REF!"
                                    | "#VALUE!"
                                    | "#GETTING_DATA"
                            ) =>
                    {
                        return Err(unsupported("formula string cache would change type"));
                    }
                    _ => validate_value(cached)?,
                }
            }
            Ok(())
        }
    }
}

// The crate automatically turns modern functions into dynamic arrays and
// rewrites future functions. Keep this first adapter to an explicit legacy
// function subset rather than silently changing a core formula's payload.
fn validate_formula(formula: &str) -> RpyxlResult<()> {
    validate_text(formula)?;
    let text = formula.strip_prefix('=').unwrap_or(formula);
    if text.is_empty()
        || text.starts_with('=')
        || text.encode_utf16().count() > 8192
        || text.contains(['{', '}', '[', ']', '@', '#', '\r', '\n'])
    {
        return Err(unsupported(
            "empty, array, external, or structured formula is unsupported",
        ));
    }
    let bytes = text.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if matches!(bytes[i], b'\'' | b'"') {
            let quote = bytes[i];
            let start = i;
            i += 1;
            loop {
                if i == bytes.len() {
                    return Err(unsupported("unterminated formula string or sheet name"));
                }
                if bytes[i] == quote {
                    i += 1;
                    if i == bytes.len() || bytes[i] != quote {
                        break;
                    }
                }
                i += 1;
            }
            // Upstream recognizes function substrings inside single-quoted
            // sheet names. Reject all quoted parentheses conservatively,
            // including prefixes, escaped quotes, and future-function names.
            if text[start..i].contains('(') {
                return Err(unsupported(
                    "quoted formula tokens containing '(' are unsupported",
                ));
            }
        } else if bytes[i].is_ascii_alphabetic() || bytes[i] == b'_' {
            let start = i;
            while i < bytes.len()
                && (bytes[i].is_ascii_alphanumeric() || matches!(bytes[i], b'_' | b'.'))
            {
                i += 1;
            }
            let mut next = i;
            while next < bytes.len() && bytes[next].is_ascii_whitespace() {
                next += 1;
            }
            if bytes.get(next) == Some(&b'(') {
                let name = text[start..i].to_ascii_uppercase();
                if !matches!(
                    name.as_str(),
                    "SUM"
                        | "AVERAGE"
                        | "MIN"
                        | "MAX"
                        | "COUNT"
                        | "COUNTA"
                        | "COUNTBLANK"
                        | "IF"
                        | "IFERROR"
                        | "AND"
                        | "OR"
                        | "NOT"
                        | "ROUND"
                        | "ROUNDUP"
                        | "ROUNDDOWN"
                        | "ABS"
                        | "INT"
                        | "MOD"
                        | "SUMIF"
                        | "SUMIFS"
                        | "COUNTIF"
                        | "COUNTIFS"
                        | "VLOOKUP"
                        | "HLOOKUP"
                        | "INDEX"
                        | "MATCH"
                        | "CONCATENATE"
                        | "LEFT"
                        | "RIGHT"
                        | "MID"
                        | "LEN"
                        | "TRIM"
                        | "UPPER"
                        | "LOWER"
                        | "TEXT"
                        | "VALUE"
                        | "DATE"
                        | "TIME"
                        | "YEAR"
                        | "MONTH"
                        | "DAY"
                        | "HOUR"
                        | "MINUTE"
                        | "SECOND"
                        | "TODAY"
                        | "NOW"
                        | "ISBLANK"
                        | "ISNUMBER"
                        | "ISTEXT"
                        | "ISERROR"
                        | "TRUE"
                        | "FALSE"
                ) {
                    return Err(unsupported(format!(
                        "formula function {name} is outside the legacy subset"
                    )));
                }
            }
        } else {
            i += 1;
        }
    }
    Ok(())
}

fn cache_result(value: Option<&CellValue>) -> String {
    match value {
        None => String::new(),
        Some(CellValue::Int(n)) => n.to_string(),
        Some(CellValue::Float(n)) => n.to_string(),
        Some(CellValue::Bool(true)) => "TRUE".into(),
        Some(CellValue::Bool(false)) => "FALSE".into(),
        Some(CellValue::String(s)) => s.clone(),
        Some(CellValue::Date(s) | CellValue::DateTime(s)) => temporal_to_excel_serial(s, false)
            .expect("validated temporal")
            .to_string(),
        _ => unreachable!("validated formula cache"),
    }
}

fn write_cell(target: &mut Worksheet, cell: &Cell, format: &Format) -> RpyxlResult<()> {
    let row = cell.row - 1;
    let col = (cell.column - 1) as u16;
    let result = match &cell.value {
        CellValue::Empty => target.write_blank(row, col, format),
        // Explicit APIs, never write(): = and URL-like strings stay literal.
        CellValue::String(s) => target.write_string_with_format(row, col, s, format),
        CellValue::Int(n) => target.write_number_with_format(row, col, *n as f64, format),
        CellValue::Float(n) => target.write_number_with_format(row, col, *n, format),
        CellValue::Bool(b) => target.write_boolean_with_format(row, col, *b, format),
        CellValue::Date(s) | CellValue::DateTime(s) => target.write_number_with_format(
            row,
            col,
            temporal_to_excel_serial(s, false).expect("validated temporal"),
            format,
        ),
        CellValue::Formula {
            formula,
            cached_value,
        } => target.write_formula_with_format(
            row,
            col,
            Formula::new(formula).set_result(cache_result(cached_value.as_deref())),
            format,
        ),
    };
    result.map(|_| ()).map_err(io_error)
}

fn component(fragment: &str, name: &str) -> RpyxlResult<XmlElement> {
    validate_style_fragment(fragment, name)?;
    let mut wrapper = parse_xml(&format!("<root xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{fragment}</root>"))?;
    Ok(wrapper.children.remove(0))
}

fn schema(node: &XmlElement, attrs: &[&str], children: &[&str]) -> RpyxlResult<()> {
    if !node.is_main(node.local_name()) || node.has_text() {
        return Err(unsupported("foreign namespace or text in style"));
    }
    for key in node.attributes.keys() {
        if key != "xmlns" && !key.starts_with("xmlns:") && !attrs.contains(&key.as_str()) {
            return Err(unsupported(format!(
                "unsupported {} attribute {key}",
                node.local_name()
            )));
        }
    }
    let mut seen = HashSet::new();
    for child in &node.children {
        if !child.is_main(child.local_name())
            || !children.contains(&child.local_name())
            || !seen.insert(child.local_name())
        {
            return Err(unsupported(format!(
                "unknown or duplicate {} child {}",
                node.local_name(),
                child.local_name()
            )));
        }
    }
    Ok(())
}

fn attr<'a>(node: &'a XmlElement, name: &str, default: &'a str) -> &'a str {
    node.attributes
        .get(name)
        .map(String::as_str)
        .unwrap_or(default)
}

fn boolean(value: &str) -> RpyxlResult<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(unsupported("invalid style boolean")),
    }
}

fn integer(value: &str, max: u16) -> RpyxlResult<u16> {
    // Python's modeled family/indent/readingOrder may serialize as 2.0/0.0.
    let n: f64 = value
        .parse()
        .map_err(|_| unsupported("invalid style integer"))?;
    if !n.is_finite() || n < 0.0 || n > f64::from(max) || n.fract() != 0.0 {
        return Err(unsupported("style integer out of range"));
    }
    Ok(n as u16)
}

fn rgb(node: &XmlElement) -> RpyxlResult<Color> {
    schema(node, &["rgb"], &[])?;
    let value = attr(node, "rgb", "");
    if value.len() != 8
        || !value.bytes().all(|b| b.is_ascii_hexdigit())
        || !value[..2].eq_ignore_ascii_case("FF")
    {
        return Err(unsupported("only opaque FF-prefixed 8-digit ARGB colors are representable (no theme/tint/indexed colors)"));
    }
    Ok(Color::RGB(
        u32::from_str_radix(&value[2..], 16).map_err(unsupported)?,
    ))
}

fn compile_style(style: Option<&RichStyle>) -> RpyxlResult<Format> {
    let mut format = base_format();
    if let Some(style) = style {
        if let Some(xml) = &style.font {
            format = font(format, &component(xml, "font")?)?;
        }
        if let Some(xml) = &style.fill {
            format = fill(format, &component(xml, "fill")?)?;
        }
        if let Some(xml) = &style.border {
            format = border(format, &component(xml, "border")?)?;
        }
        if let Some(xml) = &style.alignment {
            format = alignment(format, &component(xml, "alignment")?)?;
        }
        if let Some(xml) = &style.protection {
            let node = component(xml, "protection")?;
            schema(&node, &["locked", "hidden"], &[])?;
            if !boolean(attr(&node, "locked", "true"))? {
                format = format.set_unlocked();
            }
            if boolean(attr(&node, "hidden", "false"))? {
                format = format.set_hidden();
            }
        }
    }
    Ok(format)
}

fn font(mut format: Format, node: &XmlElement) -> RpyxlResult<Format> {
    schema(
        node,
        &[],
        &[
            "name",
            "sz",
            "b",
            "i",
            "strike",
            "u",
            "vertAlign",
            "family",
            "charset",
            "scheme",
            "color",
            "outline",
            "shadow",
            "condense",
            "extend",
        ],
    )?;
    // The crate always emits name/size. Missing family/color/scheme must not
    // inherit its implicit family=2, theme=1, scheme=minor defaults.
    format = format
        .set_font_family(0)
        .set_font_scheme(FontScheme::None)
        .set_font_color(Color::Automatic);
    if !node.children.iter().any(|n| n.local_name() == "name")
        || !node.children.iter().any(|n| n.local_name() == "sz")
    {
        return Err(unsupported("font requires explicit name and size"));
    }
    for child in &node.children {
        let tag = child.local_name();
        if tag == "color" {
            format = format.set_font_color(rgb(child)?);
            continue;
        }
        schema(child, &["val"], &[])?;
        if matches!(
            tag,
            "name" | "sz" | "family" | "charset" | "scheme" | "vertAlign"
        ) && !child.attributes.contains_key("val")
        {
            return Err(unsupported(format!("font {tag} requires val")));
        }
        let value = attr(child, "val", if tag == "u" { "single" } else { "1" });
        format = match tag {
            "name" => {
                if value.is_empty() || !child.attributes.contains_key("val") {
                    return Err(unsupported("empty font name"));
                }
                validate_attribute_text(value)?;
                format.set_font_name(value)
            }
            "sz" => {
                let size: f64 = value.parse().map_err(unsupported)?;
                if !size.is_finite() || size <= 0.0 || size > 409.0 {
                    return Err(unsupported("font size outside 0..409"));
                }
                format.set_font_size(size)
            }
            "b" => {
                if boolean(value)? {
                    format.set_bold()
                } else {
                    format
                }
            }
            "i" => {
                if boolean(value)? {
                    format.set_italic()
                } else {
                    format
                }
            }
            "strike" => {
                if boolean(value)? {
                    format.set_font_strikethrough()
                } else {
                    format
                }
            }
            "outline" | "shadow" | "condense" | "extend" => {
                if boolean(value)? {
                    return Err(unsupported(format!("font {tag} is unsupported")));
                }
                format
            }
            "family" => format.set_font_family(integer(value, 5)? as u8),
            "charset" => format.set_font_charset(integer(value, 255)? as u8),
            "scheme" => format.set_font_scheme(match value {
                "none" => FontScheme::None,
                _ => return Err(unsupported("theme font schemes are unsupported")),
            }),
            "u" => format.set_underline(match value {
                "none" => FormatUnderline::None,
                "single" => FormatUnderline::Single,
                "double" => FormatUnderline::Double,
                "singleAccounting" => FormatUnderline::SingleAccounting,
                "doubleAccounting" => FormatUnderline::DoubleAccounting,
                _ => return Err(unsupported("unsupported underline")),
            }),
            "vertAlign" => format.set_font_script(match value {
                "baseline" => FormatScript::None,
                "superscript" => FormatScript::Superscript,
                "subscript" => FormatScript::Subscript,
                _ => return Err(unsupported("unsupported font vertical alignment")),
            }),
            _ => unreachable!("validated font schema"),
        };
    }
    Ok(format)
}

fn fill(mut format: Format, node: &XmlElement) -> RpyxlResult<Format> {
    schema(node, &[], &["patternFill"])?;
    let Some(pattern) = node.children.first() else {
        return Ok(format);
    };
    schema(pattern, &["patternType"], &["fgColor", "bgColor"])?;
    // Limit the initial adapter to empty and solid fills. The crate may swap
    // foreground/background colors for other patterns.
    let solid = match attr(pattern, "patternType", "none") {
        "none" => false,
        "solid" => true,
        _ => return Err(unsupported("only none/solid pattern fills are supported")),
    };
    if !solid && !pattern.children.is_empty() {
        return Err(unsupported("colors on a none fill would be discarded"));
    }
    if solid {
        format = format.set_pattern(FormatPattern::Solid);
        if !pattern.children.iter().any(|n| n.local_name() == "fgColor") {
            return Err(unsupported("solid fill requires foreground RGB"));
        }
        for color in &pattern.children {
            match color.local_name() {
                "fgColor" => format = format.set_foreground_color(rgb(color)?),
                // The crate forces solid-fill background to indexed=64.
                _ => {
                    return Err(unsupported(
                        "explicit solid-fill background is not preserved",
                    ))
                }
            }
        }
    }
    Ok(format)
}

fn border(mut format: Format, node: &XmlElement) -> RpyxlResult<Format> {
    schema(
        node,
        &["outline", "diagonalUp", "diagonalDown"],
        &[
            "left",
            "right",
            "top",
            "bottom",
            "diagonal",
            "start",
            "end",
            "vertical",
            "horizontal",
        ],
    )?;
    if !boolean(attr(node, "outline", "true"))?
        || boolean(attr(node, "diagonalUp", "false"))?
        || boolean(attr(node, "diagonalDown", "false"))?
    {
        return Err(unsupported(
            "non-default border outline or diagonals are unsupported",
        ));
    }
    for side in &node.children {
        schema(side, &["style"], &["color"])?;
        let style = match attr(side, "style", "none") {
            "none" => FormatBorder::None,
            "thin" => FormatBorder::Thin,
            "medium" => FormatBorder::Medium,
            "dashed" => FormatBorder::Dashed,
            "dotted" => FormatBorder::Dotted,
            "thick" => FormatBorder::Thick,
            "double" => FormatBorder::Double,
            "hair" => FormatBorder::Hair,
            "mediumDashed" => FormatBorder::MediumDashed,
            "dashDot" => FormatBorder::DashDot,
            "mediumDashDot" => FormatBorder::MediumDashDot,
            "dashDotDot" => FormatBorder::DashDotDot,
            "mediumDashDotDot" => FormatBorder::MediumDashDotDot,
            "slantDashDot" => FormatBorder::SlantDashDot,
            _ => return Err(unsupported("unknown border style")),
        };
        if style == FormatBorder::None && !side.children.is_empty() {
            return Err(unsupported(
                "color on a border without a line would be lost",
            ));
        }
        let color = side.children.first().map(rgb).transpose()?;
        format = match side.local_name() {
            "left" => {
                format = format.set_border_left(style);
                if let Some(c) = color {
                    format = format.set_border_left_color(c);
                }
                format
            }
            "right" => {
                format = format.set_border_right(style);
                if let Some(c) = color {
                    format = format.set_border_right_color(c);
                }
                format
            }
            "top" => {
                format = format.set_border_top(style);
                if let Some(c) = color {
                    format = format.set_border_top_color(c);
                }
                format
            }
            "bottom" => {
                format = format.set_border_bottom(style);
                if let Some(c) = color {
                    format = format.set_border_bottom_color(c);
                }
                format
            }
            _ if style == FormatBorder::None => format,
            _ => return Err(unsupported("only four-sided visible borders are supported")),
        };
    }
    Ok(format)
}

fn alignment(mut format: Format, node: &XmlElement) -> RpyxlResult<Format> {
    schema(
        node,
        &[
            "horizontal",
            "vertical",
            "textRotation",
            "wrapText",
            "shrinkToFit",
            "indent",
            "relativeIndent",
            "justifyLastLine",
            "readingOrder",
        ],
        &[],
    )?;
    let horizontal = attr(node, "horizontal", "general");
    format = format.set_align(match horizontal {
        "general" => FormatAlign::General,
        "left" => FormatAlign::Left,
        "center" => FormatAlign::Center,
        "right" => FormatAlign::Right,
        "fill" => FormatAlign::Fill,
        "justify" => FormatAlign::Justify,
        "centerContinuous" => FormatAlign::CenterAcross,
        "distributed" => FormatAlign::Distributed,
        _ => return Err(unsupported("unknown horizontal alignment")),
    });
    let vertical = attr(node, "vertical", "bottom");
    format = format.set_align(match vertical {
        "bottom" => FormatAlign::Bottom,
        "top" => FormatAlign::Top,
        "center" => FormatAlign::VerticalCenter,
        "justify" => FormatAlign::VerticalJustify,
        "distributed" => FormatAlign::VerticalDistributed,
        _ => return Err(unsupported("unknown vertical alignment")),
    });
    let wrap = boolean(attr(node, "wrapText", "false"))?;
    let shrink = boolean(attr(node, "shrinkToFit", "false"))?;
    let indent = integer(attr(node, "indent", "0"), 250)?;
    if integer(attr(node, "relativeIndent", "0"), 255)? != 0
        || boolean(attr(node, "justifyLastLine", "false"))?
    {
        return Err(unsupported(
            "relative indent and justifyLastLine are unsupported",
        ));
    }
    if shrink && (wrap || matches!(horizontal, "fill" | "justify" | "distributed")) {
        return Err(unsupported("backend would discard conflicting shrinkToFit"));
    }
    if indent > 0
        && !matches!(horizontal, "left" | "right" | "distributed")
        && !matches!(vertical, "top" | "bottom")
    {
        return Err(unsupported(
            "backend would change alignment to apply indent",
        ));
    }
    // OOXML encodes clockwise rotations 91..180 as the crate's -1..-90.
    let rotation = integer(attr(node, "textRotation", "0"), 255)?;
    let rotation = match rotation {
        0..=90 => rotation as i16,
        91..=180 => 90 - rotation as i16,
        255 => 270,
        _ => return Err(unsupported("invalid text rotation")),
    };
    format = format
        .set_rotation(rotation)
        .set_indent(indent as u8)
        .set_reading_direction(integer(attr(node, "readingOrder", "0"), 2)? as u8);
    if wrap {
        format = format.set_text_wrap();
    }
    if shrink {
        format = format.set_shrink();
    }
    Ok(format)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pointer_cache_reuses_validated_content_without_bypassing_new_style_validation() {
        let shared = Arc::new(RichStyle {
            font: Some("<font><name val=\"Arial\"/><sz val=\"11\"/><b/></font>".into()),
            ..RichStyle::default()
        });
        let equal_content = Arc::new((*shared).clone());
        let mut formats = Formats::default();
        for _ in 0..1000 {
            assert_eq!(formats.intern(&Some(Arc::clone(&shared))).unwrap(), 0);
        }
        assert_eq!(formats.pointers.len(), 1);
        assert_eq!(formats.by_content.len(), 1);
        assert_eq!(formats.styles.len(), 1);
        assert_eq!(
            formats.intern(&Some(Arc::clone(&equal_content))).unwrap(),
            0
        );
        assert_eq!(formats.pointers.len(), 2);
        assert_eq!(formats.by_content.len(), 1);
        assert_eq!(formats.styles.len(), 1);

        let invalid = Some(Arc::new(RichStyle {
            font: Some("<font><name val=\"Arial\"/><sz val=\"11\"/><opaque/></font>".into()),
            ..RichStyle::default()
        }));
        assert!(formats.intern(&invalid).is_err());
        assert!(formats.intern(&invalid).is_err());
        assert_eq!(formats.pointers.len(), 2);
        assert_eq!(formats.by_content.len(), 1);
        assert_eq!(formats.styles.len(), 1);
    }
}
