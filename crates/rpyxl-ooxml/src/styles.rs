use std::collections::HashMap;
use std::io::Read;
use std::sync::Arc;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Writer};

use rpyxl_core::cell::RichStyle;
use rpyxl_core::error::{RpyxlError, RpyxlResult};

const SPREADSHEET_NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
// RichStyle owns its component Strings, so a small table can expand once per XF.
const MAX_RESOLVED_STYLE_BYTES: usize = 64 * 1024 * 1024;

#[derive(Debug, Default, Clone)]
pub struct Styles {
    custom_num_fmts: HashMap<u32, String>,
    cell_xfs: Vec<u32>, // style index -> numFmtId
    rich_styles: Vec<Arc<RichStyle>>,
    source_xml: Option<String>,
    pub date_1904: bool,
}

impl Styles {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn from_reader<R: Read>(mut reader: R) -> RpyxlResult<Self> {
        let mut source = Vec::new();
        reader
            .read_to_end(&mut source)
            .map_err(|err| style_error(err.to_string()))?;
        let mut xml_reader = NsReader::from_reader(source.as_slice());
        xml_reader.config_mut().trim_text(false);
        let mut styles = Self::default();
        let mut fonts = Vec::new();
        let mut fills = Vec::new();
        let mut borders = Vec::new();
        let mut base_xfs = Vec::new();
        let mut cell_xfs = Vec::new();
        let mut path: Vec<Vec<u8>> = Vec::new();
        let mut namespaces = Vec::new();
        let mut namespace_counts = Vec::new();
        let mut capture: Option<(Vec<u8>, usize, Writer<Vec<u8>>)> = None;
        let mut utf8 = true;
        let mut roots = 0;
        let mut legacy_namespace = false;

        loop {
            let (namespace, event) = xml_reader
                .read_resolved_event()
                .map_err(|err| style_error(err.to_string()))?;
            let unbound = matches!(&namespace, ResolveResult::Unbound);
            let spreadsheet = matches!(namespace,
                ResolveResult::Bound(namespace) if namespace.as_ref() == SPREADSHEET_NS);
            if matches!(event, Event::DocType(_)) {
                return Err(RpyxlError::XmlSecurityViolation(
                    "DTD / DOCTYPE declaration is forbidden in styles.xml".to_string(),
                ));
            }
            if let Some((_, _, writer)) = capture.as_mut() {
                writer
                    .write_event(event.borrow())
                    .map_err(|err| style_error(err.to_string()))?;
            }
            match &event {
                Event::Decl(decl) => utf8 &= utf8_declaration(decl),
                Event::Start(e) | Event::Empty(e) => {
                    let empty = matches!(event, Event::Empty(_));
                    if path.is_empty() {
                        roots += 1;
                        // A namespace-less root selects legacy mode without accepting mixed namespaces.
                        legacy_namespace = unbound && e.name().as_ref() == b"styleSheet";
                    }
                    let accepted_namespace = if legacy_namespace {
                        unbound
                    } else {
                        spreadsheet
                    };
                    // Foreign elements cannot match any ancestry segment in either mode.
                    let name = if accepted_namespace {
                        e.local_name().as_ref().to_vec()
                    } else {
                        Vec::new()
                    };
                    let namespace_count = namespaces.len();
                    for attr in e.attributes() {
                        let attr = attr.map_err(|err| style_error(err.to_string()))?;
                        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:")
                        {
                            namespaces.push((attr.key.as_ref().to_vec(), attr.value.into_owned()));
                        }
                    }
                    let parent = path.last().map(Vec::as_slice).unwrap_or_default();
                    let in_table = path.len() == 2 && path[0] == b"styleSheet";
                    let in_xf = path.len() == 3
                        && path[0] == b"styleSheet"
                        && matches!(path[1].as_slice(), b"cellXfs" | b"cellStyleXfs")
                        && parent == b"xf";
                    if in_table && name == b"numFmt" && parent == b"numFmts" {
                        let mut id = None;
                        let mut code = None;
                        for attr in e.attributes() {
                            let attr = attr.map_err(|err| style_error(err.to_string()))?;
                            let value = attr
                                .decode_and_unescape_value(xml_reader.decoder())
                                .map_err(|err| style_error(err.to_string()))?;
                            match attr.key.as_ref() {
                                b"numFmtId" => id = value.parse::<u32>().ok(),
                                b"formatCode" => code = Some(value.into_owned()),
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(code)) = (id, code) {
                            styles.custom_num_fmts.insert(id, code);
                        }
                    }
                    if in_table && name == b"xf" && matches!(parent, b"cellXfs" | b"cellStyleXfs") {
                        let xf = StyleXf::parse(e)?;
                        if parent == b"cellXfs" {
                            cell_xfs.push(xf);
                        } else {
                            base_xfs.push(xf);
                        }
                    }
                    let component = (in_table
                        && matches!(
                            (parent, name.as_slice()),
                            (b"fonts", b"font") | (b"fills", b"fill") | (b"borders", b"border")
                        ))
                        || (in_xf && matches!(name.as_slice(), b"alignment" | b"protection"));
                    if capture.is_none() && component {
                        let mut start = e.to_owned();
                        // Move inherited bindings onto the fragment root, retaining local overrides.
                        for (key, value) in namespaces.iter().rev() {
                            if !start
                                .attributes()
                                .flatten()
                                .any(|attr| attr.key.as_ref() == key)
                            {
                                start.push_attribute((key.as_slice(), value.as_slice()));
                            }
                        }
                        let mut writer = Writer::new(Vec::new());
                        writer
                            .write_event(if empty {
                                Event::Empty(start)
                            } else {
                                Event::Start(start)
                            })
                            .map_err(|err| style_error(err.to_string()))?;
                        capture = Some((name.clone(), path.len(), writer));
                    }
                    if !empty {
                        namespace_counts.push(namespace_count);
                        path.push(name);
                    } else {
                        namespaces.truncate(namespace_count);
                    }
                }
                Event::End(_) => {
                    path.pop();
                    if let Some(count) = namespace_counts.pop() {
                        namespaces.truncate(count);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            let complete = capture
                .as_ref()
                .is_some_and(|(_, depth, _)| path.len() == *depth);
            if complete {
                let (kind, _, writer) = capture.take().unwrap();
                let bytes = writer.into_inner();
                // quick-xml reads legacy encodings too; fragments must always be UTF-8 Strings.
                let fragment = xml_reader
                    .decoder()
                    .decode(&bytes)
                    .map_err(|err| style_error(err.to_string()))?
                    .into_owned();
                match kind.as_slice() {
                    b"font" => fonts.push(fragment),
                    b"fill" => fills.push(fragment),
                    b"border" => borders.push(fragment),
                    _ => {
                        let group = path.iter().rev().nth(1).map(Vec::as_slice);
                        let xf = if group == Some(b"cellXfs") {
                            cell_xfs.last_mut()
                        } else {
                            base_xfs.last_mut()
                        };
                        if let Some(xf) = xf {
                            if kind == b"alignment" {
                                xf.alignment = Some(fragment);
                            } else {
                                xf.protection = Some(fragment);
                            }
                        }
                    }
                }
            }
        }
        if !path.is_empty() || roots != 1 {
            return Err(style_error(
                "Incomplete or multiple styles.xml roots".into(),
            ));
        }
        let mut remaining_style_bytes = MAX_RESOLVED_STYLE_BYTES;
        for xf in cell_xfs {
            let base = xf.base.and_then(|id| base_xfs.get(id as usize));
            let id = |own, inherited, apply| effective(own, inherited, apply).unwrap_or(0);
            styles.cell_xfs.push(id(
                xf.num_fmt,
                base.and_then(|b| b.num_fmt),
                xf.apply_num_fmt,
            ));
            let font = fonts.get(id(xf.font, base.and_then(|b| b.font), xf.apply_font) as usize);
            let fill = fills.get(id(xf.fill, base.and_then(|b| b.fill), xf.apply_fill) as usize);
            let border =
                borders.get(id(xf.border, base.and_then(|b| b.border), xf.apply_border) as usize);
            let alignment = effective(
                xf.alignment.as_ref(),
                base.and_then(|b| b.alignment.as_ref()),
                xf.apply_alignment,
            );
            let protection = effective(
                xf.protection.as_ref(),
                base.and_then(|b| b.protection.as_ref()),
                xf.apply_protection,
            );
            // Check the aggregate before cloning any selected component, including inherited XML.
            // Include the RichStyle, Arc counters and per-XF index/handle storage as well as text.
            for bytes in std::iter::once(
                std::mem::size_of::<RichStyle>()
                    + 2 * std::mem::size_of::<usize>()
                    + std::mem::size_of::<Arc<RichStyle>>()
                    + std::mem::size_of::<u32>(),
            )
            .chain(
                [font, fill, border, alignment, protection]
                    .into_iter()
                    .flatten()
                    .map(String::len),
            ) {
                remaining_style_bytes = remaining_style_bytes.checked_sub(bytes).ok_or_else(|| {
                    RpyxlError::XmlSecurityViolation(format!(
                        "Resolved styles.xml allocation exceeds {MAX_RESOLVED_STYLE_BYTES} byte budget"
                    ))
                })?;
            }
            // Keep one identity per original XF, even when its resolved components are equal.
            styles.rich_styles.push(Arc::new(RichStyle {
                font: font.cloned(),
                fill: fill.cloned(),
                border: border.cloned(),
                alignment: alignment.cloned(),
                protection: protection.cloned(),
            }));
        }
        styles.source_xml = Some(if utf8 {
            String::from_utf8(source).map_err(|err| style_error(err.to_string()))?
        } else {
            let decoded = xml_reader
                .decoder()
                .decode(&source)
                .map_err(|err| style_error(err.to_string()))?;
            normalize_source_xml(&decoded)?
        });
        Ok(styles)
    }

    pub fn rich_style(&self, index: usize) -> Option<Arc<RichStyle>> {
        self.rich_styles.get(index).cloned()
    }

    pub fn rich_styles(&self) -> &[Arc<RichStyle>] {
        &self.rich_styles
    }

    /// Retained stylesheet for generated saves; legacy input is normalized to UTF-8.
    pub fn source_xml(&self) -> Option<&str> {
        self.source_xml.as_deref()
    }

    pub fn number_format(&self, style_idx: usize) -> Option<&str> {
        let id = *self.cell_xfs.get(style_idx)?;
        self.custom_num_fmts
            .get(&id)
            .map(String::as_str)
            .or_else(|| builtin_number_format(id))
    }

    pub fn is_date_format(&self, style_idx: usize) -> bool {
        if let Some(code) = self.number_format(style_idx) {
            return is_date_format_code(code);
        }
        matches!(
            self.cell_xfs.get(style_idx),
            Some(14..=22 | 27..=36 | 45..=47 | 50..=58)
        )
    }

    pub fn is_duration_format(&self, style_idx: usize) -> bool {
        self.number_format(style_idx)
            .is_some_and(|code| format_fields(code).2)
    }

    pub fn date_string(&self, serial: f64) -> Option<String> {
        excel_to_date_string(serial, self.date_1904)
    }

    pub fn date_string_for_style(&self, serial: f64, _style_idx: usize) -> Option<String> {
        let mut value = self.date_string(serial)?;
        // Numeric Excel dates are datetimes even when displayed with a date-only
        // format. Standalone times already contain a colon and remain times.
        if !value.contains(':') {
            value.push_str(" 00:00:00");
        }
        Some(value)
    }

    pub fn excel_to_date_string(serial: f64) -> Option<String> {
        excel_to_date_string(serial, false)
    }
}

pub(crate) fn utf8_declaration(decl: &BytesDecl<'_>) -> bool {
    match decl.encoding() {
        None => true,
        Some(Ok(encoding)) => encoding.eq_ignore_ascii_case(b"UTF-8"),
        Some(Err(_)) => false,
    }
}

// Retained source is used by generated saves, not as a raw-byte worksheet template.
// Rewrite declarations too: UTF-8 bytes with a legacy declaration would be decoded twice.
fn normalize_source_xml(source: &str) -> RpyxlResult<String> {
    let mut reader = quick_xml::Reader::from_str(source);
    let mut writer = Writer::new(Vec::with_capacity(source.len()));
    loop {
        let event = reader
            .read_event()
            .map_err(|err| style_error(err.to_string()))?;
        match event {
            Event::Eof => break,
            Event::Decl(decl) => {
                let version = decl.version().map_err(|err| style_error(err.to_string()))?;
                let version =
                    std::str::from_utf8(&version).map_err(|err| style_error(err.to_string()))?;
                let standalone = decl
                    .standalone()
                    .transpose()
                    .map_err(|err| style_error(err.to_string()))?;
                let standalone = standalone
                    .as_deref()
                    .map(std::str::from_utf8)
                    .transpose()
                    .map_err(|err| style_error(err.to_string()))?;
                writer.write_event(Event::Decl(BytesDecl::new(
                    version,
                    Some("UTF-8"),
                    standalone,
                )))
            }
            event => writer.write_event(event),
        }
        .map_err(|err| style_error(err.to_string()))?;
    }
    String::from_utf8(writer.into_inner()).map_err(|err| style_error(err.to_string()))
}

fn style_error(message: String) -> RpyxlError {
    RpyxlError::IoError(format!("Error parsing styles.xml: {message}"))
}

fn effective<T>(own: Option<T>, inherited: Option<T>, apply: Option<bool>) -> Option<T> {
    if apply == Some(false) && inherited.is_some() {
        inherited
    } else {
        own.or(inherited)
    }
}

#[derive(Default)]
struct StyleXf {
    base: Option<u32>,
    num_fmt: Option<u32>,
    font: Option<u32>,
    fill: Option<u32>,
    border: Option<u32>,
    alignment: Option<String>,
    protection: Option<String>,
    apply_num_fmt: Option<bool>,
    apply_font: Option<bool>,
    apply_fill: Option<bool>,
    apply_border: Option<bool>,
    apply_alignment: Option<bool>,
    apply_protection: Option<bool>,
}

impl StyleXf {
    fn parse(element: &BytesStart<'_>) -> RpyxlResult<Self> {
        let mut xf = Self::default();
        for attr in element.attributes() {
            let attr = attr.map_err(|err| style_error(err.to_string()))?;
            let value = std::str::from_utf8(&attr.value).unwrap_or_default();
            let id = value.parse::<u32>().ok();
            let apply = match value {
                "1" | "true" => Some(true),
                "0" | "false" => Some(false),
                _ => None,
            };
            match attr.key.as_ref() {
                b"xfId" => xf.base = id,
                b"numFmtId" => xf.num_fmt = id,
                b"fontId" => xf.font = id,
                b"fillId" => xf.fill = id,
                b"borderId" => xf.border = id,
                b"applyNumberFormat" => xf.apply_num_fmt = apply,
                b"applyFont" => xf.apply_font = apply,
                b"applyFill" => xf.apply_fill = apply,
                b"applyBorder" => xf.apply_border = apply,
                b"applyAlignment" => xf.apply_alignment = apply,
                b"applyProtection" => xf.apply_protection = apply,
                _ => {}
            }
        }
        Ok(xf)
    }
}

fn excel_to_date_string(serial: f64, date_1904: bool) -> Option<String> {
    if !serial.is_finite() || !(-700_000.0..3_000_000.0).contains(&serial) {
        return None;
    }
    let mut days = serial.floor() as i64;
    let millis = ((serial - serial.floor()) * 86_400_000.0).round_ties_even() as i64;
    // Like Excel/openpyxl, serials below one are standalone times unless
    // millisecond rounding carries into the following day.
    if (0.0..1.0).contains(&serial) && millis < 86_400_000 {
        let time = NaiveTime::from_hms_milli_opt(
            (millis / 3_600_000) as u32,
            (millis / 60_000 % 60) as u32,
            (millis / 1000 % 60) as u32,
            (millis % 1000) as u32,
        )?;
        return Some(time.format("%H:%M:%S%.f").to_string());
    }
    // Excel's fictitious 1900-02-29 shares 1900-02-28's representation.
    if !date_1904 && serial > 0.0 && serial < 60.0 {
        days += 1;
    }
    let base = if date_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1)?
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30)?
    };
    let date = base.checked_add_signed(Duration::days(days))?;
    let dt = NaiveDateTime::new(date, NaiveTime::MIN)
        .checked_add_signed(Duration::milliseconds(millis))?;
    if dt.time() == NaiveTime::MIN {
        Some(dt.format("%Y-%m-%d").to_string())
    } else {
        Some(dt.format("%Y-%m-%d %H:%M:%S%.f").to_string())
    }
}

pub(crate) fn builtin_number_format(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);(\"$\"#,##0)",
        6 => "\"$\"#,##0_);[Red](\"$\"#,##0)",
        7 => "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        8 => "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        41 => "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

/// Recognizes date/time and elapsed-time format codes, ignoring literals and directives.
pub fn is_date_format_code(code: &str) -> bool {
    let (date, time, _) = format_fields(code);
    date || time
}

fn format_fields(code: &str) -> (bool, bool, bool) {
    let mut fields = (false, false, false);
    // Only the first section determines date conversion. Ignore literal text,
    // escaped characters, spacing directives, colors, conditions and locales.
    let lower = code.to_ascii_lowercase();
    let mut chars = lower.chars();
    while let Some(c) = chars.next() {
        match c {
            ';' => break,
            '"' => {
                for c in chars.by_ref() {
                    if c == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => {
                let bracket: String = chars.by_ref().take_while(|&c| c != ']').collect();
                if matches!(bracket.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
                    fields.1 = true;
                    fields.2 = true;
                }
            }
            'y' | 'd' | 'm' => fields.0 = true,
            'h' | 's' => fields.1 = true,
            _ => {}
        }
    }
    fields
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_excel_date_conversion() {
        assert_eq!(
            Styles::excel_to_date_string(44927.0).as_deref(),
            Some("2023-01-01")
        );
        assert_eq!(
            Styles::excel_to_date_string(44927.5).as_deref(),
            Some("2023-01-01 12:00:00")
        );
        assert_eq!(
            Styles::excel_to_date_string(0.5).as_deref(),
            Some("12:00:00")
        );
        assert_eq!(
            Styles::excel_to_date_string(59.0).as_deref(),
            Some("1900-02-28")
        );
        assert_eq!(
            Styles::excel_to_date_string(60.0).as_deref(),
            Some("1900-02-28")
        );
        assert_eq!(
            Styles::excel_to_date_string(1.0).as_deref(),
            Some("1900-01-01")
        );
        assert_eq!(Styles::excel_to_date_string(f64::NAN), None);
        assert_eq!(
            Styles::excel_to_date_string(44927.999999999).as_deref(),
            Some("2023-01-02")
        );
        assert_eq!(
            Styles::excel_to_date_string(0.123456789).as_deref(),
            Some("02:57:46.667")
        );
        let styles = Styles {
            date_1904: true,
            ..Styles::default()
        };
        assert_eq!(styles.date_string(1.0).as_deref(), Some("1904-01-02"));
    }

    #[test]
    fn test_format_literals_and_attributes() {
        let xml = br#"<styleSheet><numFmts><numFmt numFmtId="164" formatCode="0.00 &quot;USD&quot;"/></numFmts><cellXfs><xf numFmtId="164"></xf><xf numFmtId="14"/></cellXfs></styleSheet>"#;
        let styles = Styles::from_reader(&xml[..]).unwrap();
        assert_eq!(styles.number_format(0), Some("0.00 \"USD\""));
        assert!(!styles.is_date_format(0));
        assert!(styles.is_date_format(1));
        assert!(!is_date_format_code(r#"0 \m "days" [Red]"#));
        assert!(is_date_format_code("[hh]:mm:ss"));
    }
}
