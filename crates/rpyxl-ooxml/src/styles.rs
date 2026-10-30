use std::collections::HashMap;
use std::io::Read;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use rpyxl_core::error::{RpyxlError, RpyxlResult};

#[derive(Debug, Default, Clone)]
pub struct Styles {
    custom_num_fmts: HashMap<u32, String>,
    cell_xfs: Vec<u32>, // style index -> numFmtId
    pub date_1904: bool,
}

impl Styles {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn from_reader<R: Read>(reader: R) -> RpyxlResult<Self> {
        let mut xml_reader = Reader::from_reader(std::io::BufReader::new(reader));
        xml_reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut styles = Self::default();
        let mut in_cell_xfs = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    match e.local_name().as_ref() {
                        b"cellXfs" => in_cell_xfs = true,
                        b"numFmt" => {
                            let mut id = None;
                            let mut code = None;
                            for attr in e.attributes().flatten() {
                                let value = attr
                                    .decode_and_unescape_value(xml_reader.decoder())
                                    .map_err(|err| {
                                        RpyxlError::IoError(format!("Invalid number format: {err}"))
                                    })?;
                                match attr.key.local_name().as_ref() {
                                    b"numFmtId" => id = value.parse::<u32>().ok(),
                                    b"formatCode" => code = Some(value.into_owned()),
                                    _ => {}
                                }
                            }
                            if let (Some(id), Some(code)) = (id, code) {
                                styles.custom_num_fmts.insert(id, code);
                            }
                        }
                        b"xf" if in_cell_xfs => {
                            let mut id = 0;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"numFmtId" {
                                    id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(0);
                                }
                            }
                            styles.cell_xfs.push(id);
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => {
                    in_cell_xfs = false;
                }
                Ok(Event::DocType(_)) => {
                    return Err(RpyxlError::XmlSecurityViolation(
                        "DTD / DOCTYPE declaration is forbidden in styles.xml".to_string(),
                    ))
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(RpyxlError::IoError(format!(
                        "Error parsing styles.xml: {err}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(styles)
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
