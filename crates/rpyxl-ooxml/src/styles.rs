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
}

impl Styles {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn from_reader<R: Read>(reader: R) -> RpyxlResult<Self> {
        let mut xml_reader = Reader::from_reader(std::io::BufReader::new(reader));
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut custom_num_fmts = HashMap::new();
        let mut cell_xfs = Vec::new();
        let mut in_cell_xfs = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"cellXfs" {
                        in_cell_xfs = true;
                    }
                }
                Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                    b"numFmt" => {
                        let mut num_fmt_id = None;
                        let mut format_code = None;

                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"numFmtId" {
                                if let Ok(val_str) = std::str::from_utf8(&attr.value) {
                                    num_fmt_id = val_str.parse::<u32>().ok();
                                }
                            } else if attr.key.local_name().as_ref() == b"formatCode" {
                                if let Ok(val_str) = std::str::from_utf8(&attr.value) {
                                    format_code = Some(val_str.to_string());
                                }
                            }
                        }

                        if let (Some(id), Some(code)) = (num_fmt_id, format_code) {
                            custom_num_fmts.insert(id, code);
                        }
                    }
                    b"xf" if in_cell_xfs => {
                        let mut num_fmt_id = 0;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"numFmtId" {
                                if let Ok(val_str) = std::str::from_utf8(&attr.value) {
                                    num_fmt_id = val_str.parse::<u32>().unwrap_or(0);
                                }
                            }
                        }
                        cell_xfs.push(num_fmt_id);
                    }
                    _ => {}
                },
                Ok(Event::End(ref e)) => {
                    if e.local_name().as_ref() == b"cellXfs" {
                        in_cell_xfs = false;
                    }
                }
                Ok(Event::DocType(_)) => {
                    return Err(RpyxlError::XmlSecurityViolation(
                        "DTD / DOCTYPE declaration is forbidden in styles.xml".to_string(),
                    ));
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(RpyxlError::IoError(format!(
                        "Error parsing styles.xml: {:?}",
                        err
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(Self {
            custom_num_fmts,
            cell_xfs,
        })
    }

    pub fn is_date_format(&self, style_idx: usize) -> bool {
        let num_fmt_id = match self.cell_xfs.get(style_idx) {
            Some(&id) => id,
            None => return false,
        };

        // Standard Excel Date/Time format IDs
        match num_fmt_id {
            14..=22 | 27..=36 | 45..=47 | 50..=58 => true,
            _ => {
                if let Some(code) = self.custom_num_fmts.get(&num_fmt_id) {
                    is_date_format_code(code)
                } else {
                    false
                }
            }
        }
    }

    pub fn excel_to_date_string(serial: f64) -> Option<String> {
        if serial < 0.0 || serial > 2_958_465.0 {
            return None;
        }

        let base_date = NaiveDate::from_ymd_opt(1899, 12, 30)?;
        let whole_days = serial.floor() as i64;
        let frac = serial - (whole_days as f64);

        let date = base_date.checked_add_signed(Duration::days(whole_days))?;
        let total_seconds = (frac * 86400.0).round() as u32;

        if total_seconds == 0 {
            Some(date.format("%Y-%m-%d").to_string())
        } else {
            let hours = (total_seconds / 3600) % 24;
            let minutes = (total_seconds % 3600) / 60;
            let seconds = total_seconds % 60;
            let time = NaiveTime::from_hms_opt(hours, minutes, seconds)?;
            let dt = NaiveDateTime::new(date, time);
            Some(dt.format("%Y-%m-%d %H:%M:%S").to_string())
        }
    }
}

fn is_date_format_code(code: &str) -> bool {
    let lower = code.to_ascii_lowercase();
    let mut in_quote = false;
    let mut in_bracket = false;

    for c in lower.chars() {
        match c {
            '"' => in_quote = !in_quote,
            '[' => in_bracket = true,
            ']' => in_bracket = false,
            'y' | 'd' | 'h' | 's' if !in_quote && !in_bracket => return true,
            'm' if !in_quote && !in_bracket => return true,
            _ => {}
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_excel_date_conversion() {
        // Excel 44927 is 2023-01-01
        assert_eq!(
            Styles::excel_to_date_string(44927.0),
            Some("2023-01-01".to_string())
        );

        // With time 44927.5 is 2023-01-01 12:00:00
        assert_eq!(
            Styles::excel_to_date_string(44927.5),
            Some("2023-01-01 12:00:00".to_string())
        );
    }
}
