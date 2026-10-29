use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::io::Read;

use rpyxl_core::error::{RpyxlError, RpyxlResult};

#[derive(Debug, Default, Clone)]
pub struct SharedStringsTable {
    strings: Vec<String>,
}

impl SharedStringsTable {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn from_reader<R: Read>(reader: R) -> RpyxlResult<Self> {
        let mut xml_reader = Reader::from_reader(std::io::BufReader::new(reader));
        xml_reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut strings = Vec::new();

        let mut in_si = false;
        let mut in_t = false;
        let mut current_string = String::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_string.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                },

                Ok(Event::Text(ref e)) => {
                    if in_t {
                        if let Ok(text) = e.unescape() {
                            current_string.push_str(&text);
                        }
                    }
                }
                Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                    b"t" => {
                        in_t = false;
                    }
                    b"si" => {
                        in_si = false;
                        if strings.len() >= 10_000_000 {
                            return Err(RpyxlError::XmlSecurityViolation(
                                "Shared strings table exceeds 10,000,000 entries limit".to_string(),
                            ));
                        }
                        strings.push(current_string.clone());
                    }
                    _ => {}
                },
                Ok(Event::DocType(_)) => {
                    return Err(RpyxlError::XmlSecurityViolation(
                        "DTD / DOCTYPE declaration is forbidden in sharedStrings.xml".to_string(),
                    ));
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(RpyxlError::IoError(format!(
                        "Error parsing sharedStrings.xml: {:?}",
                        err
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(Self { strings })
    }

    pub fn get(&self, idx: usize) -> Option<&str> {
        self.strings.get(idx).map(|s| s.as_str())
    }

    pub fn len(&self) -> usize {
        self.strings.len()
    }

    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
            <si><t>First</t></si>
            <si><t>Second</t></si>
            <si><r><t>Part1 </t></r><r><t>Part2</t></r></si>
        </sst>"#;

        let table = SharedStringsTable::from_reader(xml.as_bytes()).unwrap();
        assert_eq!(table.len(), 3);
        assert_eq!(table.get(0), Some("First"));
        assert_eq!(table.get(1), Some("Second"));
        assert_eq!(table.get(2), Some("Part1 Part2"));
        assert_eq!(table.get(3), None);
    }
}
