use std::sync::Arc;

use rpyxl_core::cell::CellValue;
use rpyxl_core::worksheet::Worksheet;
use rpyxl_ooxml::{
    parse_worksheet_into, parse_worksheet_streaming, read_single_cell, SharedStringsTable,
    StreamingRowReader, Styles,
};

#[test]
fn text_fidelity_across_all_reader_paths() {
    let cases = [
        ("  padded  ", "  padded  "),
        ("   ", "   "),
        ("\t padded\t", "\t padded\t"),
        ("line1\nline2", "line1\nline2"),
        ("line1\r\nline2", "line1\nline2"),
        ("line1\rline2", "line1\nline2"),
        ("line1&#13;line2", "line1\rline2"),
        ("line1&#13;\nline2", "line1\r\nline2"),
        ("line1&#10;line2", "line1\nline2"),
        (
            "hÃ©llo æ—¥æœ¬èªž ðŸ˜€ &lt;&amp;&gt;",
            "hÃ©llo æ—¥æœ¬èªž ðŸ˜€ <&>",
        ),
    ];

    for (encoded, expected) in cases {
        for cell_type in ["inlineStr", "str", "s"] {
            let shared_xml = format!("<sst><si><t xml:space=\"preserve\">{encoded}</t></si></sst>");
            let shared = SharedStringsTable::from_reader(shared_xml.as_bytes()).unwrap();
            let content = match cell_type {
                "inlineStr" => format!("<is><t xml:space=\"preserve\">{encoded}</t></is>"),
                "str" => format!("<v>{encoded}</v>"),
                _ => "<v>0</v>".to_string(),
            };
            let xml = format!(
                "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"{cell_type}\">{content}</c></row></sheetData></worksheet>"
            );
            let styles = Styles::default();
            let value = CellValue::String(expected.to_string());
            let context = format!("{cell_type}: {encoded:?}");

            let mut sheet = Worksheet::new("Text").unwrap();
            parse_worksheet_into(xml.as_bytes(), &mut sheet, &shared, &styles, false).unwrap();
            assert_eq!(
                sheet.get_cell(1, 1).map(|cell| &cell.value),
                Some(&value),
                "{context}"
            );

            let rows = parse_worksheet_streaming(
                xml.as_bytes(),
                &shared,
                &styles,
                None,
                None,
                None,
                None,
                false,
            )
            .unwrap();
            assert_eq!(rows, vec![vec![(1, 1, value.clone())]], "{context}");

            let single = read_single_cell(xml.as_bytes(), 1, 1, &shared, &styles, false).unwrap();
            assert_eq!(single, value, "{context}");

            let mut streaming = StreamingRowReader::new(
                xml.as_bytes(),
                Arc::new(shared),
                Arc::new(styles),
                1,
                1,
                1,
                1,
                false,
            );
            assert_eq!(
                streaming.next_xml_row().unwrap(),
                Some((1, vec![(1, value)])),
                "{context}"
            );
            assert_eq!(streaming.next_xml_row().unwrap(), None);
        }
    }
}

#[test]
fn scalar_xml_whitespace_is_still_accepted() {
    let shared = SharedStringsTable::new();
    let styles = Styles::default();
    for (kind, encoded, expected) in [
        ("n", " 42 ", CellValue::Int(42)),
        ("b", " 1 ", CellValue::Bool(true)),
    ] {
        let xml = format!(
            r#"<worksheet><sheetData><row r="1"><c r="A1" t="{kind}"><v>{encoded}</v></c></row></sheetData></worksheet>"#
        );
        assert_eq!(
            read_single_cell(xml.as_bytes(), 1, 1, &shared, &styles, false).unwrap(),
            expected
        );
    }
}
