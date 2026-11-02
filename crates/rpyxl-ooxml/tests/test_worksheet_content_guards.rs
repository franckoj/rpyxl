use std::io::{Cursor, Read, Write};
use std::sync::Arc;

use rpyxl_core::cell::CellValue;
use rpyxl_core::package::PackageTemplate;
use rpyxl_core::workbook::Workbook;
use rpyxl_core::worksheet::Worksheet;
use rpyxl_ooxml::writer::{
    save_workbook, save_workbook_generated, write_workbook, write_workbook_generated,
};
use rpyxl_ooxml::{parse_worksheet_into, SharedStringsTable, Styles, XlsxPackage};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const VALUES: &str =
    r#"<row r="1"><c r="A1"><v>11</v></c><c r="B1"><v>22</v></c><c r="C1"><v>33</v></c></row>"#;

fn worksheet(data: &str, after: &str, namespace: bool) -> String {
    let namespace = if namespace {
        format!(r#" xmlns="{NS}""#)
    } else {
        String::new()
    };
    format!(
        r#"<worksheet{namespace} xmlns:e="urn:foreign"><sheetData>{data}</sheetData>{after}</worksheet>"#
    )
}

fn parse(xml: &str) -> Worksheet {
    let mut sheet = Worksheet::new("Data").unwrap();
    parse_worksheet_into(
        xml.as_bytes(),
        &mut sheet,
        &SharedStringsTable::new(),
        &Styles::default(),
        false,
    )
    .unwrap();
    sheet
}

#[test]
fn foreign_and_nested_merge_cells_do_not_clear_values() {
    let impostors = [
        r#"<e:mergeCells><e:mergeCell ref="A1:C1"/></e:mergeCells>"#,
        r#"<mergeCells><e:mergeCell ref="A1:C1"/></mergeCells>"#,
        r#"<e:mergeCells><mergeCell ref="A1:C1"/></e:mergeCells>"#,
        r#"<mergeCell ref="A1:C1"/>"#,
        r#"<mergeCells><e:wrapper><mergeCell ref="A1:C1"/></e:wrapper></mergeCells>"#,
        r#"<mergeCells><mergeCell><mergeCell ref="A1:C1"/></mergeCell></mergeCells>"#,
        r#"<extLst><ext uri="private"><mergeCells><mergeCell ref="A1:C1"/></mergeCells></ext></extLst>"#,
        r#"<mergeCells xmlns="urn:foreign"><mergeCell ref="A1:C1"/></mergeCells>"#,
    ];
    for namespace in [false, true] {
        for impostor in impostors {
            let sheet = parse(&worksheet(VALUES, impostor, namespace));
            assert!(sheet.merged_ranges.is_empty(), "{impostor}");
            assert!(!sheet.unsupported_cell_content, "{impostor}");
            for (col, value) in [(1, 11), (2, 22), (3, 33)] {
                assert_eq!(
                    sheet.get_cell(1, col).unwrap().value,
                    CellValue::Int(value),
                    "namespace={namespace}: {impostor}"
                );
            }
        }
    }
}

#[test]
fn namespace_rebinding_cannot_make_a_fake_merge_genuine() {
    for impostor in [
        format!(
            r#"<mergeCells xmlns="urn:foreign"><mergeCell xmlns="{NS}" ref="A1:C1"/></mergeCells>"#
        ),
        r#"<mergeCells><mergeCell xmlns="" ref="A1:C1"/></mergeCells>"#.to_owned(),
        r#"<mergeCells xmlns=""><mergeCell ref="A1:C1"/></mergeCells>"#.to_owned(),
    ] {
        let sheet = parse(&worksheet(VALUES, &impostor, true));
        assert!(sheet.merged_ranges.is_empty(), "{impostor}");
        assert_eq!(sheet.get_cell(1, 2).unwrap().value, CellValue::Int(22));
        assert_eq!(sheet.get_cell(1, 3).unwrap().value, CellValue::Int(33));
    }
}

#[test]
fn a_foreign_or_nested_worksheet_cannot_supply_merge_ranges() {
    for xml in [
        format!(
            r#"<e:worksheet xmlns:e="urn:foreign" xmlns="{NS}"><mergeCells><mergeCell ref="A1:C1"/></mergeCells></e:worksheet>"#
        ),
        format!(
            r#"<wrapper xmlns="{NS}"><worksheet><mergeCells><mergeCell ref="A1:C1"/></mergeCells></worksheet></wrapper>"#
        ),
        worksheet(
            VALUES,
            r#"<worksheet><mergeCells><mergeCell ref="A1:C1"/></mergeCells></worksheet>"#,
            true,
        ),
    ] {
        assert!(parse(&xml).merged_ranges.is_empty(), "{xml}");
    }
}

#[test]
fn genuine_merges_clear_only_non_anchor_values() {
    for namespace in [false, true] {
        for merge in [
            r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            r#"<mergeCells><mergeCell ref="A1:B1"></mergeCell></mergeCells>"#,
        ] {
            let sheet = parse(&worksheet(VALUES, merge, namespace));
            assert_eq!(sheet.merged_ranges, vec![(1, 1, 1, 2)]);
            assert_eq!(sheet.get_cell(1, 1).unwrap().value, CellValue::Int(11));
            assert_eq!(sheet.get_cell(1, 2).unwrap().value, CellValue::Empty);
            assert_eq!(sheet.get_cell(1, 3).unwrap().value, CellValue::Int(33));
            assert!(!sheet.unsupported_cell_content);
        }
    }
}

#[test]
fn prefixed_spreadsheet_elements_are_resolved_by_namespace() {
    let xml = format!(
        r#"<s:worksheet xmlns:s="{NS}" xmlns:m="{NS}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>11</s:v></s:c><s:c r="B1"><s:v>22</s:v></s:c></s:row></s:sheetData><m:mergeCells><s:mergeCell ref="A1:B1"/></m:mergeCells></s:worksheet>"#
    );
    let sheet = parse(&xml);
    assert_eq!(sheet.merged_ranges, vec![(1, 1, 1, 2)]);
    assert_eq!(sheet.get_cell(1, 1).unwrap().value, CellValue::Int(11));
    assert_eq!(sheet.get_cell(1, 2).unwrap().value, CellValue::Empty);
    assert!(!sheet.unsupported_cell_content);
}

fn unsupported_data() -> Vec<(&'static str, String)> {
    let cells = [
        (
            "shared formula",
            r#"<c r="A1"><f t="shared" si="0" ref="A1:B1">1+1</f><v>2</v></c>"#,
        ),
        (
            "shared formula follower",
            r#"<c r="A1"><f t="shared" si="0"/><v>2</v></c>"#,
        ),
        (
            "array formula",
            r#"<c r="A1"><f t="array" ref="A1:B1">1+1</f><v>2</v></c>"#,
        ),
        (
            "other formula attribute",
            r#"<c r="A1"><f ca="1">1+1</f><v>2</v></c>"#,
        ),
        (
            "foreign formula attribute",
            r#"<c r="A1"><f e:flag="1">1+1</f><v>2</v></c>"#,
        ),
        (
            "rich inline text",
            r#"<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>rich</t></r><r><t> text</t></r></is></c>"#,
        ),
        ("cell metadata", r#"<c r="A1" cm="1"><v>2</v></c>"#),
        ("value metadata", r#"<c r="A1" vm="1"><v>2</v></c>"#),
        ("phonetic flag", r#"<c r="A1" ph="1"><v>2</v></c>"#),
        ("empty cell metadata", r#"<c r="A1" cm="1"/>"#),
        (
            "foreign style attribute",
            r#"<c r="A1" e:s="1"><v>2</v></c>"#,
        ),
        ("empty foreign style attribute", r#"<c r="A1" e:s="1"/>"#),
        (
            "foreign coordinate attribute",
            r#"<c r="A1" e:r="B1"><v>2</v></c>"#,
        ),
        (
            "foreign type attribute",
            r#"<c r="A1" e:t="str"><v>2</v></c>"#,
        ),
        (
            "unknown cell descendant",
            r#"<c r="A1"><v>2</v><extLst><ext uri="private"/></extLst></c>"#,
        ),
        ("foreign value", r#"<c r="A1"><e:v>2</e:v></c>"#),
        ("wrong value ancestry", r#"<c r="A1"><v><v>2</v></v></c>"#),
        (
            "CDATA value",
            r#"<c r="A1" t="str"><v><![CDATA[text]]></v></c>"#,
        ),
    ];
    let mut data: Vec<_> = cells
        .into_iter()
        .map(|(name, cell)| (name, format!(r#"<row r="1">{cell}</row>"#)))
        .collect();
    data.extend([
        ("unknown sheetData descendant", "<unknown/>".to_owned()),
        (
            "foreign row",
            r#"<e:row r="1"><c r="A1"><v>2</v></c></e:row>"#.to_owned(),
        ),
        ("cell without row", r#"<c r="A1"><v>2</v></c>"#.to_owned()),
        (
            "comment in sheetData",
            "<!-- not regenerated -->".to_owned(),
        ),
    ]);
    data
}

#[test]
fn unsupported_cell_content_is_detected_including_empty_elements() {
    for namespace in [false, true] {
        for (name, data) in unsupported_data() {
            let sheet = parse(&worksheet(&data, "", namespace));
            assert!(
                sheet.unsupported_cell_content,
                "{name}, namespace={namespace}"
            );
        }
    }
}

#[test]
fn rich_inline_text_remains_readable_but_is_not_safe_to_regenerate() {
    let data = r#"<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>rich</t></r><r><t xml:space="preserve"> text</t></r></is></c></row>"#;
    let sheet = parse(&worksheet(data, "", true));
    assert!(sheet.unsupported_cell_content);
    assert_eq!(
        sheet.get_cell(1, 1).unwrap().value,
        CellValue::String("rich text".to_owned())
    );
}

#[test]
fn foreign_cell_attributes_never_override_exact_attributes() {
    for cell in [
        r#"<c r="A1" e:r="B1" e:s="7" e:t="str"><v>2</v></c>"#,
        r#"<c r="A1" e:r="B1" e:s="7" e:t="str"/>"#,
    ] {
        let sheet = parse(&worksheet(&format!(r#"<row r="1">{cell}</row>"#), "", true));
        assert!(sheet.unsupported_cell_content);
        let cell = sheet.get_cell(1, 1).unwrap();
        assert_eq!(cell.source_xf_index(), Some(0));
        assert!(!cell.source_xf_explicit);
        assert!(sheet.get_cell(1, 2).is_none());
    }
}

#[test]
fn nested_merge_in_sheet_data_is_unsupported_and_does_not_erase_values() {
    let data = format!(
        r#"{VALUES}<row r="2"><c r="A2"><mergeCells><mergeCell ref="A1:C1"/></mergeCells></c></row>"#
    );
    let sheet = parse(&worksheet(&data, "", true));
    assert!(sheet.unsupported_cell_content);
    assert!(sheet.merged_ranges.is_empty());
    assert_eq!(sheet.get_cell(1, 2).unwrap().value, CellValue::Int(22));
    assert_eq!(sheet.get_cell(1, 3).unwrap().value, CellValue::Int(33));
}

fn load_for_save(xml: &str, preserve_package: bool) -> Workbook {
    let mut seed = Cursor::new(Vec::new());
    write_workbook(&mut seed, &Workbook::new()).unwrap();
    let mut source = ZipArchive::new(Cursor::new(seed.into_inner())).unwrap();
    let mut output = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        output
            .start_file(entry.name(), SimpleFileOptions::default())
            .unwrap();
        if entry.name() == "xl/worksheets/sheet1.xml" {
            output.write_all(xml.as_bytes()).unwrap();
        } else {
            std::io::copy(&mut entry, &mut output).unwrap();
        }
    }
    let archive: Arc<[u8]> = output.finish().unwrap().into_inner().into();
    let mut package =
        XlsxPackage::from_reader(Cursor::new(Arc::clone(&archive)), false, false).unwrap();
    let mut wb = package.to_workbook().unwrap();
    if preserve_package {
        wb.package_template = Some(Arc::new(PackageTemplate {
            archive,
            workbook_path: package.workbook_path.clone(),
            styles_path: package.styles_path.clone().unwrap(),
            styles_xml: package.styles.source_xml().unwrap().to_owned(),
            sheet_parts: package
                .sheets
                .iter()
                .map(|sheet| sheet.target_path.clone())
                .collect(),
            sheet_names: wb.sheetnames(),
            active_index: wb.active_index(),
            date_1904: wb.date_1904,
        }));
    }
    wb
}

#[test]
fn every_save_mode_rejects_unsupported_content_without_touching_output() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("existing.xlsx");
    for (name, data) in unsupported_data() {
        for preserve_package in [false, true] {
            let wb = load_for_save(&worksheet(&data, "", true), preserve_package);
            assert!(wb.sheets()[0].unsupported_cell_content, "{name}");
            for generated in [false, true] {
                let context = format!("{name}, preserve={preserve_package}, generated={generated}");
                let mut output = Cursor::new(b"untouched stream".to_vec());
                let error = if generated {
                    write_workbook_generated(&mut output, &wb)
                } else {
                    write_workbook(&mut output, &wb)
                }
                .unwrap_err()
                .to_string();
                assert!(
                    error.to_ascii_lowercase().contains("unsupported"),
                    "{context}: {error}"
                );
                assert_eq!(output.into_inner(), b"untouched stream", "{context}");

                std::fs::write(&path, b"untouched file").unwrap();
                let result = if generated {
                    save_workbook_generated(&path, &wb)
                } else {
                    save_workbook(&path, &wb)
                };
                assert!(result.is_err(), "{context}");
                assert_eq!(
                    std::fs::read(&path).unwrap(),
                    b"untouched file",
                    "{context}"
                );
            }
        }
    }
}

#[test]
fn plain_supported_cell_content_can_still_be_saved() {
    let data = r#"<row r="1" s="0" customFormat="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="B1" t="inlineStr"><is><t xml:space="preserve"> padded </t></is></c><c r="C1"/><c r="D1" t="b"><v>1</v></c></row>"#;
    for namespace in [false, true] {
        let xml = worksheet(data, "", namespace);
        assert!(!parse(&xml).unsupported_cell_content);
        let wb = load_for_save(&xml, false);
        for generated in [false, true] {
            let mut output = Cursor::new(Vec::new());
            if generated {
                write_workbook_generated(&mut output, &wb).unwrap();
            } else {
                write_workbook(&mut output, &wb).unwrap();
            }
            let mut archive = ZipArchive::new(Cursor::new(output.into_inner())).unwrap();
            let mut saved = String::new();
            archive
                .by_name("xl/worksheets/sheet1.xml")
                .unwrap()
                .read_to_string(&mut saved)
                .unwrap();
            assert!(!parse(&saved).unsupported_cell_content);
            assert_eq!(
                parse(&saved).get_cell(1, 2).unwrap().value,
                CellValue::String(" padded ".to_owned())
            );
        }
    }
}
