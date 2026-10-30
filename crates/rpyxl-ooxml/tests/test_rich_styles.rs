use std::collections::HashMap;
use std::io::{Cursor, Read};
use std::sync::Arc;

use quick_xml::{events::Event, Reader};
use rpyxl_core::cell::{CellValue, RichStyle};
use rpyxl_core::workbook::Workbook;
use rpyxl_ooxml::writer::{save_workbook, write_workbook};

fn package_xml(wb: &Workbook, path: &str) -> String {
    let mut bytes = Cursor::new(Vec::new());
    write_workbook(&mut bytes, wb).unwrap();
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes.into_inner())).unwrap();
    let mut xml = String::new();
    zip.by_name(path).unwrap().read_to_string(&mut xml).unwrap();
    xml
}

fn attributes(xml: &str, name: &[u8]) -> Vec<HashMap<String, String>> {
    let mut reader = Reader::from_str(xml);
    let mut result = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == name => {
                result.push(
                    e.attributes()
                        .map(|a| {
                            let a = a.unwrap();
                            (
                                String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                                a.decode_and_unescape_value(reader.decoder())
                                    .unwrap()
                                    .into_owned(),
                            )
                        })
                        .collect(),
                );
            }
            Event::Eof => return result,
            _ => {}
        }
    }
}

fn rich_style() -> Arc<RichStyle> {
    Arc::new(RichStyle {
        font: Some("<s:font xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><s:b/><s:name val=\"A &amp; B\"/></s:font>".into()),
        fill: Some("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFF0000\"/></patternFill></fill>".into()),
        border: Some("<border><left style=\"thin\"/></border>".into()),
        alignment: Some("<alignment horizontal=\"center\" wrapText=\"1\"/>".into()),
        protection: Some("<protection locked=\"0\"/>".into()),
    })
}

#[test]
fn shared_styles_and_equal_distinct_arcs_are_deduplicated_across_sheets() {
    let mut wb = Workbook::new();
    let style = rich_style();
    for row in 1..=100 {
        let cell = wb
            .get_sheet_mut("Sheet")
            .unwrap()
            .set_cell_value(row, 1, CellValue::Int(row.into()))
            .unwrap();
        cell.rich_style = Some(Arc::clone(&style));
    }
    let cell = wb
        .create_sheet(Some("Second"), None)
        .unwrap()
        .set_cell_value(1, 1, CellValue::Int(1))
        .unwrap();
    cell.rich_style = Some(Arc::new((*style).clone()));

    let styles = package_xml(&wb, "xl/styles.xml");
    assert_eq!(attributes(&styles, b"fonts")[0]["count"], "2");
    assert_eq!(attributes(&styles, b"fills")[0]["count"], "3");
    assert_eq!(attributes(&styles, b"borders")[0]["count"], "2");
    assert_eq!(attributes(&styles, b"cellXfs")[0]["count"], "5");
    assert_eq!(attributes(&styles, b"alignment").len(), 1);
    assert_eq!(attributes(&styles, b"protection").len(), 1);
    for path in ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"] {
        assert!(attributes(&package_xml(&wb, path), b"c")
            .iter()
            .all(|c| c["s"] == "4"));
    }
}

#[test]
fn rich_styles_combine_with_number_formats_and_temporal_fallbacks() {
    let mut wb = Workbook::new();
    let style = rich_style();
    for (column, value, format) in [
        (1, CellValue::Float(0.25), "0.00%"),
        (2, CellValue::Float(12.5), "#,##0.00"),
        (3, CellValue::Date("2026-10-30".into()), "General"),
        (4, CellValue::DateTime("12:34:56".into()), "General"),
        (
            5,
            CellValue::formula_with_cached("A3", CellValue::Date("2026-10-30".into())),
            "0.00",
        ),
    ] {
        let cell = wb
            .get_sheet_mut("Sheet")
            .unwrap()
            .set_cell_value(1, column, value)
            .unwrap();
        cell.number_format = Some(format.into());
        cell.rich_style = Some(Arc::clone(&style));
    }
    let styles = package_xml(&wb, "xl/styles.xml");
    let cells = attributes(&package_xml(&wb, "xl/worksheets/sheet1.xml"), b"c");
    let native_styles = rpyxl_ooxml::Styles::from_reader(styles.as_bytes()).unwrap();
    let xfs = attributes(&styles, b"xf");
    let formats = attributes(&styles, b"numFmt")
        .into_iter()
        .map(|a| (a["numFmtId"].clone(), a["formatCode"].clone()))
        .collect::<HashMap<_, _>>();
    for (cell, expected) in cells.iter().zip([
        "0.00%",
        "#,##0.00",
        "yyyy-mm-dd",
        "h:mm:ss.000",
        "yyyy-mm-dd",
    ]) {
        // The first xf belongs to cellStyleXfs, not cellXfs.
        let xf = &xfs[cell["s"].parse::<usize>().unwrap() + 1];
        assert_eq!(formats[&xf["numFmtId"]], expected);
        assert_eq!(
            native_styles.number_format(cell["s"].parse().unwrap()),
            Some(expected)
        );
        assert_eq!(xf["fontId"], "1");
    }
    assert_ne!(cells[0]["s"], cells[1]["s"]);
    assert_eq!(cells[2]["s"], cells[4]["s"]);
}

#[test]
fn equivalent_component_combinations_share_an_xf() {
    let mut wb = Workbook::new();
    let styles = [
        RichStyle::default(),
        RichStyle {
            font: Some(
                "<font><name val=\"Calibri\"/><family val=\"2\"/><sz val=\"11\"/></font>".into(),
            ),
            ..RichStyle::default()
        },
        RichStyle {
            alignment: Some("<alignment horizontal=\"center\"/>".into()),
            ..RichStyle::default()
        },
    ];
    for (index, style) in styles.into_iter().enumerate() {
        wb.get_sheet_mut("Sheet")
            .unwrap()
            .set_cell_value(1, index as u32 + 1, CellValue::Int(1))
            .unwrap()
            .rich_style = Some(Arc::new(style));
    }
    let xml = package_xml(&wb, "xl/styles.xml");
    assert_eq!(attributes(&xml, b"fonts")[0]["count"], "1");
    assert_eq!(attributes(&xml, b"cellXfs")[0]["count"], "6");
    let cells = attributes(&package_xml(&wb, "xl/worksheets/sheet1.xml"), b"c");
    assert_eq!(cells[0]["s"], cells[1]["s"]);
    assert_ne!(cells[0]["s"], cells[2]["s"]);
}

#[test]
fn styled_blanks_are_written_even_without_a_value() {
    let mut wb = Workbook::new();
    let ws = wb.get_sheet_mut("Sheet").unwrap();
    ws.get_or_create_cell(1, 1).unwrap().rich_style = Some(rich_style());
    ws.get_or_create_cell(1, 2).unwrap().rich_style = Some(Arc::new(RichStyle::default()));
    ws.get_or_create_cell(1, 3).unwrap();
    ws.set_cell_number_format(1, 4, Some("0.00".into()))
        .unwrap();
    let xml = package_xml(&wb, "xl/worksheets/sheet1.xml");
    let cells = attributes(&xml, b"c");
    assert_eq!(cells.len(), 3);
    assert_eq!(
        cells.iter().map(|c| c["r"].as_str()).collect::<Vec<_>>(),
        ["A1", "B1", "D1"]
    );
    assert!(cells.iter().all(|c| c["s"] != "0"));
    assert!(!xml.contains("<v>"));
}

#[test]
fn spreadsheet_fragments_inherit_the_stylesheet_namespace() {
    let mut wb = Workbook::new();
    let fragment = "<font xml:space=\"preserve\"><!-- retained --><b/><extLst/></font>";
    wb.get_sheet_mut("Sheet")
        .unwrap()
        .get_or_create_cell(1, 1)
        .unwrap()
        .rich_style = Some(Arc::new(RichStyle {
        font: Some(fragment.into()),
        ..RichStyle::default()
    }));
    let styles = package_xml(&wb, "xl/styles.xml");
    assert!(styles.contains(fragment));
    let mut reader = quick_xml::NsReader::from_str(&styles);
    let mut fonts = 0;
    loop {
        let (namespace, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"font" => {
                assert!(
                    matches!(namespace, quick_xml::name::ResolveResult::Bound(ns)
                    if ns.as_ref() == b"http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                );
                fonts += 1;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    assert_eq!(fonts, 2);
}

#[test]
fn extension_payloads_preserve_foreign_namespaces_text_and_cdata() {
    let fragment = r#"<font><b/><extLst><ext uri="urn:vendor:rich-style"><v:payload xmlns:v="urn:vendor:rich-style" xml:lang="en">A &amp; B<!-- payload comment --><v:label><![CDATA[<literal> & not-an-entity;]]></v:label><local xmlns="">local text</local> tail &#x1F600; ]]&gt;</v:payload></ext></extLst></font>"#;
    let mut wb = Workbook::new();
    for column in 1..=2 {
        wb.get_sheet_mut("Sheet")
            .unwrap()
            .get_or_create_cell(1, column)
            .unwrap()
            .rich_style = Some(Arc::new(RichStyle {
            font: Some(fragment.into()),
            ..RichStyle::default()
        }));
    }
    let styles = package_xml(&wb, "xl/styles.xml");
    assert_eq!(styles.matches(fragment).count(), 1);
    assert_eq!(attributes(&styles, b"fonts")[0]["count"], "2");
    assert_eq!(attributes(&styles, b"cellXfs")[0]["count"], "5");
}

#[test]
fn malformed_fragments_fail_before_touching_the_destination() {
    for fragment in [
        "<font><b></font>",
        "<font/ ></font>",
        "<fill/>",
        "<font/><font/>",
        "<font/>trailing",
        "prefix<font/>",
        "<font>",
        "",
        "<!DOCTYPE font [<!ENTITY x 'injected'>]><font/>",
        "<font xmlns=\"urn:wrong\"/>",
        "<font xmlns=''/>",
        "<font><b xmlns=''/></font>",
        "<font xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><b xmlns=''/></font>",
        "<font><extLst xmlns=''><ext uri='urn:test'/></extLst></font>",
        "<font><extLst><ext xmlns='' uri='urn:test'/></extLst></font>",
        "<font/><fill/>",
        "</rpyxl-style-fragment><font/>",
        "<font/></rpyxl-style-fragment><rpyxl-style-fragment>",
        "<font/></rpyxl-style-fragment><!-- hidden closing tag",
        "<?xml version='1.0'?><font/>",
        "<?injected instruction?><font/>",
        "<font><?injected instruction?></font>",
        "<font><extLst><ext uri='urn:test'><?injected instruction?></ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'><!DOCTYPE payload></ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'><v:payload xmlns:v='urn:vendor'>&unknown;</v:payload></ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'><v:payload xmlns:v='urn:vendor'>&#1;</v:payload></ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'><v:payload xmlns:v='urn:vendor' val='&unknown;'/></ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'><v:payload/></ext></extLst></font>",
        "<font><extLst><v:payload xmlns:v='urn:vendor'/></extLst></font>",
        "<font><extLst><ext uri='urn:test'><v:payload xmlns:v='urn:vendor'></v:other></ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'>invalid ]]&broken;</ext></extLst></font>",
        "<font><extLst><ext uri='urn:test'>invalid ]]></ext></extLst></font>",
        "<font><![CDATA[not extension content]]></font>",
        "<font><extLst><ext uri='urn:test'><![CDATA[\u{1}]]></ext></extLst></font>",
        "<x:font/>",
        "<font><b xmlns=\"urn:wrong\"/></font>",
        "<font><name val=\"&unknown;\"/></font>",
        "<font><name val=\"<injection\"/></font>",
        "<font><!-- invalid -- comment --></font>",
        "<font>\u{1}</font>",
        "<font><1invalid/></font>",
        "<font><name val=\"&#1;\"/></font>",
        "<font x:unknown=\"1\"/>",
        "<font/>\u{a0}",
    ] {
        let mut wb = Workbook::new();
        wb.get_sheet_mut("Sheet")
            .unwrap()
            .get_or_create_cell(1, 1)
            .unwrap()
            .rich_style = Some(Arc::new(RichStyle {
            font: Some(fragment.into()),
            ..RichStyle::default()
        }));
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("existing.xlsx");
        std::fs::write(&path, b"original contents").unwrap();
        assert!(save_workbook(&path, &wb).is_err(), "accepted {fragment:?}");
        assert_eq!(std::fs::read(&path).unwrap(), b"original contents");
        let mut output = Cursor::new(Vec::new());
        assert!(write_workbook(&mut output, &wb).is_err());
        assert!(output.into_inner().is_empty());
    }
}
