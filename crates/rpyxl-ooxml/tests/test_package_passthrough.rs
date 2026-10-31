use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::{NsReader, Reader};
use rpyxl_core::cell::{CellValue, RichStyle};
use rpyxl_core::package::PackageTemplate;
use rpyxl_core::workbook::Workbook;
use rpyxl_ooxml::writer::{
    save_workbook, save_workbook_generated, write_workbook, write_workbook_generated,
};
use rpyxl_ooxml::{Styles, XlsxPackage};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPC: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const BOOK: &str = "custom/books/book.xml";
const STYLE: &str = "custom/format/native.xml";
const SHEET: &str = "custom/grids/data.xml";

fn source_styles(prefixed: bool) -> String {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{NS}" xmlns:e="urn:private" e:keep="root">
<numFmts count="1" e:keep="formats"><numFmt numFmtId="500" formatCode="0.000 &quot;€&quot;"/></numFmts>
<fonts count="2" e:keep="fonts"><font><name val="Calibri"/></font><font><b/><name val="日本語"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1" e:keep="base"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" e:keep="base-xf"/></cellStyleXfs>
<cellXfs count="3" e:keep="xfs">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="500" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" quotePrefix="1" pivotButton="1" e:keep="first"><alignment horizontal="center"/><protection locked="0"/><extLst><ext uri="private"><e:payload e:value="unchanged"/></ext></extLst></xf>
<xf numFmtId="500" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" quotePrefix="0" pivotButton="1" e:keep="second"><alignment horizontal="center"/><protection locked="0"/><extLst><ext uri="private"><e:payload e:value="unchanged"/></ext></extLst></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="My Normal" xfId="0" builtinId="0" e:keep="named"/></cellStyles>
<dxfs count="1" e:keep="dxfs"><dxf><font><color rgb="FFFF0000"/></font><numFmt numFmtId="501" formatCode="0.0%"/></dxf></dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
<extLst><ext uri="root-extension"><e:payload>keep &amp; preserve</e:payload></ext></extLst>
</styleSheet>"#
    );
    if !prefixed {
        return xml;
    }
    // A stylesheet need not bind SpreadsheetML as its default namespace. New
    // native records must still be in SpreadsheetML, not in the empty namespace.
    let mut reader = Reader::from_str(&xml);
    let mut output = quick_xml::Writer::new(Vec::new());
    loop {
        let event = reader.read_event().unwrap();
        match event {
            Event::Start(mut element) | Event::Empty(mut element) => {
                let empty = xml.as_bytes()[reader.buffer_position() as usize - 2] == b'/';
                if !element.name().as_ref().contains(&b':') {
                    let name = format!("s:{}", String::from_utf8_lossy(element.name().as_ref()));
                    element.set_name(name.as_bytes());
                }
                let mut replacement = element.to_owned();
                replacement.clear_attributes();
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if attribute.key.as_ref() == b"xmlns" {
                        replacement.push_attribute(("xmlns:s", NS));
                    } else {
                        replacement.push_attribute(attribute);
                    }
                }
                output
                    .write_event(if empty {
                        Event::Empty(replacement)
                    } else {
                        Event::Start(replacement)
                    })
                    .unwrap();
            }
            Event::End(element) => {
                let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                let name = if name.contains(':') {
                    name
                } else {
                    format!("s:{name}")
                };
                output
                    .write_event(Event::End(quick_xml::events::BytesEnd::new(name)))
                    .unwrap();
            }
            Event::Eof => break,
            event => output.write_event(event).unwrap(),
        }
    }
    String::from_utf8(output.into_inner()).unwrap()
}

fn fixture(prefixed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{NS}" xmlns:r="{REL}"><workbookPr date1904="0" codeName="OriginalBook"/><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Data" sheetId="19" r:id="grid"/></sheets><definedNames><definedName name="Preserved">Data!$A$1</definedName></definedNames><calcPr calcId="123"/></workbook>"#
    );
    let worksheet = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{NS}" xmlns:r="{REL}" xmlns:e="urn:private" e:keep="worksheet">
<sheetPr codeName="OriginalSheet"><tabColor rgb="FF112233"/></sheetPr><dimension ref="A1:C7"/>
<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="17"/>
<cols><col min="1" max="2" width="23" customWidth="1" style="2"/></cols>
<sheetData e:keep="data"><row r="1" ht="31" customHeight="1" s="2" customFormat="1"><c r="A1" s="2"><v>12.5</v></c><c r="B1" s="1"><v>7</v></c><c r="C1"><v>4</v></c></row><row r="7" s="1" customFormat="1" hidden="1"/></sheetData>
<autoFilter ref="A1:C7"/><mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells>
<conditionalFormatting sqref="A1:A7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>
<hyperlinks><hyperlink ref="B1" r:id="link"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="drawing"/>
<extLst><ext uri="sheet-extension"><e:metadata e:value="日本語"/></ext></extLst></worksheet>"#
    );
    let parts = vec![
        ("_rels/.rels", format!(r#"<Relationships xmlns="{OPC}"><Relationship Id="book" Type="{REL}/officeDocument" Target="/{BOOK}"/></Relationships>"#)),
        ("[Content_Types].xml", format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/{BOOK}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/{STYLE}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/{SHEET}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#)),
        (BOOK, workbook),
        ("custom/books/_rels/book.xml.rels", format!(r#"<Relationships xmlns="{OPC}"><Relationship Id="grid" Type="{REL}/worksheet" Target="../grids/data.xml"/><Relationship Id="styles" Type="{REL}/styles" Target="../format/native.xml"/><Relationship Id="theme" Type="{REL}/theme" Target="../themes/theme.xml"/></Relationships>"#)),
        (STYLE, source_styles(prefixed)),
        (SHEET, worksheet),
        ("custom/grids/_rels/data.xml.rels", format!(r#"<Relationships xmlns="{OPC}"><Relationship Id="link" Type="{REL}/hyperlink" Target="https://example.com/" TargetMode="External"/><Relationship Id="drawing" Type="{REL}/drawing" Target="../drawings/drawing.xml"/></Relationships>"#)),
        ("custom/themes/theme.xml", "<theme xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Keep theme\"/>".into()),
        ("custom/drawings/drawing.xml", "<drawing keep=\"unchanged\"/>".into()),
        ("docProps/custom.xml", "<properties keep=\"日本語 &amp; café\"/>".into()),
        ("opaque/unknown.bin", "opaque payload: do not parse or decompress this member".repeat(100)),
    ];
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    zip.set_comment("original package comment");
    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .unix_permissions(0o640);
    for (name, xml) in parts {
        zip.start_file(name, options).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
    }
    zip.start_file("custom/media/image.png", options).unwrap();
    zip.write_all(b"\x89PNG\r\n\x1a\nopaque image payload")
        .unwrap();
    zip.finish().unwrap().into_inner()
}

fn load(bytes: Vec<u8>) -> Workbook {
    let archive: Arc<[u8]> = bytes.into();
    let mut package =
        XlsxPackage::from_reader(Cursor::new(Arc::clone(&archive)), false, false).unwrap();
    let mut wb = package.to_workbook().unwrap();
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
    wb
}

fn save_bytes(wb: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    write_workbook(&mut output, wb).unwrap();
    output.into_inner()
}

fn part(bytes: &[u8], path: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    archive
        .by_name(path)
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    xml
}

fn attributes(xml: &str, tag: &str) -> Vec<HashMap<String, String>> {
    let mut reader = Reader::from_str(xml);
    let mut result = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == tag.as_bytes() =>
            {
                result.push(
                    element
                        .attributes()
                        .map(|attribute| {
                            let attribute = attribute.unwrap();
                            (
                                String::from_utf8(attribute.key.as_ref().to_vec()).unwrap(),
                                attribute
                                    .decode_and_unescape_value(reader.decoder())
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

fn records(xml: &str, group: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0;
    let mut group_depth = None;
    let mut start = None;
    let mut result = Vec::new();
    loop {
        let offset = reader.buffer_position() as usize;
        match reader.read_event().unwrap() {
            Event::Start(element) => {
                if element.local_name().as_ref() == group.as_bytes() && depth == 1 {
                    group_depth = Some(depth + 1);
                } else if group_depth == Some(depth) {
                    start = Some(offset);
                }
                depth += 1;
            }
            Event::Empty(_) if group_depth == Some(depth) => {
                result.push(xml[offset..reader.buffer_position() as usize].to_owned())
            }
            Event::End(_) => {
                depth -= 1;
                if group_depth == Some(depth) {
                    result.push(
                        xml[start.take().unwrap()..reader.buffer_position() as usize].to_owned(),
                    );
                } else if group_depth == Some(depth + 1) {
                    group_depth = None;
                }
            }
            Event::Eof => return result,
            _ => {}
        }
    }
}

#[test]
fn explicit_xf_zero_survives_for_all_value_kinds_without_changing_inheritance() {
    let original = fixture(false);
    let mut archive = ZipArchive::new(Cursor::new(&original)).unwrap();
    let mut replacement = ZipWriter::new(Cursor::new(Vec::new()));
    let data = r#"<row r="1" s="2" customFormat="1"><c r="A1" s="0"><v>5</v></c><c r="B1" s="0" t="inlineStr"><is><t>text</t></is></c><c r="C1" s="0" t="b"><v>1</v></c><c r="D1" s="0"><f>1+1</f><v>2</v></c><c r="E1" s="0"/><c r="F1"><v>9</v></c><c r="G1"/></row>"#;
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).unwrap();
        replacement
            .start_file(file.name(), SimpleFileOptions::default())
            .unwrap();
        if file.name() == SHEET {
            write!(replacement, r#"<worksheet xmlns="{NS}"><cols><col min="1" max="7" style="2"/></cols><sheetData>{data}</sheetData></worksheet>"#).unwrap();
        } else {
            std::io::copy(&mut file, &mut replacement).unwrap();
        }
    }
    let wb = load(replacement.finish().unwrap().into_inner());
    let sheet = wb.get_sheet("Data").unwrap();
    for column in 1..=5 {
        assert!(sheet.get_cell(1, column).unwrap().source_xf_explicit);
    }
    for column in 6..=7 {
        assert!(!sheet.get_cell(1, column).unwrap().source_xf_explicit);
    }
    for generated in [false, true] {
        let mut saved = Cursor::new(Vec::new());
        let path = if generated {
            write_workbook_generated(&mut saved, &wb).unwrap();
            "xl/worksheets/sheet1.xml"
        } else {
            write_workbook(&mut saved, &wb).unwrap();
            SHEET
        };
        let cells = attributes(&part(&saved.into_inner(), path), "c");
        for column in 0..5 {
            assert_eq!(cells[column].get("s").map(String::as_str), Some("0"));
        }
        let implicit = cells.iter().find(|cell| cell["r"] == "F1").unwrap();
        assert!(!implicit.contains_key("s"));
        assert!(cells
            .iter()
            .filter(|cell| !cell.contains_key("s"))
            .all(|cell| ["F1", "G1"].contains(&cell["r"].as_str())));
    }
}

#[test]
fn untouched_members_are_raw_copied_and_native_metadata_survives() {
    let original = fixture(false);
    let wb = load(original.clone());
    let saved = save_bytes(&wb);
    let mut before = ZipArchive::new(Cursor::new(&original)).unwrap();
    let mut after = ZipArchive::new(Cursor::new(&saved)).unwrap();
    assert_eq!(before.len(), after.len());
    assert_eq!(before.comment(), after.comment());
    for index in 0..before.len() {
        let mut old = before.by_index_raw(index).unwrap();
        if [SHEET, STYLE].contains(&old.name()) {
            continue;
        }
        let mut new = after.by_index_raw(index).unwrap();
        assert_eq!(old.name(), new.name());
        assert_eq!(old.crc32(), new.crc32());
        assert_eq!(old.compression(), new.compression());
        assert_eq!(old.unix_mode(), new.unix_mode());
        let mut old_compressed = Vec::new();
        let mut new_compressed = Vec::new();
        old.read_to_end(&mut old_compressed).unwrap();
        new.read_to_end(&mut new_compressed).unwrap();
        assert_eq!(
            old_compressed,
            new_compressed,
            "raw bytes for {}",
            old.name()
        );
    }
    assert_eq!(part(&saved, STYLE), part(&original, STYLE));
    let sheet = part(&saved, SHEET);
    for tag in [
        "sheetPr",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "autoFilter",
        "mergeCells",
        "conditionalFormatting",
        "hyperlinks",
        "drawing",
        "extLst",
    ] {
        assert_eq!(
            attributes(&sheet, tag),
            attributes(&part(&original, SHEET), tag),
            "{tag}"
        );
    }
    let rows = attributes(&sheet, "row");
    assert_eq!(rows[0]["s"], "2");
    assert_eq!(rows[0]["ht"], "31");
    assert_eq!(rows[1]["r"], "7");
    assert_eq!(rows[1]["s"], "1");
    assert_eq!(rows[1]["hidden"], "1");
    let cells = attributes(&sheet, "c");
    assert_eq!(cells[0]["s"], "2");
    assert_eq!(cells[1]["s"], "1");
    assert!(!cells[2].contains_key("s"));
    assert!(sheet.find("<sheetPr").unwrap() < sheet.find("<dimension").unwrap());
    assert!(sheet.find("<dimension").unwrap() < sheet.find("<sheetViews").unwrap());
    let reread = load(saved);
    assert_eq!(
        reread.get_sheet("Data").unwrap().get_cell_value(1, 1),
        CellValue::Float(12.5)
    );
}

#[test]
fn changed_format_clones_original_xf_and_appends_without_renumbering() {
    let original = fixture(false);
    let mut wb = load(original.clone());
    let sheet = wb.get_sheet_mut("Data").unwrap();
    let cell = sheet.get_cell_mut(1, 1).unwrap();
    cell.value = CellValue::Float(99.125);
    cell.number_format = Some("#,##0.000".into());
    let saved = save_bytes(&wb);
    let xml = part(&saved, STYLE);
    for group in [
        "fonts",
        "fills",
        "borders",
        "numFmts",
        "cellXfs",
        "cellStyleXfs",
        "cellStyles",
        "dxfs",
        "tableStyles",
        "extLst",
    ] {
        let before = records(&part(&original, STYLE), group);
        let after = records(&xml, group);
        assert_eq!(
            &after[..before.len()],
            before.as_slice(),
            "original {group} records"
        );
    }
    let cells = attributes(&part(&saved, SHEET), "c");
    let new_id: usize = cells[0]["s"].parse().unwrap();
    assert_eq!(new_id, 3);
    assert_eq!(cells[1]["s"], "1");
    let new_xf = &records(&xml, "cellXfs")[new_id];
    let xf_attrs = &attributes(new_xf, "xf")[0];
    assert_eq!(xf_attrs["quotePrefix"], "0");
    assert_eq!(xf_attrs["pivotButton"], "1");
    assert_eq!(xf_attrs["e:keep"], "second");
    // ID 501 is already owned by a differential format.
    assert_eq!(xf_attrs["numFmtId"], "502");
    assert!(new_xf.contains("<alignment horizontal=\"center\"/>"));
    assert!(new_xf.contains("<e:payload e:value=\"unchanged\"/>"));
    assert_eq!(attributes(&xml, "cellXfs")[0]["e:keep"], "xfs");
    assert_eq!(attributes(&xml, "numFmts")[0]["e:keep"], "formats");
    assert_eq!(
        Styles::from_reader(xml.as_bytes())
            .unwrap()
            .number_format(new_id),
        Some("#,##0.000")
    );
    // Repeating a save against the original Arc and saving a reloaded result must
    // not append the same new style again.
    assert_eq!(part(&save_bytes(&wb), STYLE), xml);
    let reread = load(saved);
    assert_eq!(part(&save_bytes(&reread), STYLE), xml);
}

#[test]
fn appended_components_deduplicate_and_prefixed_styles_keep_namespaces() {
    let mut wb = load(fixture(true));
    let source = wb.source_styles()[1].clone();
    let style = Arc::new(RichStyle {
        font: Some("<font><b/><name val=\"New font\"/></font>".into()),
        ..(*source).clone()
    });
    for col in [4, 5] {
        let cell = wb
            .get_sheet_mut("Data")
            .unwrap()
            .set_cell_value(2, col, CellValue::Float(0.5))
            .unwrap();
        cell.rich_style = Some(Arc::new((*style).clone()));
        cell.number_format = Some("0.00%".into());
    }
    // A source Arc with a built-in format change must clone its attributes, not
    // generate an unrelated default XF or allocate a custom numFmt for 0.00%.
    wb.get_sheet_mut("Data")
        .unwrap()
        .get_cell_mut(1, 2)
        .unwrap()
        .number_format = Some("0.00%".into());
    let saved = save_bytes(&wb);
    let xml = part(&saved, STYLE);
    assert_eq!(records(&xml, "fonts").len(), 3);
    assert_eq!(records(&xml, "fills").len(), 2);
    assert_eq!(records(&xml, "borders").len(), 1);
    assert_eq!(records(&xml, "numFmts").len(), 1);
    let cells = attributes(&part(&saved, SHEET), "c");
    assert_eq!(cells[3]["s"], cells[4]["s"]);
    let parsed = Styles::from_reader(xml.as_bytes()).unwrap();
    assert_eq!(
        parsed.number_format(cells[3]["s"].parse().unwrap()),
        Some("0.00%")
    );
    let mut reader = NsReader::from_str(&xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Start(element) | Event::Empty(element)
                if !element.name().as_ref().starts_with(b"e:") =>
            {
                assert!(
                    matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == NS.as_bytes())
                );
            }
            Event::Eof => break,
            _ => {}
        }
    }
    assert_eq!(part(&save_bytes(&load(saved)), STYLE), xml);
}

#[test]
fn same_path_save_and_subsequent_modification_are_safe() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("same.xlsx");
    std::fs::write(&path, fixture(false)).unwrap();
    let mut wb = load(std::fs::read(&path).unwrap());
    for value in [100, 200] {
        wb.get_sheet_mut("Data")
            .unwrap()
            .get_cell_mut(1, 1)
            .unwrap()
            .value = CellValue::Int(value);
        save_workbook(&path, &wb).unwrap();
        assert_eq!(
            load(std::fs::read(&path).unwrap())
                .get_sheet("Data")
                .unwrap()
                .get_cell_value(1, 1),
            CellValue::Int(value)
        );
    }
    assert!(part(&std::fs::read(path).unwrap(), BOOK).contains("OriginalBook"));
}

#[test]
fn structural_changes_require_explicit_generated_mode_with_original_style_base() {
    for change in 0..4 {
        let mut wb = load(fixture(false));
        match change {
            0 => wb.rename_sheet("Data", "Renamed").unwrap(),
            1 => {
                wb.create_sheet(Some("Extra"), None).unwrap();
            }
            2 => wb.date_1904 = true,
            _ => wb.get_sheet_mut("Data").unwrap().xml_template = None,
        }
        let mut output = Cursor::new(b"untouched stream".to_vec());
        let error = write_workbook(&mut output, &wb).unwrap_err().to_string();
        assert!(error.contains("generated"), "{error}");
        assert_eq!(output.into_inner(), b"untouched stream");
        let mut generated = Cursor::new(Vec::new());
        write_workbook_generated(&mut generated, &wb).unwrap();
        let bytes = generated.into_inner();
        assert_eq!(part(&bytes, "xl/styles.xml"), part(&fixture(false), STYLE));
        let cells = attributes(&part(&bytes, "xl/worksheets/sheet1.xml"), "c");
        assert_eq!(cells[0]["s"], "2");
        assert_eq!(cells[1]["s"], "1");
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("generated.xlsx");
        save_workbook_generated(&path, &wb).unwrap();
        assert_eq!(
            part(&std::fs::read(path).unwrap(), "xl/styles.xml"),
            part(&bytes, "xl/styles.xml")
        );
    }
}

#[test]
fn malformed_or_mismatched_templates_never_touch_the_destination() {
    for change in 0..10 {
        let mut wb = load(fixture(false));
        match change {
            0 => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).styles_xml =
                    "<styleSheet>".into()
            }
            1 => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).styles_path =
                    "not/present.xml".into()
            }
            2 => Arc::make_mut(wb.package_template.as_mut().unwrap()).sheet_parts[0] = STYLE.into(),
            3 => Arc::make_mut(wb.package_template.as_mut().unwrap()).workbook_path = SHEET.into(),
            4 => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).styles_xml =
                    source_styles(false).replace("quotePrefix=\"0\"", "quotePrefix=\"1\"")
            }
            5 => {
                wb.get_sheet_mut("Data")
                    .unwrap()
                    .xml_template
                    .as_mut()
                    .unwrap()
                    .after_data = "</wrong>".into()
            }
            6 => {
                wb.get_sheet_mut("Data")
                    .unwrap()
                    .xml_template
                    .as_mut()
                    .unwrap()
                    .before_data = "<autoFilter ref=\"A1\"/>".into()
            }
            7 => wb.set_source_styles(Vec::new()),
            8 => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).archive =
                    Arc::from(b"invalid zip".as_slice())
            }
            _ => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).styles_xml =
                    source_styles(false).replace("fontId=\"1\"", "fontId=\"99\"")
            }
        }
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("do-not-truncate.xlsx");
        std::fs::write(&path, b"existing destination").unwrap();
        assert!(save_workbook(&path, &wb).is_err(), "invalid case {change}");
        assert_eq!(std::fs::read(&path).unwrap(), b"existing destination");
        let missing = directory.path().join("must-not-create.xlsx");
        assert!(save_workbook(&missing, &wb).is_err());
        assert!(!missing.exists());
        let mut output = Cursor::new(b"existing stream".to_vec());
        assert!(write_workbook(&mut output, &wb).is_err());
        assert_eq!(output.into_inner(), b"existing stream");
    }
}

fn replace_part(bytes: Vec<u8>, path: &str, replacement: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut output = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let file = archive.by_index_raw(index).unwrap();
        if file.name() == path {
            output
                .start_file(path, SimpleFileOptions::default())
                .unwrap();
            output.write_all(replacement.as_bytes()).unwrap();
        } else {
            output.raw_copy_file(file).unwrap();
        }
    }
    output.finish().unwrap().into_inner()
}

#[test]
fn blank_source_style_zero_survives_and_temporal_formats_keep_original_xf_properties() {
    let mut wb = load(fixture(false));
    let zero = Arc::clone(&wb.source_styles()[0]);
    let sheet = wb.get_sheet_mut("Data").unwrap();
    sheet
        .set_cell_value(2, 1, CellValue::Empty)
        .unwrap()
        .rich_style = Some(zero);
    sheet.get_cell_mut(1, 1).unwrap().value = CellValue::Date("2026-10-31".into());
    let saved = save_bytes(&wb);
    let sheet = part(&saved, SHEET);
    let cells = attributes(&sheet, "c");
    let blank = cells.iter().find(|cell| cell["r"] == "A2").unwrap();
    assert_eq!(blank["s"], "0");
    let styles = part(&saved, STYLE);
    let id: usize = cells[0]["s"].parse().unwrap();
    assert_eq!(
        Styles::from_reader(styles.as_bytes())
            .unwrap()
            .number_format(id),
        Some("yyyy-mm-dd")
    );
    assert_eq!(
        attributes(&records(&styles, "cellXfs")[id], "xf")[0]["e:keep"],
        "second"
    );
    let loaded = load(saved);
    assert!(loaded.get_sheet("Data").unwrap().get_cell(2, 1).is_some());
    assert_eq!(part(&save_bytes(&loaded), STYLE), styles);
}

#[test]
fn absent_number_format_group_is_inserted_and_attribute_whitespace_roundtrips() {
    let original = fixture(true);
    let styles = part(&original, STYLE);
    let start = styles.find("<s:numFmts").unwrap();
    let end = styles.find("</s:numFmts>").unwrap() + "</s:numFmts>".len();
    let styles = format!("{}{}", &styles[..start], &styles[end..])
        .replace("numFmtId=\"500\"", "numFmtId=\"0\"");
    let mut wb = load(replace_part(original, STYLE, &styles));
    let format = "0.000 \"日本語 €\"\t\n\r";
    wb.get_sheet_mut("Data")
        .unwrap()
        .get_cell_mut(1, 1)
        .unwrap()
        .number_format = Some(format.into());
    let saved = save_bytes(&wb);
    let xml = part(&saved, STYLE);
    assert!(xml.find("<numFmts").unwrap() < xml.find("<s:fonts").unwrap());
    assert_eq!(attributes(&xml, "numFmt")[0]["formatCode"], format);
    let cells = attributes(&part(&saved, SHEET), "c");
    assert_eq!(
        Styles::from_reader(xml.as_bytes())
            .unwrap()
            .number_format(cells[0]["s"].parse().unwrap()),
        Some(format)
    );
}

#[test]
fn unavailable_original_cell_row_column_and_dxf_references_are_rejected() {
    for (from, to) in [
        ("s=\"2\"", "s=\"99\""),
        ("style=\"2\"", "style=\"99\""),
        ("dxfId=\"0\"", "dxfId=\"99\""),
    ] {
        let bytes = fixture(false);
        let sheet = part(&bytes, SHEET).replace(from, to);
        let wb = load(replace_part(bytes, SHEET, &sheet));
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("existing.xlsx");
        std::fs::write(&path, b"unchanged").unwrap();
        assert!(save_workbook(&path, &wb).is_err());
        assert_eq!(std::fs::read(path).unwrap(), b"unchanged");
    }
}

#[test]
fn ordinary_rust_reader_saves_use_the_independent_style_base() {
    for prefixed in [false, true] {
        let original = fixture(prefixed);
        let expected_styles = part(&original, STYLE);
        let mut package = XlsxPackage::from_reader(Cursor::new(original), false, false).unwrap();
        let wb = package.to_workbook().unwrap();
        assert!(wb.package_template.is_none());
        assert_eq!(
            wb.source_styles_xml.as_deref(),
            Some(expected_styles.as_str())
        );
        assert_eq!(
            wb.get_sheet("Data")
                .unwrap()
                .get_cell(1, 1)
                .unwrap()
                .source_xf_index(),
            Some(2)
        );
        let saved = save_bytes(&wb);
        assert_eq!(part(&saved, "xl/styles.xml"), expected_styles);
        let cells = attributes(&part(&saved, "xl/worksheets/sheet1.xml"), "c");
        assert_eq!(cells[0]["s"], "2");
        assert_eq!(cells[1]["s"], "1");
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("ordinary-reader.xlsx");
        save_workbook(&path, &wb).unwrap();
        assert_eq!(
            part(&std::fs::read(path).unwrap(), "xl/styles.xml"),
            expected_styles
        );
    }
}

#[test]
fn unsupported_worksheet_framing_retains_styles_in_generated_mode() {
    let original = fixture(false);
    let sheet = part(&original, SHEET)
        .replace("<worksheet ", &format!("<s:worksheet xmlns:s=\"{NS}\" "))
        .replace("</worksheet>", "</s:worksheet>");
    let mut wb = load(replace_part(original, SHEET, &sheet));
    assert!(wb.get_sheet("Data").unwrap().xml_template.is_none());
    assert!(!wb.can_preserve_package());
    let expected_styles = wb.source_styles_xml.clone().unwrap();
    // Generated mode must use the independent base even when no usable package
    // template remains, rather than taking its style base from package eligibility.
    Arc::make_mut(wb.package_template.as_mut().unwrap())
        .styles_xml
        .clear();
    let mut output = Cursor::new(Vec::new());
    write_workbook_generated(&mut output, &wb).unwrap();
    let generated = output.into_inner();
    assert_eq!(part(&generated, "xl/styles.xml"), expected_styles);
    assert_eq!(
        attributes(&part(&generated, "xl/worksheets/sheet1.xml"), "c")[0]["s"],
        "2"
    );
    wb.package_template = None;
    assert_eq!(part(&save_bytes(&wb), "xl/styles.xml"), expected_styles);
}

#[test]
fn font_edits_clone_provenance_xfs_even_when_replacement_arcs_are_shared() {
    for prefixed in [false, true] {
        for generated in [false, true] {
            let original = fixture(prefixed);
            let mut wb = load(original.clone());
            let replacement = Arc::new(RichStyle {
                font: Some("<font><b/><name val=\"Edited font\"/></font>".into()),
                ..(*wb.source_styles()[1]).clone()
            });
            for (column, source_xf) in [(1, 2), (2, 1)] {
                let cell = wb
                    .get_sheet_mut("Data")
                    .unwrap()
                    .get_cell_mut(1, column)
                    .unwrap();
                assert_eq!(cell.source_xf_index(), Some(source_xf));
                cell.rich_style = Some(Arc::clone(&replacement));
            }
            let mut output = Cursor::new(Vec::new());
            if generated {
                write_workbook_generated(&mut output, &wb).unwrap();
            } else {
                write_workbook(&mut output, &wb).unwrap();
            }
            let bytes = output.into_inner();
            let (style_path, sheet_path) = if generated {
                ("xl/styles.xml", "xl/worksheets/sheet1.xml")
            } else {
                (STYLE, SHEET)
            };
            let styles = part(&bytes, style_path);
            let original_xfs = records(&part(&original, STYLE), "cellXfs");
            let xfs = records(&styles, "cellXfs");
            assert_eq!(&xfs[..3], original_xfs.as_slice());
            assert_eq!(xfs.len(), 5);
            assert_eq!(records(&styles, "fonts").len(), 3);
            let cells = attributes(&part(&bytes, sheet_path), "c");
            assert_ne!(
                cells[0]["s"], cells[1]["s"],
                "different provenance must not share an XF cache entry"
            );
            for (column, source_xf, keep, quote) in [(0, 2, "second", "0"), (1, 1, "first", "1")] {
                let id: usize = cells[column]["s"].parse().unwrap();
                let attributes = attributes(&xfs[id], "xf");
                let attributes = &attributes[0];
                assert_eq!(attributes["fontId"], "2");
                assert_eq!(attributes["numFmtId"], "500");
                assert_eq!(attributes["quotePrefix"], quote);
                assert_eq!(attributes["pivotButton"], "1");
                assert_eq!(attributes["xfId"], "0");
                assert_eq!(attributes["e:keep"], keep);
                assert_eq!(
                    records(&format!("<root>{}</root>", xfs[id]), "xf"),
                    records(&format!("<root>{}</root>", original_xfs[source_xf]), "xf"),
                    "alignment, protection, and XF extensions must remain verbatim",
                );
                assert!(xfs[id].contains("e:value=\"unchanged\""));
                let parsed = Styles::from_reader(styles.as_bytes()).unwrap();
                assert!(parsed
                    .rich_style(id)
                    .unwrap()
                    .font
                    .as_deref()
                    .unwrap()
                    .contains("Edited font"));
            }
            if !generated {
                assert_eq!(part(&save_bytes(&load(bytes)), STYLE), styles);
            }
        }
    }
}

#[test]
fn native_core_axis_edits_refuse_default_passthrough_without_touching_files() {
    for operation in 0..4 {
        let original = fixture(false);
        let directory = tempfile::tempdir().unwrap();
        let source = directory.path().join("source.xlsx");
        let destination = directory.path().join("destination.xlsx");
        std::fs::write(&source, &original).unwrap();
        std::fs::write(&destination, b"existing destination").unwrap();
        let mut wb = load(original.clone());
        let sheet = wb.get_sheet_mut("Data").unwrap();
        match operation {
            0 => sheet.insert_rows(1, 2).unwrap(),
            1 => sheet.delete_rows(1, 1).unwrap(),
            2 => sheet.insert_cols(1, 2).unwrap(),
            _ => sheet.delete_cols(1, 1).unwrap(),
        }
        assert!(sheet.xml_template.is_none());
        assert!(!wb.can_preserve_package());
        for path in [&source, &destination] {
            let error = save_workbook(path, &wb).unwrap_err().to_string();
            assert!(error.contains("generated"), "{error}");
        }
        assert_eq!(std::fs::read(&source).unwrap(), original);
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            b"existing destination"
        );
        let mut stream = Cursor::new(b"existing stream".to_vec());
        assert!(write_workbook(&mut stream, &wb).is_err());
        assert_eq!(stream.into_inner(), b"existing stream");
        assert_eq!(
            wb.package_template.as_ref().unwrap().archive.as_ref(),
            original.as_slice()
        );
        let mut generated = Cursor::new(Vec::new());
        write_workbook_generated(&mut generated, &wb).unwrap();
        assert_eq!(
            part(&generated.into_inner(), "xl/styles.xml"),
            wb.source_styles_xml.as_deref().unwrap()
        );
    }
}

#[test]
fn legacy_namespace_less_style_fixtures_normalize_without_renumbering_xfs() {
    let styles = r#"<styleSheet><numFmts><numFmt numFmtId="500" formatCode="0.000"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="500" xfId="0" quotePrefix="1"/><xf numFmtId="500" xfId="0" quotePrefix="0" pivotButton="1"/></cellXfs></styleSheet>"#;
    let bytes = replace_part(fixture(false), STYLE, styles);
    let mut package = XlsxPackage::from_reader(Cursor::new(bytes), false, false).unwrap();
    let wb = package.to_workbook().unwrap();
    assert!(wb.package_template.is_none());
    let saved = save_bytes(&wb);
    let xml = part(&saved, "xl/styles.xml");
    let mut reader = NsReader::from_str(&xml);
    let (namespace, event) = reader.read_resolved_event().unwrap();
    assert!(matches!(event, Event::Start(_)));
    assert!(
        matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == NS.as_bytes())
    );
    assert_eq!(records(&xml, "cellXfs"), records(styles, "cellXfs"));
    let cells = attributes(&part(&saved, "xl/worksheets/sheet1.xml"), "c");
    assert_eq!(cells[0]["s"], "2");
    assert_eq!(cells[1]["s"], "1");
    let parsed = Styles::from_reader(xml.as_bytes()).unwrap();
    assert_eq!(parsed.number_format(2), Some("0.000"));
    assert_eq!(parsed.rich_styles().len(), 3);
}

#[test]
fn invalid_original_style_base_is_not_silently_rebuilt_in_generated_mode() {
    let mut wb = load(fixture(false));
    wb.source_styles_xml = Some(String::new());
    let mut output = Cursor::new(b"untouched".to_vec());
    assert!(write_workbook_generated(&mut output, &wb).is_err());
    assert_eq!(output.into_inner(), b"untouched");
}
