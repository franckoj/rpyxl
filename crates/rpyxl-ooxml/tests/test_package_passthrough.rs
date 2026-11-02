use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::{NsReader, Reader};
use rpyxl_core::cell::{CellValue, RichStyle};
use rpyxl_core::package::PackageTemplate;
use rpyxl_core::workbook::Workbook;
use rpyxl_ooxml::writer::{
    enable_plain_structural_edits, save_workbook, save_workbook_generated, write_workbook,
    write_workbook_generated,
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
    load_with_data_only(bytes, false)
}

fn load_with_data_only(bytes: Vec<u8>, data_only: bool) -> Workbook {
    let archive: Arc<[u8]> = bytes.into();
    let mut package =
        XlsxPackage::from_reader(Cursor::new(Arc::clone(&archive)), false, data_only).unwrap();
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

fn plain_fixture() -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{NS}"><workbookPr/><workbookProtection/><bookViews><workbookView visibility="visible" minimized="0" showHorizontalScroll="1" showVerticalScroll="1" showSheetTabs="1" tabRatio="600" firstSheet="0" activeTab="0" autoFilterDateGrouping="1"/></bookViews><sheets><sheet xmlns:r="{REL}" name="Data" sheetId="1" state="visible" r:id="grid"/></sheets><definedNames/><calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>"#
    );
    let worksheet = format!(
        r#"<worksheet xmlns="{NS}"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr/></sheetPr><dimension ref="A1:D8"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr baseColWidth="8" defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" s="2" t="n"><v>12.5</v></c><c r="B1" s="1" t="n"><v>7</v></c><c r="C1" s="0" t="n"></c><c r="D1" t="n"><v>4</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t xml:space="preserve"> =日本語 &amp; café </t></is></c><c r="B4" t="b"><v>1</v></c></row><row r="8"/></sheetData><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/></worksheet>"#
    );
    let bytes = replace_part(fixture(false), BOOK, &workbook);
    let bytes = replace_part(bytes, SHEET, &worksheet);
    replace_part(
        bytes,
        "custom/grids/_rels/data.xml.rels",
        &format!(r#"<Relationships xmlns="{OPC}"/>"#),
    )
}

fn apply_plain_edit(wb: &mut Workbook, operation: usize) {
    let sheet = wb.get_sheet_mut("Data").unwrap();
    match operation {
        0 => sheet.insert_rows(1, 2).unwrap(),
        1 => sheet.delete_rows(1, 1).unwrap(),
        2 => sheet.insert_cols(1, 2).unwrap(),
        3 => sheet.delete_cols(1, 1).unwrap(),
        _ => sheet.move_range("A1:H8", 1, 1).unwrap(),
    }
}

fn assert_native_cells_roundtrip(wb: &Workbook, saved: &[u8]) {
    let loaded = load(saved.to_vec());
    for (expected, actual) in wb.sheets().iter().zip(loaded.sheets()) {
        assert_eq!(expected.cells().len(), actual.cells().len());
        for (key, cell) in expected.cells() {
            let actual = actual.cells().get(key).unwrap();
            assert_eq!(cell.value, actual.value, "{key:?}");
            assert_eq!(cell.source_xf_index(), actual.source_xf_index(), "{key:?}");
            assert_eq!(
                cell.source_xf_explicit, actual.source_xf_explicit,
                "{key:?}"
            );
        }
    }
}

#[test]
fn plain_optin_accepts_only_empty_default_workbook_protection() {
    let original = plain_fixture();
    let book = part(&original, BOOK);
    assert!(book.contains("<workbookPr/><workbookProtection/><bookViews>"));
    for protection in [
        "<workbookProtection/>",
        "<workbookProtection></workbookProtection>",
        "",
    ] {
        let book = book.replace("<workbookProtection/>", protection);
        let mut wb = load(replace_part(original.clone(), BOOK, &book));
        assert!(enable_plain_structural_edits(&mut wb).unwrap());
        apply_plain_edit(&mut wb, 4);
        let saved = save_bytes(&wb);
        assert_eq!(part(&saved, BOOK), book);
        assert_native_cells_roundtrip(&wb, &saved);
        assert!(enable_plain_structural_edits(&mut load(saved)).unwrap());
    }
    for protection in [
        "<workbookProtection lockStructure=\"1\"/>",
        "<workbookProtection lockStructure=\"0\"/>",
        "<workbookProtection lockWindows=\"1\"/>",
        "<workbookProtection workbookPassword=\"ABCD\"/>",
        "<workbookProtection unknown=\"1\"/>",
        "<workbookProtection><extLst/></workbookProtection>",
        "<workbookProtection><unknown xmlns=\"urn:foreign\"/></workbookProtection>",
        "<workbookProtection>protected</workbookProtection>",
        "<workbookProtection> </workbookProtection>",
        "<workbookProtection xmlns=\"urn:foreign\"/>",
        "<workbookProtection xmlns:e=\"urn:foreign\" e:lock=\"1\"/>",
        "<workbookProtection xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>",
        "<workbookProtection/><workbookProtection/>",
    ] {
        let mut wb = load(replace_part(
            original.clone(),
            BOOK,
            &book.replace("<workbookProtection/>", protection),
        ));
        let snapshot = wb.clone();
        assert!(
            !enable_plain_structural_edits(&mut wb).unwrap(),
            "accepted {protection}"
        );
        assert_eq!(wb, snapshot);
        assert!(wb
            .sheets()
            .iter()
            .all(|sheet| !sheet.structural_passthrough_enabled()));
    }
    let misplaced = book
        .replace("<workbookProtection/>", "")
        .replace("<sheets>", "<workbookProtection/><sheets>");
    let mut wb = load(replace_part(original, BOOK, &misplaced));
    let snapshot = wb.clone();
    assert!(!enable_plain_structural_edits(&mut wb).unwrap());
    assert_eq!(wb, snapshot);
}

#[test]
fn entity_encoded_cell_attributes_match_source_values_styles_and_plain_proof() {
    let original = plain_fixture();
    let sheet = part(&original, SHEET);
    let cases = [
        ("coordinate", sheet.replace("r=\"A1\"", "r=\"A&#49;\"")),
        ("style", sheet.replace("s=\"2\"", "s=\"&#50;\"")),
        ("boolean_type", sheet.replace("t=\"b\"", "t=\"&#98;\"")),
        (
            "inline_type",
            sheet.replace("t=\"inlineStr\"", "t=\"inline&#x53;tr\""),
        ),
        (
            "string_type",
            sheet.replace("<c r=\"D1\" t=\"n\">", "<c r=\"D1\" t=\"&#115;tr\">"),
        ),
        (
            "empty_cell",
            sheet.replace(
                "<c r=\"C1\" s=\"0\" t=\"n\"></c>",
                "<c r=\"C&#49;\" s=\"&#x32;\" t=\"&#110;\"/>",
            ),
        ),
        (
            "combined",
            sheet
                .replace("r=\"A1\"", "r=\"&#x41;&#49;\"")
                .replace("s=\"2\"", "s=\"&#x32;\"")
                .replace("s=\"1\"", "s=\"&#49;\"")
                .replace("s=\"0\"", "s=\"&#48;\"")
                .replace("t=\"n\"", "t=\"&#x6e;\"")
                .replace("t=\"b\"", "t=\"&#x62;\"")
                .replace("t=\"inlineStr\"", "t=\"&#105;nlineStr\"")
                .replace("<row r=\"1\">", "<row r=\"&#49;\">")
                .replace("<row r=\"4\">", "<row r=\"&#x34;\">")
                .replace("<row r=\"8\"/>", "<row r=\"&#56;\"/>"),
        ),
    ];
    let style_xml = part(&original, STYLE);
    let source_styles = Styles::from_reader(style_xml.as_bytes()).unwrap();
    for (case, xml) in cases {
        // These expectations come from decoded source XML, not from a native
        // workbook that could have already dropped a cell or misread its XF.
        let source_cells = attributes(&xml, "c");
        assert_eq!(source_cells.len(), 6);
        let check_source = |wb: &Workbook, offset: u32| {
            let sheet = wb.get_sheet("Data").unwrap();
            assert_eq!(sheet.cell_count(), source_cells.len(), "{case}");
            for source in &source_cells {
                let reference = source["r"].as_str();
                let (row, col) = rpyxl_core::coordinate::coordinate_to_tuple(reference).unwrap();
                let expected = match reference {
                    "A1" => CellValue::Float(12.5),
                    "B1" => CellValue::Int(7),
                    "C1" => CellValue::Empty,
                    "D1" if source.get("t").map(String::as_str) == Some("str") => {
                        CellValue::String("4".into())
                    }
                    "D1" => CellValue::Int(4),
                    "A4" => CellValue::String(" =日本語 & café ".into()),
                    "B4" => CellValue::Bool(true),
                    _ => panic!("unexpected source reference {reference}"),
                };
                let cell = sheet
                    .get_cell(row + offset, col + offset)
                    .unwrap_or_else(|| panic!("missing source cell {reference} in {case}"));
                let style = source.get("s").map_or(0, |id| id.parse::<usize>().unwrap());
                assert_eq!(cell.value, expected, "{case}: {reference}");
                assert_eq!(cell.source_xf_index(), Some(style), "{case}: {reference}");
                assert_eq!(
                    cell.source_xf_explicit,
                    source.contains_key("s"),
                    "{case}: {reference}"
                );
                assert_eq!(
                    cell.rich_style,
                    source_styles.rich_style(style),
                    "{case}: {reference}"
                );
                assert_eq!(
                    cell.number_format.as_deref(),
                    source_styles
                        .number_format(style)
                        .filter(|code| !code.eq_ignore_ascii_case("General")),
                    "{case}: {reference}"
                );
            }
        };
        let bytes = replace_part(original.clone(), SHEET, &xml);
        for data_only in [false, true] {
            let mut wb = load_with_data_only(bytes.clone(), data_only);
            check_source(&wb, 0);
            assert!(enable_plain_structural_edits(&mut wb).unwrap(), "{case}");
            let unedited = save_bytes(&wb);
            check_source(&load(unedited.clone()), 0);
            assert_eq!(
                attributes(&part(&unedited, SHEET), "row"),
                attributes(&xml, "row")
            );
            wb.active_mut().move_range("A1:D4", 1, 1).unwrap();
            check_source(&wb, 1);
            assert!(wb.can_preserve_package());
            assert!(enable_plain_structural_edits(&mut wb).unwrap());
            let saved = save_bytes(&wb);
            check_source(&load(saved.clone()), 1);
            assert_eq!(part(&saved, STYLE), style_xml);
            let mut reloaded = load(saved);
            assert!(enable_plain_structural_edits(&mut reloaded).unwrap());
            reloaded.active_mut().move_range("B2:E5", 1, 1).unwrap();
            check_source(&load(save_bytes(&reloaded)), 2);
        }
    }
}

#[test]
fn normal_parser_keeps_invalid_decoded_coordinate_and_style_fallbacks() {
    let original = plain_fixture();
    let sheet = part(&original, SHEET);
    let mut bad_coordinate = load(replace_part(
        original.clone(),
        SHEET,
        &sheet.replace("r=\"A1\"", "r=\"A&#48;\""),
    ));
    assert!(bad_coordinate.active().get_cell(1, 1).is_none());
    assert_eq!(bad_coordinate.active().cell_count(), 5);
    assert!(!enable_plain_structural_edits(&mut bad_coordinate).unwrap());

    let mut bad_style = load(replace_part(
        original,
        SHEET,
        &sheet.replace("s=\"2\"", "s=\"&#120;\""),
    ));
    let cell = bad_style.active().get_cell(1, 1).unwrap();
    assert_eq!(cell.value, CellValue::Float(12.5));
    assert_eq!(cell.source_xf_index(), Some(0));
    assert!(cell.source_xf_explicit);
    assert!(!enable_plain_structural_edits(&mut bad_style).unwrap());
}

#[test]
fn plain_structural_optin_all_five_ops_repeat_save_without_ghosts_or_style_renumbering() {
    for operation in 0..5 {
        let original = plain_fixture();
        let mut wb = load(original.clone());
        assert!(wb
            .sheets()
            .iter()
            .all(|sheet| !sheet.structural_passthrough_enabled() && !sheet.source_rows_changed()));
        assert!(enable_plain_structural_edits(&mut wb).unwrap());
        assert!(enable_plain_structural_edits(&mut wb).unwrap());
        // Before a structural edit the original empty row still survives.
        assert!(attributes(&part(&save_bytes(&wb), SHEET), "row")
            .iter()
            .any(|row| row["r"] == "8"));
        for _ in 0..2 {
            apply_plain_edit(&mut wb, operation);
            assert!(wb.can_preserve_package());
            assert!(enable_plain_structural_edits(&mut wb).unwrap());
            let sheet = wb.get_sheet("Data").unwrap();
            assert!(sheet.xml_template.is_some());
            assert!(sheet.source_rows_changed());
            let saved = save_bytes(&wb);
            assert_eq!(saved, save_bytes(&wb));
            assert_native_cells_roundtrip(&wb, &saved);
            assert_eq!(part(&saved, STYLE), part(&original, STYLE));
            let xml = part(&saved, SHEET);
            let expected_rows = sheet
                .cells()
                .keys()
                .map(|&(row, _)| row.to_string())
                .collect::<std::collections::BTreeSet<_>>();
            let actual_rows = attributes(&xml, "row")
                .iter()
                .map(|row| row["r"].clone())
                .collect::<std::collections::BTreeSet<_>>();
            assert_eq!(expected_rows, actual_rows, "operation {operation}");
            assert_eq!(attributes(&xml, "row").len(), expected_rows.len());
            let mut reloaded = load(saved);
            assert!(enable_plain_structural_edits(&mut reloaded).unwrap());
            assert_native_cells_roundtrip(&reloaded, &save_bytes(&reloaded));
        }
        let mut before = ZipArchive::new(Cursor::new(&original)).unwrap();
        let saved = save_bytes(&wb);
        let mut after = ZipArchive::new(Cursor::new(&saved)).unwrap();
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
            let mut a = Vec::new();
            let mut b = Vec::new();
            old.read_to_end(&mut a).unwrap();
            new.read_to_end(&mut b).unwrap();
            assert_eq!(a, b, "raw-copy {}", old.name());
        }
    }
}

#[test]
fn plain_optin_noops_bounds_failures_and_same_path_saves_are_atomic() {
    let original = plain_fixture();
    let mut wb = load(original.clone());
    assert!(enable_plain_structural_edits(&mut wb).unwrap());
    let snapshot = wb.clone();
    let sheet = wb.get_sheet_mut("Data").unwrap();
    sheet.insert_rows(1, 0).unwrap();
    sheet.delete_rows(1, 0).unwrap();
    sheet.insert_cols(1, 0).unwrap();
    sheet.delete_cols(1, 0).unwrap();
    sheet.move_range("A1:D4", 0, 0).unwrap();
    assert_eq!(wb, snapshot);
    for operation in 0..5 {
        let sheet = wb.get_sheet_mut("Data").unwrap();
        let result = match operation {
            0 => sheet.insert_rows(1, u32::MAX),
            1 => sheet.delete_rows(0, 1),
            2 => sheet.insert_cols(1, u32::MAX),
            3 => sheet.delete_cols(0, 1),
            _ => sheet.move_range("A1:D4", -1, 0),
        };
        assert!(result.is_err());
        assert_eq!(wb, snapshot);
    }
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("plain.xlsx");
    for operation in 0..5 {
        let mut wb = load(original.clone());
        assert!(enable_plain_structural_edits(&mut wb).unwrap());
        std::fs::write(&path, &original).unwrap();
        for _ in 0..2 {
            apply_plain_edit(&mut wb, operation);
            save_workbook(&path, &wb).unwrap();
            assert_native_cells_roundtrip(&wb, &std::fs::read(&path).unwrap());
        }
        assert_eq!(
            wb.package_template.as_ref().unwrap().archive.as_ref(),
            original.as_slice()
        );
    }
}

#[test]
fn plain_optin_is_not_implicit_even_for_plain_sources() {
    for operation in 0..5 {
        let mut wb = load(plain_fixture());
        apply_plain_edit(&mut wb, operation);
        assert!(!wb.can_preserve_package());
        assert!(!enable_plain_structural_edits(&mut wb).unwrap());
        assert!(wb.active().xml_template.is_none());
    }
    assert!(!enable_plain_structural_edits(&mut Workbook::new()).unwrap());
    let mut unsafe_book = load(fixture(false));
    let snapshot = unsafe_book.clone();
    assert!(!enable_plain_structural_edits(&mut unsafe_book).unwrap());
    assert_eq!(snapshot, unsafe_book);
}

#[test]
fn plain_optin_rejects_unknown_framing_reference_features_and_cell_payloads_atomically() {
    let original = plain_fixture();
    let sheet = part(&original, SHEET);
    let changes = [
        ("<sheetPr>", "<sheetPr codeName=\"Unknown\">"),
        ("<worksheet ", "<worksheet xmlns:e=\"urn:foreign\" e:keep=\"yes\" "),
        ("<sheetData>", "<sheetData xmlns=\"urn:foreign\">"),
        ("<sheetData>", "<sheetData future=\"1\">"),
        ("<row r=\"1\">", "<row r=\"1\" hidden=\"1\">"),
        ("<row r=\"8\"/>", "<row r=\"8\" ht=\"30\" customHeight=\"1\"/>"),
        ("<row r=\"8\"/>", "<row r=\"0\"/>"),
        ("<row r=\"8\"/>", "<row r=\"4\"/>"),
        ("<row r=\"8\"/>", "<row r=\"1048577\"/>"),
        ("<v>12.5</v>", "<f>1+1</f><v>2</v>"),
        ("<v>12.5</v>", "<f t=\"shared\" si=\"0\" ref=\"A1:A2\">1</f><v>1</v>"),
        ("<v>12.5</v>", "<v>not-a-number</v>"),
        ("<v>12.5</v>", "<v>12.5</v><!--lost-->"),
        ("<v>12.5</v>", "<v>12.5</v><?unknown lost?>"),
        ("<v>12.5</v>", "<v>12.5</v><extLst/>"),
        ("<v>12.5</v>", "<v xmlns=\"urn:foreign\">12.5</v>"),
        ("r=\"A1\"", "r=\"A2\""),
        ("r=\"A1\"", "r=\"XFE1\""),
        ("s=\"2\"", "s=\"99\""),
        ("r=\"A1\"", "r=\"A1\" cm=\"1\""),
        ("<is>", "<is><r><t>Rich</t></r>"),
        ("<is>", "<is><rPh sb=\"0\" eb=\"1\"><t>Phonetic</t></rPh>"),
        ("sqref=\"A1\"", "sqref=\"B2\""),
        ("activeCell=\"A1\"", "activeCell=\"B2\""),
        ("<selection ", "<selection pane=\"bottomRight\" "),
        ("</sheetView>", "<pane state=\"frozen\" ySplit=\"1\"/></sheetView>"),
        ("<sheetData>", "<cols><col min=\"1\" max=\"1\" width=\"20\"/></cols><sheetData>"),
        ("</worksheet>", "<autoFilter ref=\"A1:D4\"/></worksheet>"),
        ("</worksheet>", "<mergeCells><mergeCell ref=\"A1:B1\"/></mergeCells></worksheet>"),
        ("</worksheet>", "<conditionalFormatting sqref=\"A1\"/></worksheet>"),
        ("</worksheet>", "<dataValidations count=\"0\"/></worksheet>"),
        ("</worksheet>", "<extLst/></worksheet>"),
        ("</worksheet>", "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"x\"/></worksheet>"),
        ("ref=\"A1:D8\"", "ref=\"A1:D8\" unknown=\"1\""),
        ("ref=\"A1:D8\"", "ref=\"A1:XFE8\""),
    ];
    for (old, new) in changes {
        let bytes = replace_part(original.clone(), SHEET, &sheet.replace(old, new));
        let mut wb = load(bytes);
        let snapshot = wb.clone();
        assert!(
            !enable_plain_structural_edits(&mut wb).unwrap(),
            "accepted {new}"
        );
        assert_eq!(wb, snapshot, "mutated on rejection {new}");
        assert!(wb
            .sheets()
            .iter()
            .all(|sheet| !sheet.structural_passthrough_enabled()));
    }
    let book = part(&original, BOOK);
    for replacement in [
        "<definedNames><definedName name=\"Ref\">Data!$A$1</definedName></definedNames>",
        "<externalReferences/>",
        "<pivotCaches/>",
        "<extLst/>",
        "<unknown xmlns=\"urn:foreign\"/>",
    ] {
        let mut wb = load(replace_part(
            original.clone(),
            BOOK,
            &book.replace("<definedNames/>", replacement),
        ));
        let snapshot = wb.clone();
        assert!(
            !enable_plain_structural_edits(&mut wb).unwrap(),
            "accepted {replacement}"
        );
        assert_eq!(wb, snapshot);
    }
    for kind in [
        "chart",
        "pivotTable",
        "externalLink",
        "calcChain",
        "drawing",
        "table",
        "comments",
        "sharedStrings",
        "unknown",
    ] {
        let rels = part(&original, "custom/books/_rels/book.xml.rels");
        let rels = rels.replace("</Relationships>", &format!(r#"<Relationship Id="unsafe" Type="{REL}/{kind}" Target="../opaque.xml"/></Relationships>"#));
        let mut wb = load(replace_part(
            original.clone(),
            "custom/books/_rels/book.xml.rels",
            &rels,
        ));
        assert!(
            !enable_plain_structural_edits(&mut wb).unwrap(),
            "accepted {kind}"
        );
    }
    // A source formula remains disqualifying even if the live value was replaced
    // (including a data-only load that hid its formula from the caller).
    let formula = sheet.replace("<v>12.5</v>", "<f>1+1</f><v>2</v>");
    for data_only in [false, true] {
        let mut wb =
            load_with_data_only(replace_part(original.clone(), SHEET, &formula), data_only);
        wb.active_mut()
            .set_cell_value(1, 1, CellValue::Int(2))
            .unwrap();
        assert!(!enable_plain_structural_edits(&mut wb).unwrap());
    }
}

#[test]
fn plain_optin_rejects_nonfinite_numbers_in_source_and_live_cells() {
    for value in ["NaN", "inf", "-inf"] {
        let original = plain_fixture();
        let sheet = part(&original, SHEET).replace("<v>12.5</v>", &format!("<v>{value}</v>"));
        let mut wb = load(replace_part(original, SHEET, &sheet));
        assert!(!enable_plain_structural_edits(&mut wb).unwrap());
        assert!(!wb.active().structural_passthrough_enabled());
    }
    let mut wb = load(plain_fixture());
    assert!(enable_plain_structural_edits(&mut wb).unwrap());
    wb.active_mut()
        .set_cell_value(1, 1, CellValue::Float(f64::NAN))
        .unwrap();
    assert!(!wb.can_preserve_package());
    assert!(!enable_plain_structural_edits(&mut wb).unwrap());
    let mut output = Cursor::new(b"existing stream".to_vec());
    assert!(write_workbook(&mut output, &wb).is_err());
    assert_eq!(output.into_inner(), b"existing stream");
}

#[test]
fn plain_optin_live_content_and_stale_snapshots_fail_closed_before_writes() {
    for mutation in 0..12 {
        let original = plain_fixture();
        let mut wb = load(original.clone());
        assert!(enable_plain_structural_edits(&mut wb).unwrap());
        apply_plain_edit(&mut wb, 0);
        match mutation {
            0 => wb
                .active_mut()
                .set_cell_comment(3, 1, Some("comment".into()))
                .unwrap(),
            1 => wb
                .active_mut()
                .set_cell_hyperlink(3, 1, Some("https://example.com".into()))
                .unwrap(),
            2 => {
                wb.active_mut()
                    .set_cell_value(3, 1, CellValue::formula("1+1"))
                    .unwrap();
            }
            3 => wb.active_mut().merged_ranges.push((1, 1, 2, 2)),
            4 => wb.active_mut().unsupported_cell_content = true,
            5 => wb
                .active_mut()
                .xml_template
                .as_mut()
                .unwrap()
                .before_data
                .push_str("<cols/>"),
            6 => wb.active_mut().xml_template = None,
            7 => {
                let source = Arc::make_mut(wb.package_template.as_mut().unwrap());
                source.sheet_parts[0] = "wrong.xml".into();
            }
            8 => wb.source_styles_xml = Some("<styleSheet/>".into()),
            9 => wb.active_mut().get_cell_mut(3, 1).unwrap().row = 99,
            10 => wb.package_template = None,
            _ => wb.source_styles_xml = None,
        }
        assert!(!wb.can_preserve_package(), "mutation {mutation}");
        let snapshot = wb.clone();
        assert!(!enable_plain_structural_edits(&mut wb).unwrap());
        assert_eq!(snapshot, wb);
        let mut output = Cursor::new(b"existing stream".to_vec());
        assert!(write_workbook(&mut output, &wb).is_err());
        assert_eq!(output.into_inner(), b"existing stream");
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("source.xlsx");
        std::fs::write(&path, &original).unwrap();
        assert!(save_workbook(&path, &wb).is_err());
        assert_eq!(std::fs::read(&path).unwrap(), original);
    }
}

#[test]
fn plain_optin_rejects_native_annotations_and_bad_source_snapshots_without_changes() {
    for mutation in 0..7 {
        let mut wb = load(plain_fixture());
        match mutation {
            0 => wb
                .active_mut()
                .set_cell_comment(1, 1, Some("comment".into()))
                .unwrap(),
            1 => wb
                .active_mut()
                .set_cell_hyperlink(1, 1, Some("link".into()))
                .unwrap(),
            2 => {
                wb.active_mut()
                    .set_cell_value(1, 1, CellValue::formula("1+1"))
                    .unwrap();
            }
            3 => wb.active_mut().unsupported_cell_content = true,
            4 => wb
                .active_mut()
                .xml_template
                .as_mut()
                .unwrap()
                .before_data
                .push_str("<cols/>"),
            5 => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).styles_xml =
                    "<styleSheet/>".into()
            }
            _ => {
                Arc::make_mut(wb.package_template.as_mut().unwrap()).archive =
                    Arc::from(&b"not a ZIP"[..])
            }
        }
        let snapshot = wb.clone();
        assert!(!enable_plain_structural_edits(&mut wb).unwrap());
        assert_eq!(snapshot, wb);
    }
}

#[test]
fn plain_optin_empty_after_delete_and_restored_template_do_not_restore_source_rows() {
    let mut wb = load(plain_fixture());
    assert!(enable_plain_structural_edits(&mut wb).unwrap());
    wb.active_mut().delete_rows(1, u32::MAX).unwrap();
    assert!(wb.active().is_empty());
    for operation in 0..5 {
        apply_plain_edit(&mut wb, operation);
        let saved = save_bytes(&wb);
        assert!(attributes(&part(&saved, SHEET), "row").is_empty());
        assert_native_cells_roundtrip(&wb, &saved);
    }
    let mut wb = load(plain_fixture());
    assert!(enable_plain_structural_edits(&mut wb).unwrap());
    let frame = wb.active().xml_template.clone();
    wb.active_mut()
        .set_cell_comment(1, 1, Some("temporary".into()))
        .unwrap();
    apply_plain_edit(&mut wb, 4);
    assert!(wb.active().xml_template.is_none());
    assert!(wb.active().source_rows_changed());
    wb.active_mut().set_cell_comment(2, 2, None).unwrap();
    wb.active_mut().xml_template = frame;
    assert!(wb.can_preserve_package());
    let saved = save_bytes(&wb);
    let rows = attributes(&part(&saved, SHEET), "row");
    assert_eq!(
        rows.iter().map(|row| row["r"].as_str()).collect::<Vec<_>>(),
        ["2", "5"]
    );
    assert_native_cells_roundtrip(&wb, &saved);
}

#[test]
fn plain_optin_rejects_orphan_known_reference_parts_and_external_relationships() {
    let original = plain_fixture();
    let types = part(&original, "[Content_Types].xml");
    for kind in [
        "worksheet",
        "calcChain",
        "comments",
        "externalLink",
        "pivotTable",
        "sharedStrings",
    ] {
        let types = types.replace("</Types>", &format!(r#"<Override PartName="/orphan.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.{kind}+xml"/></Types>"#));
        let mut wb = load(replace_part(
            original.clone(),
            "[Content_Types].xml",
            &types,
        ));
        assert!(
            !enable_plain_structural_edits(&mut wb).unwrap(),
            "accepted orphan {kind}"
        );
    }
    let path = "custom/books/_rels/book.xml.rels";
    for changed in [
        part(&original, path).replace("Id=\"theme\"", "Id=\"theme\" TargetMode=\"External\""),
        part(&original, path).replace(&format!("{REL}/theme"), "urn:foreign/theme"),
        part(&original, path).replace("</Relationships>", &format!(r#"<Relationship Id="orphan" Type="{REL}/worksheet" Target="../grids/orphan.xml"/></Relationships>"#)),
    ] {
        let mut wb = load(replace_part(original.clone(), path, &changed));
        assert!(!enable_plain_structural_edits(&mut wb).unwrap());
    }
}

fn add_plain_second_sheet(bytes: Vec<u8>, unsafe_second: bool) -> Vec<u8> {
    let book = part(&bytes, BOOK).replace(
        "</sheets>",
        &format!(r#"<sheet xmlns:r="{REL}" name="Other" sheetId="2" r:id="other"/></sheets>"#),
    );
    let bytes = replace_part(bytes, BOOK, &book);
    let path = "custom/books/_rels/book.xml.rels";
    let rels = part(&bytes, path).replace("</Relationships>", &format!(r#"<Relationship Id="other" Type="{REL}/worksheet" Target="../grids/other.xml"/></Relationships>"#));
    let bytes = replace_part(bytes, path, &rels);
    let mut sheet = part(&bytes, SHEET);
    if unsafe_second {
        sheet = sheet.replace("<v>12.5</v>", "<f>Data!A1</f><v>12.5</v>");
    }
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut output = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        output
            .raw_copy_file(archive.by_index_raw(index).unwrap())
            .unwrap();
    }
    output
        .start_file("custom/grids/other.xml", SimpleFileOptions::default())
        .unwrap();
    output.write_all(sheet.as_bytes()).unwrap();
    output.finish().unwrap().into_inner()
}

#[test]
fn plain_optin_is_workbook_atomic_and_keeps_unedited_sheet_row_frames() {
    let mut unsafe_book = load(add_plain_second_sheet(plain_fixture(), true));
    // The first sheet and all live cells pass; only the second source sheet
    // fails, so a proof must not have been partially installed on the first.
    unsafe_book
        .get_sheet_mut("Other")
        .unwrap()
        .set_cell_value(1, 1, CellValue::Float(12.5))
        .unwrap();
    let snapshot = unsafe_book.clone();
    assert!(!enable_plain_structural_edits(&mut unsafe_book).unwrap());
    assert_eq!(unsafe_book, snapshot);
    assert!(unsafe_book
        .sheets()
        .iter()
        .all(|sheet| !sheet.structural_passthrough_enabled()));
    let mut wb = load(add_plain_second_sheet(plain_fixture(), false));
    assert!(enable_plain_structural_edits(&mut wb).unwrap());
    apply_plain_edit(&mut wb, 4);
    let saved = save_bytes(&wb);
    assert!(!attributes(&part(&saved, SHEET), "row")
        .iter()
        .any(|row| row["r"] == "8"));
    assert!(attributes(&part(&saved, "custom/grids/other.xml"), "row")
        .iter()
        .any(|row| row["r"] == "8"));
    assert!(!wb.get_sheet("Other").unwrap().source_rows_changed());
    wb.get_sheet_mut("Other")
        .unwrap()
        .set_cell_value(1, 1, CellValue::formula("Data!A1"))
        .unwrap();
    assert!(!wb.can_preserve_package());
    assert!(!enable_plain_structural_edits(&mut wb).unwrap());
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
