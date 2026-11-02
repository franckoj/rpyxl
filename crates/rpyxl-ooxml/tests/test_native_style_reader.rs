use std::io::{Cursor, Write};
use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::NsReader;
use rpyxl_core::cell::{CellValue, RichStyle};
use rpyxl_core::error::RpyxlError;
use rpyxl_core::worksheet::Worksheet;
use rpyxl_ooxml::writer::write_workbook_generated;
use rpyxl_ooxml::{parse_worksheet_into, SharedStringsTable, Styles, XlsxPackage};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn read_sheet(xml: &[u8], styles: &Styles) -> Worksheet {
    let mut sheet = Worksheet::new("Native").unwrap();
    parse_worksheet_into(xml, &mut sheet, &SharedStringsTable::new(), styles, false).unwrap();
    sheet
}

fn rich_styles() -> Styles {
    Styles::from_reader(format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{NS}">
        <fonts><font><name val="Calibri"/></font><font><b/><name val="日本語 😀 &amp; café"/></font></fonts>
        <fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFABCDEF"/></patternFill></fill></fills>
        <borders><border/><border><left style="thin"/></border></borders>
        <numFmts><numFmt numFmtId="164" formatCode="0.00 &quot;€&quot;"/></numFmts>
        <cellStyleXfs><xf numFmtId="164" fontId="1" fillId="1" borderId="1"><alignment horizontal="center"/><protection locked="0"/></xf></cellStyleXfs>
        <cellXfs>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          <xf xfId="0"/>
          <xf xfId="0" quotePrefix="1"/>
          <xf xfId="0" fontId="0" fillId="0" borderId="0" numFmtId="0" applyFont="false" applyFill="0" applyBorder="0" applyNumberFormat="0" applyAlignment="0" applyProtection="0"><alignment horizontal="right"/><protection locked="1"/></xf>
          <xf xfId="0" fontId="0" applyFont="1" numFmtId="14" applyAlignment="1"><alignment horizontal="right"/></xf>
        </cellXfs></styleSheet>"#
    ).as_bytes()).unwrap()
}

#[test]
fn resolves_inheritance_and_keeps_one_identity_per_original_xf() {
    let styles = rich_styles();
    assert_eq!(styles.rich_styles().len(), 5);
    assert_eq!(styles.number_format(0), Some("General"));
    assert_eq!(styles.number_format(1), Some("0.00 \"€\""));
    assert_eq!(styles.number_format(3), styles.number_format(1));
    assert!(styles.is_date_format(4));
    assert!(styles.rich_style(5).is_none());
    assert!(styles.number_format(5).is_none());

    let inherited = styles.rich_style(1).unwrap();
    assert!(inherited
        .font
        .as_deref()
        .unwrap()
        .contains("日本語 😀 &amp; café"));
    assert!(inherited.fill.as_deref().unwrap().contains("FFABCDEF"));
    assert!(inherited.border.as_deref().unwrap().contains("thin"));
    assert!(inherited.alignment.as_deref().unwrap().contains("center"));
    assert!(inherited
        .protection
        .as_deref()
        .unwrap()
        .contains("locked=\"0\""));
    assert_eq!(*inherited, *styles.rich_style(2).unwrap());
    assert_eq!(*inherited, *styles.rich_style(3).unwrap());
    assert!(!Arc::ptr_eq(&inherited, &styles.rich_style(2).unwrap()));
    assert!(Arc::ptr_eq(&inherited, &styles.rich_style(1).unwrap()));
    let overridden = styles.rich_style(4).unwrap();
    assert!(overridden.font.as_deref().unwrap().contains("Calibri"));
    assert!(overridden.alignment.as_deref().unwrap().contains("right"));
    assert_eq!(styles.clone().source_xml(), styles.source_xml());
}

#[test]
fn resolved_style_budget_bounds_aggregate_component_expansion() {
    // Less than 1 MiB of component XML expands beyond 64 MiB across repeated XFs.
    // Each component alone is below budget: all five selected fields must be charged.
    let payload = "x".repeat(128 * 1024);
    let tables = format!(
        r#"<fonts><font><!--{payload}--></font></fonts>
        <fills><fill><!--{payload}--></fill></fills>
        <borders><border><!--{payload}--></border></borders>
        <cellStyleXfs><xf><alignment><!--{payload}--></alignment><protection><!--{payload}--></protection></xf></cellStyleXfs>"#
    );
    let xml = |count| {
        format!(
            r#"<styleSheet xmlns="{NS}">{tables}<cellXfs>{}</cellXfs></styleSheet>"#,
            r#"<xf xfId="0"/>"#.repeat(count)
        )
    };
    let small = Styles::from_reader(xml(2).as_bytes()).unwrap();
    assert_eq!(small.rich_styles().len(), 2);
    let first = small.rich_style(0).unwrap();
    let second = small.rich_style(1).unwrap();
    assert_eq!(first, second);
    assert!(!Arc::ptr_eq(&first, &second));
    assert!(matches!(
        Styles::from_reader(xml(128).as_bytes()),
        Err(RpyxlError::XmlSecurityViolation(message))
            if message.contains("Resolved styles.xml allocation") && message.contains("budget")
    ));
}

#[test]
fn resolved_style_budget_ignores_overridden_inherited_components() {
    let payload = "x".repeat(1024 * 1024);
    let xml = format!(
        r#"<styleSheet xmlns="{NS}"><cellStyleXfs><xf><alignment><!--{payload}--></alignment><protection><!--{payload}--></protection></xf></cellStyleXfs><cellXfs>{}</cellXfs></styleSheet>"#,
        r#"<xf xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="right"/><protection locked="1"/></xf>"#.repeat(128)
    );
    let styles = Styles::from_reader(xml.as_bytes()).unwrap();
    assert_eq!(styles.rich_styles().len(), 128);
    for style in styles.rich_styles() {
        assert!(style.alignment.as_deref().unwrap().contains("right"));
        assert!(style
            .protection
            .as_deref()
            .unwrap()
            .contains("locked=\"1\""));
        assert!(!style.alignment.as_deref().unwrap().contains(&payload));
        assert!(!style.protection.as_deref().unwrap().contains(&payload));
    }
}

#[test]
fn component_fragments_carry_inherited_namespaces_and_local_overrides() {
    let xml = format!(
        r#"<s:styleSheet xmlns:s="{NS}" xmlns:x="urn:root" xmlns:r="{REL_NS}">
        <s:fonts xmlns:x="urn:group"><s:font><s:name val="é字"/><x:extra r:id="rId1"/></s:font></s:fonts>
        <s:cellXfs><s:xf><s:alignment xmlns:x="urn:local" x:mode="keep"/></s:xf></s:cellXfs>
        </s:styleSheet>"#
    );
    let styles = Styles::from_reader(xml.as_bytes()).unwrap();
    assert_eq!(styles.source_xml(), Some(xml.as_str()));
    let rich = styles.rich_style(0).unwrap();
    for fragment in [
        rich.font.as_deref().unwrap(),
        rich.alignment.as_deref().unwrap(),
    ] {
        assert!(fragment.contains(&format!("xmlns:s=\"{NS}\"")));
        assert!(fragment.contains(&format!("xmlns:r=\"{REL_NS}\"")));
        let mut reader = NsReader::from_str(fragment);
        loop {
            let (namespace, event) = reader.read_resolved_event().unwrap();
            if let Event::Start(_) | Event::Empty(_) = event {
                assert!(matches!(
                    namespace,
                    quick_xml::name::ResolveResult::Bound(_)
                ));
            } else if matches!(event, Event::Eof) {
                break;
            }
        }
    }
    assert!(rich
        .font
        .as_deref()
        .unwrap()
        .contains("xmlns:x=\"urn:group\""));
    assert!(rich
        .alignment
        .as_deref()
        .unwrap()
        .contains("xmlns:x=\"urn:local\""));
}

#[test]
fn opaque_extensions_cannot_add_source_xfs_or_override_style_tables() {
    let opaque = r#"<extLst><ext uri="opaque">
        <e:cellXfs><e:xf numFmtId="14"/></e:cellXfs>
        <cellXfs><xf numFmtId="14"/></cellXfs>
        <numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
        <fonts><font><name val="Nested impostor"/></font></fonts>
        <cellStyleXfs><xf numFmtId="14"/></cellStyleXfs>
        </ext></extLst>"#;
    let xml = format!(
        r#"<styleSheet xmlns="{NS}" xmlns:e="urn:opaque">
        <e:fonts><e:font><e:name val="Foreign table"/></e:font></e:fonts>
        <fonts><e:font><e:name val="Foreign entry"/></e:font><font><name val="Real font"/></font></fonts>
        <fills><e:fill/><fill><patternFill patternType="solid"/></fill></fills>
        <borders><e:border/><border><left style="thin"/></border></borders>
        <numFmts><e:numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
        <cellStyleXfs><e:xf numFmtId="14"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0"><alignment horizontal="center"/><protection locked="0"/></xf></cellStyleXfs>
        <cellXfs><e:xf numFmtId="14"/><xf xfId="0" quotePrefix="1"><e:alignment horizontal="right"/><e:protection locked="1"/></xf></cellXfs>
        {opaque}</styleSheet>"#
    );
    let styles = Styles::from_reader(xml.as_bytes()).unwrap();
    assert_eq!(styles.rich_styles().len(), 1);
    assert!(styles.rich_style(1).is_none());
    assert_eq!(styles.number_format(0), Some("0.000"));
    assert!(styles.number_format(1).is_none());
    assert!(!styles.is_date_format(0));
    let rich = styles.rich_style(0).unwrap();
    assert!(rich.font.as_deref().unwrap().contains("Real font"));
    assert!(rich.fill.as_deref().unwrap().contains("solid"));
    assert!(rich.border.as_deref().unwrap().contains("thin"));
    assert!(rich.alignment.as_deref().unwrap().contains("center"));
    assert!(rich.protection.as_deref().unwrap().contains("locked=\"0\""));
    assert_eq!(styles.source_xml(), Some(xml.as_str()));
    assert!(styles.source_xml().unwrap().contains(opaque));
}

#[test]
fn style_tables_require_a_supported_root_namespace_and_exact_ancestry() {
    let tables = r#"<fonts><font/></fonts><numFmts><numFmt numFmtId="164" formatCode="0.000"/></numFmts><cellXfs><xf numFmtId="164"/></cellXfs>"#;
    for xml in [
        format!(r#"<styleSheet xmlns="urn:foreign">{tables}</styleSheet>"#),
        format!("<wrapper><styleSheet>{tables}</styleSheet></wrapper>"),
        format!("<styleSheet><extLst><ext><styleSheet>{tables}</styleSheet></ext></extLst></styleSheet>"),
        format!(r#"<notStyleSheet xmlns="{NS}">{tables}</notStyleSheet>"#),
        format!(r#"<wrapper xmlns="{NS}"><styleSheet>{tables}</styleSheet></wrapper>"#),
        format!(
            r#"<styleSheet xmlns="{NS}"><extLst><ext uri="nested"><styleSheet>{tables}</styleSheet></ext></extLst></styleSheet>"#
        ),
    ] {
        let styles = Styles::from_reader(xml.as_bytes()).unwrap();
        assert!(styles.rich_styles().is_empty(), "{xml}");
        assert!(styles.number_format(0).is_none(), "{xml}");
        assert_eq!(styles.source_xml(), Some(xml.as_str()));
    }
}

#[test]
fn namespace_less_legacy_styles_keep_date_helpers_and_reject_foreign_lookups() {
    let builtin = Styles::from_reader(
        &br#"<styleSheet><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>"#[..],
    )
    .unwrap();
    assert_eq!(builtin.number_format(0), Some("mm-dd-yy"));
    assert!(builtin.is_date_format(0));

    let xml = format!(
        r#"<styleSheet xmlns:e="urn:foreign" xmlns:s="{NS}">
        <fonts><e:font/><s:font/><font><name val="Legacy font"/></font></fonts>
        <numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><e:numFmt numFmtId="164" formatCode="0.000"/><s:numFmt numFmtId="164" formatCode="0.000"/></numFmts>
        <e:numFmts><numFmt numFmtId="164" formatCode="0.000"/></e:numFmts>
        <cellXfs><xf numFmtId="164"/><e:xf numFmtId="0"/><s:xf numFmtId="0"/></cellXfs>
        <e:cellXfs><xf numFmtId="0"/></e:cellXfs><s:cellXfs><s:xf numFmtId="0"/></s:cellXfs>
        <extLst><ext><e:cellXfs><e:xf numFmtId="0"/></e:cellXfs><cellXfs><xf numFmtId="0"/></cellXfs><numFmts><numFmt numFmtId="164" formatCode="0.000"/></numFmts></ext></extLst>
        </styleSheet>"#
    );
    let styles = Styles::from_reader(xml.as_bytes()).unwrap();
    assert_eq!(styles.rich_styles().len(), 1);
    assert_eq!(styles.number_format(0), Some("yyyy-mm-dd"));
    assert!(styles.is_date_format(0));
    assert!(styles.number_format(1).is_none());
    assert!(styles.rich_style(1).is_none());
    assert!(styles
        .rich_style(0)
        .unwrap()
        .font
        .as_deref()
        .unwrap()
        .contains("Legacy font"));
    assert_eq!(
        styles.date_string_for_style(44927.0, 0).as_deref(),
        Some("2023-01-01 00:00:00")
    );
    assert_eq!(styles.source_xml(), Some(xml.as_str()));
    let sheet = read_sheet(
        br#"<worksheet><sheetData><row r="1"><c r="A1"><v>44927</v></c></row></sheetData></worksheet>"#,
        &styles,
    );
    assert_eq!(
        sheet.get_cell_value(1, 1),
        CellValue::DateTime("2023-01-01 00:00:00".into())
    );
}

#[test]
fn root_namespace_mode_applies_to_each_direct_table_and_record() {
    for (root_namespace, other_namespace) in [
        (String::new(), format!(r#"xmlns="{NS}""#)),
        (format!(r#"xmlns="{NS}""#), "xmlns=\"\"".to_string()),
    ] {
        let xml = format!(
            r#"<styleSheet {root_namespace}>
            <fonts><font {other_namespace}/><font><name val="Original"/></font></fonts>
            <numFmts><numFmt numFmtId="164" formatCode="0.000"/><numFmt {other_namespace} numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
            <numFmts {other_namespace}><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
            <cellXfs {other_namespace}><xf numFmtId="14"/></cellXfs>
            <cellXfs><xf numFmtId="164"/><xf {other_namespace} numFmtId="14"/></cellXfs>
            </styleSheet>"#
        );
        let styles = Styles::from_reader(xml.as_bytes()).unwrap();
        assert_eq!(styles.rich_styles().len(), 1, "{xml}");
        assert_eq!(styles.number_format(0), Some("0.000"), "{xml}");
        assert!(!styles.is_date_format(0), "{xml}");
        assert!(
            styles
                .rich_style(0)
                .unwrap()
                .font
                .as_deref()
                .unwrap()
                .contains("Original"),
            "{xml}"
        );
    }
}

#[test]
fn attaches_shared_styles_to_normal_and_blank_cells_including_style_zero() {
    let styles = rich_styles();
    let xml = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1">
        <c r="A1"><v>42</v></c><c r="B1"/><c r="C1" s="0"></c>
        <c r="D1" s="1" t="inlineStr"><is><t xml:space="preserve">  é字 😀 &amp; &#13;  </t></is></c>
        <c r="E1" s="1"/><c r="F1" s="2"><v>3</v></c><c r="G1" s="4"><v>44927</v></c>
        <c r="H1" t="d"><v>2026-10-31</v></c>
        </row></sheetData></worksheet>"#
    );
    let sheet = read_sheet(xml.as_bytes(), &styles);
    assert_eq!(sheet.cell_count(), 8);
    assert_eq!(sheet.get_cell_value(1, 1), CellValue::Int(42));
    for col in 1..=3 {
        let cell = sheet.get_cell(1, col).unwrap();
        assert_eq!(cell.source_xf_index(), Some(0));
        assert_eq!(cell.source_xf_id.unwrap().get(), 1);
        assert!(Arc::ptr_eq(
            cell.rich_style.as_ref().unwrap(),
            &styles.rich_style(0).unwrap()
        ));
        assert!(
            cell.number_format.is_none(),
            "General must not be allocated per cell"
        );
    }
    assert_eq!(
        sheet.get_cell_value(1, 4),
        CellValue::String("  é字 😀 & \r  ".into())
    );
    for col in [4, 5] {
        let cell = sheet.get_cell(1, col).unwrap();
        assert_eq!(cell.source_xf_index(), Some(1));
        assert!(Arc::ptr_eq(
            cell.rich_style.as_ref().unwrap(),
            &styles.rich_style(1).unwrap()
        ));
        assert_eq!(cell.number_format.as_deref(), Some("0.00 \"€\""));
    }
    assert_eq!(sheet.get_cell_value(1, 5), CellValue::Empty);
    assert!(!Arc::ptr_eq(
        sheet.get_cell(1, 5).unwrap().rich_style.as_ref().unwrap(),
        sheet.get_cell(1, 6).unwrap().rich_style.as_ref().unwrap()
    ));
    assert_eq!(
        sheet.get_cell_value(1, 7),
        CellValue::DateTime("2023-01-01 00:00:00".into())
    );
    assert_eq!(
        sheet.get_cell(1, 8).unwrap().number_format.as_deref(),
        Some("yyyy-mm-dd")
    );
}

#[test]
fn source_xf_provenance_uses_checked_one_based_indices() {
    let xml = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" s="4294967294"/><c r="B1" s="4294967295"/><c r="C1" s="{}"><v>7</v></c></row></sheetData></worksheet>"#,
        usize::MAX
    );
    let sheet = read_sheet(xml.as_bytes(), &Styles::default());
    assert_eq!(
        sheet.get_cell(1, 1).unwrap().source_xf_id.unwrap().get(),
        u32::MAX
    );
    assert_eq!(
        sheet.get_cell(1, 1).unwrap().source_xf_index(),
        Some(u32::MAX as usize - 1)
    );
    assert!(sheet.get_cell(1, 2).unwrap().source_xf_id.is_none());
    assert!(sheet.get_cell(1, 3).unwrap().source_xf_id.is_none());
    assert_eq!(sheet.get_cell_value(1, 3), CellValue::Int(7));
}

#[test]
fn captures_framing_and_nested_metadata_without_cell_data_or_dimension() {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}" xmlns:r="{REL_NS}" xmlns:x="urn:extension">
        <dimension ref="A1:Z99"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" state="frozen"/></sheetView></sheetViews>
        <cols><col min="1" max="1" width="23"/></cols><!-- before -->
        <sheetData custom="original"><row r="1"><c r="A1"><v>17</v></c></row></sheetData>
        <mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>
        <hyperlinks><hyperlink ref="A1" r:id="rId9"/></hyperlinks><drawing r:id="rId10"/>
        <extLst><ext uri="keep"><x:container><x:c r="Z99"><x:t>not a cell é字</x:t></x:c></x:container></ext></extLst>
        <!-- after --></worksheet>"#
    );
    let sheet = read_sheet(xml.as_bytes(), &Styles::default());
    assert_eq!(sheet.cell_count(), 1);
    let frame = sheet.xml_template.as_ref().unwrap();
    assert!(frame.root_start.starts_with("<?xml"));
    assert!(frame.root_start.contains(&format!("xmlns:r=\"{REL_NS}\"")));
    assert!(!frame.before_data.contains("dimension"));
    assert!(frame
        .before_data
        .contains("<pane xSplit=\"1\" state=\"frozen\"/>"));
    assert!(frame.before_data.contains("<!-- before -->"));
    assert_eq!(frame.data_start, "<sheetData custom=\"original\">");
    assert!(!frame.after_data.contains("<row"));
    assert!(!frame.after_data.contains("</sheetData>"));
    assert!(frame.after_data.contains("r:id=\"rId9\""));
    assert!(frame.after_data.contains("not a cell é字"));
    assert!(frame.after_data.ends_with("</worksheet>"));
    let reconstructed = format!(
        "{}{}{}{}</sheetData>{}",
        frame.root_start,
        frame.before_data,
        frame.data_start,
        "<row r=\"1\"><c r=\"A1\"><v>17</v></c></row>",
        frame.after_data
    );
    assert_eq!(
        read_sheet(reconstructed.as_bytes(), &Styles::default()).get_cell_value(1, 1),
        CellValue::Int(17)
    );
}

#[test]
fn only_safe_complete_utf8_default_namespace_sheets_get_templates() {
    for xml in [
        format!(r#"<worksheet xmlns="{NS}"><sheetData/></worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}"><sheetData></sheetData></worksheet>"#),
    ] {
        let sheet = read_sheet(xml.as_bytes(), &Styles::default());
        let frame = sheet.xml_template.unwrap();
        assert_eq!(frame.data_start, "<sheetData>");
        assert_eq!(frame.after_data, "</worksheet>");
    }
    for xml in [
        "<worksheet><sheetData/></worksheet>".to_string(),
        format!(
            r#"<s:worksheet xmlns:s="{NS}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>7</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
        ),
        format!(
            r#"<?xml version="1.0" encoding="ISO-8859-1"?><worksheet xmlns="{NS}"><sheetData/></worksheet>"#
        ),
        format!(r#"<worksheet xmlns="{NS}"><sheetData/>"#),
        format!(r#"<worksheet xmlns="{NS}"><sheetData/><sheetData/></worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}"><sheetData xmlns="urn:other"/></worksheet>"#),
        format!(
            r#"<!-- before --><?xml version="1.0"?><worksheet xmlns="{NS}"><sheetData/></worksheet>"#
        ),
        format!(r#"<worksheet xmlns="{NS}"><sheetData/></worksheet>not XML whitespace"#),
    ] {
        let sheet = read_sheet(xml.as_bytes(), &Styles::default());
        assert!(sheet.xml_template.is_none(), "{xml}");
        if xml.contains("s:c") {
            assert_eq!(sheet.get_cell_value(1, 1), CellValue::Int(7));
        }
    }
}

#[test]
fn normal_reads_use_cell_coordinates_even_when_first_row_and_dimension_disagree() {
    let xml = format!(
        r#"<worksheet xmlns="{NS}"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="C7"><v>42</v></c></row></sheetData></worksheet>"#
    );
    let sheet = read_sheet(xml.as_bytes(), &Styles::default());
    assert_eq!(sheet.cell_count(), 1);
    assert_eq!(sheet.get_cell_value(7, 3), CellValue::Int(42));
    assert_eq!(sheet.get_cell_value(1, 1), CellValue::Empty);
    assert_eq!(sheet.dimensions().unwrap(), "C7");
    assert!(sheet.xml_template.is_some());
}

#[test]
fn legacy_encoding_is_decoded_for_styles_and_values_without_a_raw_template() {
    let mut xml = b"<?xml version=\"1.0\" encoding=\"windows-1252\"?><styleSheet><fonts><font><name val=\"caf".to_vec();
    xml.push(0xe9);
    xml.extend_from_slice(b"\"/></font></fonts><cellXfs><xf/></cellXfs></styleSheet>");
    let styles = Styles::from_reader(xml.as_slice()).unwrap();
    let retained = styles.source_xml().unwrap();
    assert!(retained.starts_with(r#"<?xml version="1.0" encoding="UTF-8"?>"#));
    assert!(retained.contains("café"));
    let reparsed = Styles::from_reader(retained.as_bytes()).unwrap();
    assert_eq!(styles.rich_styles(), reparsed.rich_styles());
    assert!(styles
        .rich_style(0)
        .unwrap()
        .font
        .as_deref()
        .unwrap()
        .contains("café"));
    let mut xml = format!(r#"<?xml version="1.0" encoding="windows-1252"?><worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>caf"#).into_bytes();
    xml.push(0xe9);
    xml.extend_from_slice(b"</t></is></c></row></sheetData></worksheet>");
    let sheet = read_sheet(&xml, &styles);
    assert_eq!(sheet.get_cell_value(1, 1), CellValue::String("café".into()));
    assert!(sheet.xml_template.is_none());
}

#[test]
fn legacy_styles_and_formats_survive_generated_package_roundtrip() {
    for encoding in ["windows-1252", "ISO-8859-1"] {
        let mut styles_xml = format!(
            r#"<?xml version="1.0" encoding="{encoding}" standalone="yes"?><styleSheet xmlns="{NS}"><fonts><font><name val="caf"#
        ).into_bytes();
        styles_xml.push(0xe9);
        styles_xml.extend_from_slice(
            b"\"/></font></fonts><fills><fill><patternFill patternType=\"none\"/></fill></fills><borders><border/></borders><numFmts><numFmt numFmtId=\"164\" formatCode=\"0.00 &quot;caf",
        );
        styles_xml.push(0xe9);
        styles_xml.extend_from_slice(
            b"&quot;\"/></numFmts><cellXfs><xf/><xf numFmtId=\"164\"/></cellXfs></styleSheet>",
        );
        let mut sheet_xml = format!(
            r#"<?xml version="1.0" encoding="{encoding}"?><worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>caf"#
        ).into_bytes();
        sheet_xml.push(0xe9);
        sheet_xml.extend_from_slice(b"</t></is></c><c r=\"B1\" s=\"1\"><v>7</v></c><c r=\"C1\" s=\"1\"/></row></sheetData></worksheet>");
        let workbook_xml = format!(
            r#"<workbook xmlns="{NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Native" sheetId="1" r:id="s1"/></sheets></workbook>"#
        );
        let rels = format!(
            r#"<Relationships><Relationship Id="s1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="style" Type="{REL_NS}/styles" Target="styles.xml"/></Relationships>"#
        );
        let source = package_bytes(&[
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", rels.as_bytes()),
            ("xl/styles.xml", &styles_xml),
            ("xl/worksheets/sheet1.xml", &sheet_xml),
        ]);
        let mut package = XlsxPackage::from_reader(source, false, false).unwrap();
        let book = package.to_workbook().unwrap();
        assert!(book.package_template.is_none());
        assert!(book.active().xml_template.is_none());
        let retained = book.source_styles_xml.as_deref().unwrap();
        assert!(retained.starts_with(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#));
        let retained = Styles::from_reader(retained.as_bytes()).unwrap();
        assert_eq!(retained.number_format(1), Some("0.00 \"café\""));

        let mut saved = Cursor::new(Vec::new());
        write_workbook_generated(&mut saved, &book).unwrap();
        saved.set_position(0);
        let mut reopened = XlsxPackage::from_reader(saved, false, false).unwrap();
        assert_eq!(reopened.styles.number_format(1), Some("0.00 \"café\""));
        let reopened = reopened.to_workbook().unwrap();
        assert_eq!(book.source_styles(), reopened.source_styles());
        assert_eq!(
            reopened.active().get_cell_value(1, 1),
            CellValue::String("café".into())
        );
        assert_eq!(reopened.active().get_cell_value(1, 2), CellValue::Int(7));
        assert_eq!(reopened.active().get_cell_value(1, 3), CellValue::Empty);
        for col in 1..=3 {
            assert_eq!(
                reopened
                    .active()
                    .get_cell(1, col)
                    .unwrap()
                    .source_xf_index(),
                Some(1)
            );
        }
    }
}

#[test]
fn merged_cleanup_is_sparse_preserves_anchor_styles_and_original_metadata() {
    let styles = rich_styles();
    let xml = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" s="1" t="str"><v>anchor</v></c><c r="B1" s="1"><f>A1</f><v>9</v></c></row><row r="1048576"><c r="XFD1048576" s="2"><v>99</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells></worksheet>"#
    );
    let sheet = read_sheet(xml.as_bytes(), &styles);
    assert_eq!(sheet.cell_count(), 3);
    assert_eq!(
        sheet.get_cell_value(1, 1),
        CellValue::String("anchor".into())
    );
    assert_eq!(sheet.get_cell_value(1, 2), CellValue::Empty);
    assert_eq!(sheet.get_cell_value(1_048_576, 16_384), CellValue::Empty);
    assert_eq!(sheet.merged_ranges, vec![(1, 1, 1_048_576, 16_384)]);
    assert_eq!(sheet.get_cell(1, 2).unwrap().source_xf_index(), Some(1));
    assert!(Arc::ptr_eq(
        sheet.get_cell(1, 2).unwrap().rich_style.as_ref().unwrap(),
        &styles.rich_style(1).unwrap()
    ));
    assert!(sheet
        .xml_template
        .unwrap()
        .after_data
        .contains("ref=\"A1:XFD1048576\""));
}

fn package(parts: &[(&str, &str)]) -> Cursor<Vec<u8>> {
    let parts: Vec<_> = parts
        .iter()
        .map(|&(path, xml)| (path, xml.as_bytes()))
        .collect();
    package_bytes(&parts)
}

fn package_bytes(parts: &[(&str, &[u8])]) -> Cursor<Vec<u8>> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for &(path, xml) in parts {
        zip.start_file(path, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(xml).unwrap();
    }
    zip.finish().unwrap()
}

#[test]
fn package_exposes_resolved_paths_and_shares_source_style_identities_across_sheets() {
    let root_rels = format!(
        r#"<Relationships><Relationship Id="book" Type="{REL_NS}/officeDocument" Target="/custom/books/main.xml"/></Relationships>"#
    );
    let rels = format!(
        r#"<Relationships><Relationship Id="s1" Type="{REL_NS}/worksheet" Target="../../sheets/first.xml"/><Relationship Id="s2" Type="{REL_NS}/worksheet" Target="/sheets/second.xml"/><Relationship Id="style" Type="{REL_NS}/styles" Target="../assets/format.xml"/></Relationships>"#
    );
    let workbook = r#"<workbook xmlns:r="relationships"><sheets><sheet name="First" sheetId="1" r:id="s1"/><sheet name="Second" sheetId="2" r:id="s2"/></sheets></workbook>"#;
    let styles_xml = rich_styles().source_xml().unwrap().to_string();
    let sheet_xml = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" s="1"><v>5</v></c><c r="B1" s="2"/></row></sheetData></worksheet>"#
    );
    let mut package = XlsxPackage::from_reader(
        package(&[
            ("_rels/.rels", &root_rels),
            ("custom/books/main.xml", workbook),
            ("custom/books/_rels/main.xml.rels", &rels),
            ("custom/assets/format.xml", &styles_xml),
            ("sheets/first.xml", &sheet_xml),
            ("sheets/second.xml", &sheet_xml),
        ]),
        false,
        false,
    )
    .unwrap();
    assert_eq!(package.workbook_path, "custom/books/main.xml");
    assert_eq!(
        package.styles_path.as_deref(),
        Some("custom/assets/format.xml")
    );
    let mut book = package.to_workbook().unwrap();
    assert_eq!(book.source_styles().len(), 5);
    assert!(book.package_template.is_none());
    assert_eq!(book.source_styles_xml.as_deref(), Some(styles_xml.as_str()));
    for name in ["First", "Second"] {
        let sheet = book.get_sheet(name).unwrap();
        for (col, index) in [(1, 1), (2, 2)] {
            let cell = sheet.get_cell(1, col).unwrap();
            assert_eq!(cell.source_xf_index(), Some(index));
            let style = cell.rich_style.as_ref().unwrap();
            assert!(Arc::ptr_eq(style, &book.source_styles()[index]));
            assert_eq!(book.source_style_index(style), Some(index));
        }
        assert!(sheet.xml_template.is_some());
    }
    let replacement = Arc::new(RichStyle {
        alignment: Some("<alignment horizontal=\"right\"/>".into()),
        ..RichStyle::default()
    });
    let cell = book
        .get_sheet_mut("Second")
        .unwrap()
        .get_cell_mut(1, 2)
        .unwrap();
    cell.rich_style = Some(Arc::clone(&replacement));
    assert_eq!(cell.source_xf_index(), Some(2));
    assert_eq!(book.source_style_index(&replacement), None);
    assert_eq!(book.source_styles_xml.as_deref(), Some(styles_xml.as_str()));

    let mut without_styles =
        XlsxPackage::from_reader(package_without_styles(), false, false).unwrap();
    assert!(without_styles.styles_path.is_none());
    let without_styles = without_styles.to_workbook().unwrap();
    assert!(without_styles.source_styles().is_empty());
    assert!(without_styles.source_styles_xml.is_none());
}

fn package_without_styles() -> Cursor<Vec<u8>> {
    package(&[("xl/workbook.xml", "<workbook><sheets/></workbook>")])
}

#[test]
fn rejects_style_dtds_and_preserves_normal_mismatched_first_cell_errors() {
    for xml in [
        "<!DOCTYPE styleSheet><styleSheet><cellXfs><xf/></cellXfs></styleSheet>",
        "<styleSheet><fonts><font><!DOCTYPE font></font></fonts></styleSheet>",
    ] {
        assert!(matches!(
            Styles::from_reader(xml.as_bytes()),
            Err(RpyxlError::XmlSecurityViolation(_))
        ));
    }
    let xml = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>1</v></row></sheetData></worksheet>"#
    );
    let mut sheet = Worksheet::new("Broken").unwrap();
    assert!(parse_worksheet_into(
        xml.as_bytes(),
        &mut sheet,
        &SharedStringsTable::new(),
        &Styles::default(),
        false
    )
    .is_err());
    assert!(sheet.xml_template.is_none());
}
