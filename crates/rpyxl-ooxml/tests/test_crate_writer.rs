use std::collections::HashMap;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::num::NonZeroU32;
use std::rc::Rc;
use std::sync::Arc;

use quick_xml::events::Event;
use rpyxl_core::cell::{CellValue, RichStyle};
use rpyxl_core::workbook::Workbook;
use rpyxl_ooxml::crate_writer::{
    save_workbook_rust_xlsxwriter as save, write_workbook_rust_xlsxwriter as write,
};
use rpyxl_ooxml::reader::XlsxPackage;
use tempfile::{tempdir, NamedTempFile};
use zip::ZipArchive;

type TestResult = Result<(), Box<dyn std::error::Error>>;

fn bytes(book: &Workbook) -> Vec<u8> {
    let mut stream = Cursor::new(Vec::new());
    write(&mut stream, book).unwrap();
    stream.into_inner()
}

fn part(bytes: &[u8], name: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    xml
}

fn read(bytes: &[u8], data_only: bool) -> Workbook {
    let file = NamedTempFile::new().unwrap();
    std::fs::write(file.path(), bytes).unwrap();
    XlsxPackage::open(file.path(), false, data_only)
        .unwrap()
        .to_workbook()
        .unwrap()
}

fn attributes(xml: &str, element: &[u8]) -> Vec<HashMap<String, String>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut result = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == element => {
                result.push(
                    e.attributes()
                        .map(|a| {
                            let a = a.unwrap();
                            (
                                String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                                xml_normalized_attribute(a.value.as_ref()),
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

// Apply XML 1.0 EOL/attribute-whitespace normalization before expanding
// character references. quick-xml's raw attribute decoder does not do this;
// independent expectations must not accidentally bless literal CR/tab output.
fn xml_normalized_attribute(raw: &[u8]) -> String {
    let normalized = std::str::from_utf8(raw)
        .unwrap()
        .replace("\r\n", "\n")
        .replace(['\r', '\n', '\t'], " ");
    quick_xml::escape::unescape(&normalized)
        .unwrap()
        .into_owned()
}

fn cell_attrs(xml: &str, coordinate: &str) -> HashMap<String, String> {
    attributes(xml, b"c")
        .into_iter()
        .find(|a| a.get("r").is_some_and(|r| r == coordinate))
        .unwrap()
}

fn assert_number(value: &CellValue, expected: f64) {
    match value {
        CellValue::Int(n) => assert_eq!(*n as f64, expected),
        CellValue::Float(n) => assert_eq!(*n, expected),
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

// Assert all three destination modes, including non-zero stream position. Put
// invalid input late in a workbook to catch validation that happens mid-write.
fn assert_rejected(book: &Workbook, reason: &str) {
    let directory = tempdir().unwrap();
    let existing = directory.path().join("existing.xlsx");
    std::fs::write(&existing, b"do not truncate").unwrap();
    let error = save(&existing, book).unwrap_err();
    assert!(error.to_string().contains(reason), "{error}");
    assert_eq!(std::fs::read(existing).unwrap(), b"do not truncate");
    let absent = directory.path().join("absent.xlsx");
    assert!(save(&absent, book).is_err());
    assert!(!absent.exists());
    let mut stream = Cursor::new(b"prefix and existing content".to_vec());
    stream.set_position(7);
    assert!(write(&mut stream, book).is_err());
    assert_eq!(stream.position(), 7);
    assert_eq!(stream.into_inner(), b"prefix and existing content");
}

#[test]
fn values_sparse_sheets_literal_strings_and_active_sheet() -> TestResult {
    let mut book = Workbook::new();
    let values = [
        CellValue::String("=not a formula".into()),
        CellValue::String("https://example.com/a?b=c&d=e".into()),
        CellValue::String("mailto:test@example.com".into()),
        CellValue::Int(-42),
        CellValue::Float(123.125),
        CellValue::Bool(true),
        CellValue::Bool(false),
        CellValue::Int(9_007_199_254_740_992),
        CellValue::Int(-9_007_199_254_740_992),
        CellValue::Int(9_007_199_254_740_994),
        CellValue::String("Unicode: 日本語 🦀 & < >".into()),
    ];
    // Deliberately insert out of order; serialized coordinates must not depend
    // on the original edit order.
    for (i, value) in values.iter().enumerate().rev() {
        book.active_mut()
            .set_cell_value(i as u32 + 1, 1, value.clone())?;
    }
    book.create_sheet(Some("Second"), None)?.set_cell_value(
        1000,
        16384,
        CellValue::String("last column".into()),
    )?;
    book.create_sheet(Some("Empty"), None)?;
    book.set_active_index(1)?;
    let original = book.clone();
    let output = bytes(&book);
    assert_eq!(book, original);
    let loaded = read(&output, false);
    assert_eq!(loaded.sheetnames(), ["Sheet", "Second", "Empty"]);
    assert_eq!(loaded.active_index(), 1);
    for (i, value) in values.iter().enumerate() {
        assert_eq!(
            &loaded.sheets()[0].get_cell(i as u32 + 1, 1).unwrap().value,
            value
        );
    }
    assert_eq!(
        loaded.sheets()[1].get_cell_value(1000, 16384),
        CellValue::String("last column".into())
    );
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    assert!(!xml.contains("<f>"));
    assert!(!xml.contains("<hyperlink"));
    assert_eq!(cell_attrs(&xml, "A1")["t"], "s");
    assert!(part(&output, "xl/sharedStrings.xml").contains("=not a formula"));
    assert!(xml.contains("<v>-42</v>"));
    assert!(xml.contains("<v>9007199254740994</v>"));
    assert!(part(&output, "xl/worksheets/sheet2.xml").contains("XFD1000"));
    Ok(())
}

#[test]
fn whitespace_cr_and_formula_caches_are_faithful() -> TestResult {
    let mut book = Workbook::new();
    let values = [
        CellValue::String("  leading and trailing  ".into()),
        CellValue::String("\ttab\nnewline\t".into()),
        CellValue::String("\rleading\r\nline2\r".into()),
        CellValue::formula("=SUM(1,2)"),
        CellValue::formula_with_cached("=1+2", CellValue::Int(3)),
        CellValue::formula_with_cached("=1/4", CellValue::Float(0.25)),
        CellValue::formula_with_cached("=TRUE()", CellValue::Bool(true)),
        CellValue::formula_with_cached("=FALSE()", CellValue::Bool(false)),
        CellValue::formula_with_cached("=\"answer\"", CellValue::String(" answer\r\n ".into())),
        // Formula text is not an inline-string escape target.
        CellValue::formula_with_cached("=\"_x000D_\"", CellValue::String("literal formula".into())),
    ];
    for (i, value) in values.iter().enumerate() {
        book.active_mut()
            .set_cell_value(i as u32 + 1, 1, value.clone())?;
    }
    let output = bytes(&book);
    let loaded = read(&output, false);
    for (i, value) in values.iter().enumerate() {
        assert_eq!(
            &loaded.active().get_cell(i as u32 + 1, 1).unwrap().value,
            value
        );
    }
    let data = read(&output, true);
    assert_eq!(data.active().get_cell_value(4, 1), CellValue::Empty);
    assert_eq!(
        data.active().get_resolved_cell_value(4, 1, true),
        CellValue::Empty
    );
    assert_number(&data.active().get_cell_value(5, 1), 3.0);
    assert_eq!(data.active().get_cell_value(7, 1), CellValue::Bool(true));
    assert_eq!(data.active().get_cell_value(8, 1), CellValue::Bool(false));
    assert_eq!(
        data.active().get_cell_value(9, 1),
        CellValue::String(" answer\r\n ".into())
    );
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    let shared = part(&output, "xl/sharedStrings.xml");
    assert!(shared.contains("&#13;leading&#13;\nline2&#13;"));
    assert!(shared.contains("xml:space=\"preserve\""));
    assert!(!shared.contains("_x000D_"));
    assert!(xml.contains("<f>SUM(1,2)</f><v></v>"));
    assert!(!cell_attrs(&xml, "A4").contains_key("t"));
    assert_eq!(cell_attrs(&xml, "A7")["t"], "b");
    assert_eq!(cell_attrs(&xml, "A9")["t"], "str");
    assert!(xml.contains("<f>\"_x000D_\"</f>"));
    assert!(!part(&output, "xl/worksheets/sheet1.xml").contains("t=\"array\""));
    Ok(())
}

#[test]
fn uncached_formula_fixup_also_runs_without_cr() -> TestResult {
    let mut book = Workbook::new();
    book.active_mut()
        .set_cell_value(1, 1, CellValue::formula("=1+1"))?;
    let output = bytes(&book);
    assert_eq!(
        read(&output, false).active().get_cell_value(1, 1),
        CellValue::formula("=1+1")
    );
    assert_eq!(
        read(&output, true).active().get_cell_value(1, 1),
        CellValue::Empty
    );
    assert!(!cell_attrs(&part(&output, "xl/worksheets/sheet1.xml"), "A1").contains_key("t"));
    Ok(())
}

#[test]
fn temporal_defaults_effective_formats_and_1900_leap_boundary() -> TestResult {
    let mut book = Workbook::new();
    for (row, value) in [
        CellValue::Date("1900-01-01".into()),
        CellValue::Date("1900-02-28".into()),
        CellValue::Date("1900-03-01".into()),
        CellValue::DateTime("2026-11-02T03:04:05.125".into()),
        CellValue::DateTime("12:30:45.250".into()),
        CellValue::formula_with_cached("=DATE(2026,11,2)", CellValue::Date("2026-11-02".into())),
    ]
    .into_iter()
    .enumerate()
    {
        book.active_mut().set_cell_value(row as u32 + 1, 1, value)?;
        book.active_mut()
            .set_cell_number_format(row as u32 + 1, 1, Some("General".into()))?;
    }
    book.active_mut()
        .set_cell_number_format(4, 1, Some("yyyy-mm-dd hh:mm:ss.000".into()))?;
    let output = bytes(&book);
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("<v>1</v>"));
    assert!(xml.contains("<v>59</v>"));
    assert!(xml.contains("<v>61</v>"));
    let loaded = read(&output, false);
    for row in 1..=6 {
        let cell = loaded.active().get_cell(row, 1).unwrap();
        assert!(cell
            .number_format
            .as_deref()
            .unwrap()
            .contains(if row == 5 { "h:mm:ss" } else { "yyyy-mm-dd" }));
    }
    assert_eq!(
        loaded.active().get_cell_value(1, 1),
        CellValue::DateTime("1900-01-01 00:00:00".into())
    );
    assert_eq!(
        loaded.active().get_cell_value(4, 1),
        CellValue::DateTime("2026-11-02 03:04:05.125".into())
    );
    assert_eq!(
        loaded.active().get_cell_value(5, 1),
        CellValue::DateTime("12:30:45.250".into())
    );
    assert!(matches!(
        read(&output, true).active().get_cell_value(6, 1),
        CellValue::DateTime(_)
    ));
    Ok(())
}

#[test]
fn styled_blank_and_other_modeled_font_alignment_fields() -> TestResult {
    let mut book = Workbook::new();
    book.active_mut().set_cell_value(2, 3, CellValue::Empty)?;
    let cell = book.active_mut().get_cell_mut(2, 3).unwrap();
    cell.number_format = Some("#,##0.00 \"USD & <tax>\"".into());
    cell.rich_style = Some(Arc::new(RichStyle {
        font: Some("<font><name val=\"Times New Roman\"/><sz val=\"10.5\"/><family val=\"1.0\"/><charset val=\"1\"/><u val=\"doubleAccounting\"/><strike/><vertAlign val=\"superscript\"/></font>".into()),
        alignment: Some("<alignment horizontal=\"right\" vertical=\"top\" indent=\"2.0\" textRotation=\"255\" shrinkToFit=\"true\" readingOrder=\"1\"/>".into()),
        protection: Some("<protection/>".into()),
        ..RichStyle::default()
    }));
    let output = bytes(&book);
    let loaded = read(&output, false);
    let blank = loaded
        .active()
        .get_cell(2, 3)
        .expect("styled blank must materialize");
    assert_eq!(blank.value, CellValue::Empty);
    assert_eq!(
        blank.number_format.as_deref(),
        Some("#,##0.00 \"USD & <tax>\"")
    );
    let styles = part(&output, "xl/styles.xml");
    assert_eq!(attributes(&styles, b"charset")[0]["val"], "1");
    assert_eq!(attributes(&styles, b"vertAlign")[0]["val"], "superscript");
    assert_eq!(attributes(&styles, b"u")[0]["val"], "doubleAccounting");
    let alignment = &attributes(&styles, b"alignment")[0];
    assert_eq!(alignment["horizontal"], "right");
    assert_eq!(alignment["indent"], "2");
    assert_eq!(alignment["textRotation"], "255");
    assert_eq!(alignment["shrinkToFit"], "1");
    Ok(())
}

#[test]
fn string_size_boundaries_and_empty_new_workbook() -> TestResult {
    let empty = bytes(&Workbook::new());
    assert_eq!(read(&empty, false).sheetnames(), ["Sheet"]);
    let mut book = Workbook::new();
    book.active_mut()
        .set_cell_value(1, 1, CellValue::String("a".repeat(32767)))?;
    assert_eq!(
        read(&bytes(&book), false).active().get_cell_value(1, 1),
        CellValue::String("a".repeat(32767))
    );
    book.active_mut()
        .set_cell_value(1, 1, CellValue::String("🦀".repeat(16384)))?;
    assert_rejected(&book, "limit");
    Ok(())
}

fn common_style() -> RichStyle {
    RichStyle {
        font: Some("<font><name val=\"Arial\"/><sz val=\"12\"/><b/><i val=\"true\"/><u val=\"single\"/><color rgb=\"FF112233\"/></font>".into()),
        fill: Some("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFCCEEFF\"/></patternFill></fill>".into()),
        border: Some("<border outline=\"true\" diagonalUp=\"false\" diagonalDown=\"0\"><left style=\"thin\"><color rgb=\"FF445566\"/></left><right style=\"double\"/></border>".into()),
        alignment: Some("<alignment horizontal=\"center\" vertical=\"center\" textRotation=\"135\" wrapText=\"true\" shrinkToFit=\"false\" indent=\"0.0\" relativeIndent=\"0.0\" readingOrder=\"2.0\" justifyLastLine=\"false\"/>".into()),
        protection: Some("<protection locked=\"false\" hidden=\"true\"/>".into()),
    }
}

#[test]
fn styles_compile_deduplicate_and_roundtrip_without_losing_argb() -> TestResult {
    let mut book = Workbook::new();
    let first = common_style();
    let mut equivalent = first.clone();
    equivalent.font = Some("<s:font xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><s:color rgb=\"FF112233\"/><s:sz val=\"12.0\"/><s:name val=\"Arial\"/><s:b val=\"1\"/><s:i/><s:u/></s:font>".into());
    for row in 1..=100 {
        book.active_mut()
            .set_cell_value(row, 1, CellValue::Float(12.5))?;
        let cell = book.active_mut().get_cell_mut(row, 1).unwrap();
        // Distinct Arc allocations with equal XML and equivalent XML all share
        // the crate's final XF. A second number format creates only one more.
        cell.rich_style = Some(Arc::new(if row % 2 == 0 {
            first.clone()
        } else {
            equivalent.clone()
        }));
        cell.number_format = Some(if row <= 50 { "0.00" } else { "0.000" }.into());
    }
    let output = bytes(&book);
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    let cells = attributes(&xml, b"c");
    assert!(cells[..50].iter().all(|c| c["s"] == cells[0]["s"]));
    assert!(cells[50..].iter().all(|c| c["s"] == cells[50]["s"]));
    assert_ne!(cells[0]["s"], cells[50]["s"]);
    let styles = part(&output, "xl/styles.xml");
    assert_eq!(attributes(&styles, b"cellXfs")[0]["count"], "3");
    assert!(styles.contains("rgb=\"FF112233\""));
    assert!(styles.contains("rgb=\"FFCCEEFF\""));
    assert!(styles.contains("rgb=\"FF445566\""));
    assert!(styles.contains("textRotation=\"135\""));
    assert!(styles.contains("locked=\"0\""));
    assert!(styles.contains("hidden=\"1\""));
    let loaded = read(&output, false);
    let cell = loaded.active().get_cell(1, 1).unwrap();
    assert_eq!(cell.number_format.as_deref(), Some("0.00"));
    let rich = cell.rich_style.as_ref().unwrap();
    assert!(rich.font.as_ref().unwrap().contains("FF112233"));
    assert!(rich.fill.as_ref().unwrap().contains("FFCCEEFF"));
    assert!(rich.border.as_ref().unwrap().contains("FF445566"));
    assert!(rich
        .alignment
        .as_ref()
        .unwrap()
        .contains("readingOrder=\"2\""));
    assert!(rich.protection.as_ref().unwrap().contains("locked=\"0\""));
    Ok(())
}

#[test]
fn modeled_false_none_and_default_metadata_are_accepted() -> TestResult {
    let mut book = Workbook::new();
    let style = RichStyle {
        font: Some("<font><name val=\"Calibri\"/><sz val=\"11\"/><b val=\"false\"/><i val=\"0\"/><strike val=\"false\"/><outline val=\"false\"/><shadow val=\"0\"/><condense val=\"false\"/><extend val=\"false\"/></font>".into()),
        fill: Some("<fill><patternFill/></fill>".into()),
        border: Some("<border outline=\"1\" diagonalUp=\"false\" diagonalDown=\"false\"/>".into()),
        alignment: Some("<alignment textRotation=\"0\" indent=\"0.0\" relativeIndent=\"0.0\" readingOrder=\"0.0\"/>".into()),
        protection: Some("<protection locked=\"true\" hidden=\"false\"/>".into()),
    };
    book.active_mut()
        .set_cell_value(1, 1, CellValue::String("default models".into()))?;
    book.active_mut().get_cell_mut(1, 1).unwrap().rich_style = Some(Arc::new(style));
    let output = bytes(&book);
    let styles = part(&output, "xl/styles.xml");
    for tag in [
        b"b".as_slice(),
        b"i",
        b"strike",
        b"outline",
        b"shadow",
        b"scheme",
        b"color",
    ] {
        assert!(
            attributes(&styles, tag).is_empty(),
            "unexpected {} in {styles}",
            String::from_utf8_lossy(tag)
        );
    }
    assert!(!styles.contains("locked=\"0\""));
    assert!(!styles.contains("wrapText=\"1\""));
    assert_eq!(
        read(&output, false).active().get_cell_value(1, 1),
        CellValue::String("default models".into())
    );
    Ok(())
}

#[test]
fn repeated_saves_do_not_mutate_book_and_stream_need_not_be_send() -> TestResult {
    struct LocalStream {
        inner: Cursor<Vec<u8>>,
        _local: Rc<()>,
    }
    impl Write for LocalStream {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.inner.write(bytes)
        }
        fn flush(&mut self) -> std::io::Result<()> {
            self.inner.flush()
        }
    }
    impl Seek for LocalStream {
        fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
            self.inner.seek(pos)
        }
    }
    let mut book = Workbook::new();
    book.active_mut()
        .set_cell_value(1, 1, CellValue::String("one\r\ntwo".into()))?;
    let original = book.clone();
    let directory = tempdir()?;
    let path = directory.path().join("book.xlsx");
    for _ in 0..3 {
        save(&path, &book)?;
        let output = std::fs::read(&path)?;
        assert_eq!(
            read(&output, false).active().get_cell_value(1, 1),
            CellValue::String("one\r\ntwo".into())
        );
        assert_eq!(book, original);
    }
    let mut local = LocalStream {
        inner: Cursor::new(b"prefix".to_vec()),
        _local: Rc::new(()),
    };
    local.seek(SeekFrom::End(0))?;
    write(&mut local, &book)?;
    local.write_all(b"still open")?;
    assert_eq!(&local.inner.get_ref()[..6], b"prefix");
    assert_eq!(book, original);
    // Loaded books must never be accepted even if this backend produced them.
    assert_rejected(&read(&bytes(&book), false), "provenance");
    Ok(())
}

#[test]
fn value_validation_is_atomic() -> TestResult {
    for value in [
        CellValue::Int(9_007_199_254_740_993),
        CellValue::Int(-9_007_199_254_740_993),
        CellValue::Int(i64::MAX),
        CellValue::Int(i64::MIN),
        CellValue::Int(1_i64 << 60),
        CellValue::Float(f64::NAN),
        CellValue::Float(f64::INFINITY),
        CellValue::Float(f64::NEG_INFINITY),
        CellValue::String(String::new()),
        CellValue::String("literal _x000D_".into()),
        CellValue::String("literal _x005f_".into()),
        CellValue::String("bad\u{1}".into()),
        CellValue::String("a".repeat(32768)),
        CellValue::Date("not-a-date".into()),
        CellValue::DateTime("2026-11-02T12:00:00+01:00".into()),
        CellValue::DateTime("23:59:60".into()),
        CellValue::formula("=FILTER(A1:A2,A1:A2>0)"),
        CellValue::formula("=XLOOKUP(A1,A2,A3)"),
        CellValue::formula("={1,2}"),
        CellValue::formula("="),
        CellValue::formula("=\"unclosed"),
        CellValue::formula_with_cached("=1", CellValue::Empty),
        CellValue::formula_with_cached("=1", CellValue::formula("=1")),
        CellValue::formula_with_cached("=1", CellValue::Int(i64::MAX)),
        CellValue::formula_with_cached("=1", CellValue::Float(f64::NAN)),
    ] {
        let mut book = Workbook::new();
        book.active_mut()
            .set_cell_value(1, 1, CellValue::String("valid first cell".into()))?;
        book.create_sheet(Some("Late failure"), None)?
            .set_cell_value(500, 1, value)?;
        assert_rejected(&book, "rust_xlsxwriter backend");
    }
    for cached in [
        "",
        "123",
        "0.25",
        "NaN",
        "TRUE",
        "FALSE",
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NULL!",
        "#NUM!",
        "#REF!",
        "#VALUE!",
        "#GETTING_DATA",
    ] {
        let mut book = Workbook::new();
        book.active_mut().set_cell_value(
            1,
            1,
            CellValue::formula_with_cached("=1", CellValue::String(cached.into())),
        )?;
        assert_rejected(&book, "cache");
    }
    Ok(())
}

#[test]
fn quoted_function_like_tokens_are_rejected_without_touching_destinations() -> TestResult {
    for (name, formula) in [
        ("FILTER(", "='FILTER('!A1"),
        ("prefixFILTER(", "='prefixFILTER('!A1"),
        ("BETA.DIST(", "='BETA.DIST('!A1"),
        ("_xlfn.FILTER(", "='_xlfn.FILTER('!A1"),
        ("O'FILTER(", "='O''FILTER('!A1"),
        ("sum(", "='sum('!A1"),
        ("Text", "=\"FILTER(\""),
    ] {
        let mut book = Workbook::new();
        book.active_mut().set_cell_value(1, 1, CellValue::Int(1))?;
        book.create_sheet(Some(name), None)?
            .set_cell_value(1, 1, CellValue::Int(2))?;
        book.active_mut()
            .set_cell_value(1000, 2, CellValue::formula(formula))?;
        assert_rejected(&book, "quoted formula tokens");
    }
    let mut book = Workbook::new();
    book.create_sheet(Some("Sales Q1"), None)?
        .set_cell_value(1, 1, CellValue::Int(2))?;
    let value = CellValue::formula_with_cached("=SUM('Sales Q1'!A1,1)", CellValue::Int(3));
    book.active_mut().set_cell_value(2, 1, value.clone())?;
    let output = bytes(&book);
    assert_eq!(read(&output, false).active().get_cell_value(2, 1), value);
    let xml = part(&output, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("<f>SUM('Sales Q1'!A1,1)</f>"));
    assert!(!xml.contains("t=\"array\""));
    Ok(())
}

#[test]
fn attribute_whitespace_is_rejected_instead_of_normalized_by_xml() -> TestResult {
    for (raw, expected) in [
        ("A\rB", "A B"),
        ("A\tB", "A B"),
        ("A\r\nB", "A B"),
        ("A&#13;B", "A\rB"),
        ("A&#9;B", "A\tB"),
    ] {
        assert_eq!(
            attributes(&format!("<name val=\"{raw}\"/>"), b"name")[0]["val"],
            expected
        );
    }
    // Demonstrate the upstream defect through an independent XML-normalized
    // expectation, not through our reader's raw quick-xml attribute decoding.
    let raw_file = NamedTempFile::new()?;
    let mut raw_book = rust_xlsxwriter::Workbook::new();
    let raw_format = rust_xlsxwriter::Format::new()
        .set_font_name("Font\tName")
        .set_num_format("0.00\rX");
    raw_book
        .add_worksheet()
        .write_number_with_format(0, 0, 1.0, &raw_format)?;
    raw_book.save(raw_file.path())?;
    let raw_styles = part(&std::fs::read(raw_file.path())?, "xl/styles.xml");
    assert!(attributes(&raw_styles, b"name")
        .iter()
        .any(|a| a.get("val").is_some_and(|v| v == "Font Name")));
    assert_eq!(
        attributes(&raw_styles, b"numFmt")[0]["formatCode"],
        "0.00 X"
    );

    for whitespace in ["\r", "\n", "\t", "&#13;", "&#10;", "&#9;"] {
        let mut book = Workbook::new();
        let cell = book.active_mut().set_cell_value(1, 1, CellValue::Int(1))?;
        cell.rich_style = Some(Arc::new(RichStyle {
            font: Some(format!(
                "<font><name val=\"Font{whitespace}Name\"/><sz val=\"11\"/></font>"
            )),
            ..RichStyle::default()
        }));
        assert_rejected(&book, "control whitespace");
    }
    for whitespace in ['\r', '\n', '\t'] {
        for value in [CellValue::Int(1), CellValue::Date("2026-11-02".into())] {
            let mut book = Workbook::new();
            let cell = book.active_mut().set_cell_value(1, 1, value)?;
            // Validate the supplied code even when the effective temporal
            // format would otherwise replace this non-temporal code.
            cell.number_format = Some(format!("0.00{whitespace}X"));
            assert_rejected(&book, "control whitespace");
        }
    }
    let mut book = Workbook::new();
    let cell = book.active_mut().set_cell_value(1, 1, CellValue::Int(1))?;
    cell.rich_style = Some(Arc::new(RichStyle {
        font: Some("<font><name val=\"A &amp; B\"/><sz val=\"11\"/></font>".into()),
        ..RichStyle::default()
    }));
    cell.number_format = Some("0.00 \"USD & <tax>\"".into());
    let styles = part(&bytes(&book), "xl/styles.xml");
    assert!(attributes(&styles, b"name")
        .iter()
        .any(|a| a.get("val").is_some_and(|v| v == "A & B")));
    assert_eq!(
        attributes(&styles, b"numFmt")[0]["formatCode"],
        "0.00 \"USD & <tax>\""
    );
    Ok(())
}

#[test]
fn explicit_new_styles_and_blank_defaults_are_not_loaded_provenance() -> TestResult {
    let mut book = Workbook::new();
    let cell = book.active_mut().set_cell_value(1, 1, CellValue::Int(7))?;
    cell.rich_style = Some(Arc::new(common_style()));
    cell.number_format = Some("0.00".into());
    cell.source_xf_explicit = true;
    let blank_default = book.active_mut().set_cell_value(2, 1, CellValue::Empty)?;
    blank_default.source_xf_explicit = true;
    let styled_blank = book.active_mut().set_cell_value(3, 1, CellValue::Empty)?;
    styled_blank.rich_style = Some(Arc::new(common_style()));
    styled_blank.source_xf_explicit = true;
    let original = book.clone();
    let output = bytes(&book);
    let loaded = read(&output, false);
    assert_eq!(loaded.active().get_cell_value(1, 1), CellValue::Int(7));
    assert_eq!(loaded.active().get_cell_value(2, 1), CellValue::Empty);
    assert!(loaded.active().get_cell(3, 1).unwrap().rich_style.is_some());
    assert_eq!(book, original);
    // The flag is benign only without actual workbook/cell provenance.
    book.active_mut().get_cell_mut(1, 1).unwrap().source_xf_id = NonZeroU32::new(1);
    assert_rejected(&book, "loaded XF provenance");
    book.active_mut().get_cell_mut(1, 1).unwrap().source_xf_id = None;
    book.source_styles_xml = Some("<styleSheet/>".into());
    assert_rejected(&book, "provenance");
    Ok(())
}

#[test]
fn workbook_metadata_and_provenance_guards_are_atomic() -> TestResult {
    let mutations: &[fn(&mut Workbook)] = &[
        |b| b.read_only = true,
        |b| b.data_only = true,
        |b| b.date_1904 = true,
        |b| b.source_styles_xml = Some("<styleSheet/>".into()),
        |b| b.set_source_styles(vec![Arc::new(RichStyle::default())]),
        |b| b.active_mut().unsupported_cell_content = true,
        |b| b.active_mut().merged_ranges.push((1, 1, 2, 2)),
        |b| b.active_mut().title = "History".into(),
        |b| b.active_mut().get_cell_mut(1, 1).unwrap().comment = Some("comment".into()),
        |b| {
            b.active_mut().get_cell_mut(1, 1).unwrap().hyperlink =
                Some("https://example.com".into())
        },
        |b| b.active_mut().get_cell_mut(1, 1).unwrap().source_xf_id = NonZeroU32::new(1),
        |b| b.active_mut().get_cell_mut(1, 1).unwrap().row = 2,
        |b| b.active_mut().get_cell_mut(1, 1).unwrap().column = 0,
        |b| b.active_mut().get_cell_mut(1, 1).unwrap().number_format = Some(String::new()),
    ];
    for mutate in mutations {
        let mut book = Workbook::new();
        book.active_mut().set_cell_value(1, 1, CellValue::Int(1))?;
        mutate(&mut book);
        assert_rejected(&book, "rust_xlsxwriter backend");
    }
    assert_rejected(&Workbook::with_options(false, false), "active worksheet");
    let mut duplicate = Workbook::new();
    duplicate.create_sheet(Some("Second"), None)?.title = "sheet".into();
    assert_rejected(&duplicate, "duplicate");
    // Exercise sheet-template provenance independently of workbook fields.
    let loaded = read(&bytes(&Workbook::new()), false);
    let mut fresh = Workbook::new();
    fresh.active_mut().xml_template = loaded.active().xml_template.clone();
    assert!(fresh.active().xml_template.is_some());
    assert_rejected(&fresh, "source");
    Ok(())
}

#[test]
fn strict_style_schema_and_nonrepresentable_features_fail_before_output() -> TestResult {
    let mut bad_styles = Vec::new();
    for font in [
        "<font><name val=\"Arial\"/><sz val=\"11\"/><color rgb=\"00112233\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><color rgb=\"112233\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><color rgb=\"é123456\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><color rgb=\"FF112233\" tint=\"0.5\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><color theme=\"1\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><scheme val=\"minor\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><shadow/></font>",
        "<font><name val=\"Arial\"/><sz val=\"NaN\"/></font>",
        "<font><name val=\"Arial\"/><sz/></font>",
        "<font><b val=\"None\"/></font>",
        "<font><b/></font>",
        "<font><extLst/></font>",
        "<font unknown=\"1\"/>",
        "<font xmlns=\"urn:wrong\"/>",
        "<font xmlns=\"\"/>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><b/><b/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><b xmlns=\"urn:wrong\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><b val=\"false\" bogus=\"1\"/></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><color rgb=\"FF112233\"><x/></color></font>",
        "<font><name val=\"Arial\"/><sz val=\"11\"/><b val=\"None\"/></font>",
        "<font><b></font>",
        "<font/><font/>",
        "<!DOCTYPE font><font/>",
    ] {
        bad_styles.push(RichStyle {
            font: Some(font.into()),
            ..RichStyle::default()
        });
    }
    for fill in [
        "<fill><gradientFill/></fill>",
        "<fill><patternFill patternType=\"darkGrid\"/></fill>",
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"00112233\"/></patternFill></fill>",
        "<fill><patternFill><fgColor rgb=\"FF112233\"/></patternFill></fill>",
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF112233\"/><bgColor rgb=\"FF445566\"/></patternFill></fill>",
    ] { bad_styles.push(RichStyle { fill: Some(fill.into()), ..RichStyle::default() }); }
    for border in [
        "<border outline=\"0\"/>",
        "<border diagonalUp=\"true\"/>",
        "<border><start style=\"thin\"/></border>",
        "<border><left><color rgb=\"FF112233\"/></left></border>",
        "<border><left style=\"thin\"><color indexed=\"1\"/></left></border>",
    ] {
        bad_styles.push(RichStyle {
            border: Some(border.into()),
            ..RichStyle::default()
        });
    }
    for alignment in [
        "<alignment wrapText=\"None\"/>",
        "<alignment textRotation=\"200\"/>",
        "<alignment relativeIndent=\"1\"/>",
        "<alignment justifyLastLine=\"true\"/>",
        "<alignment shrinkToFit=\"true\" wrapText=\"true\"/>",
        "<alignment horizontal=\"fill\" shrinkToFit=\"true\"/>",
        "<alignment horizontal=\"center\" vertical=\"center\" indent=\"1\"/>",
        "<alignment readingOrder=\"3\"/>",
        "<alignment indent=\"0.5\"/>",
        "<alignment unknown=\"0\"/>",
        "<alignment><x/></alignment>",
    ] {
        bad_styles.push(RichStyle {
            alignment: Some(alignment.into()),
            ..RichStyle::default()
        });
    }
    bad_styles.push(RichStyle {
        protection: Some("<protection locked=\"None\"/>".into()),
        ..RichStyle::default()
    });
    bad_styles.push(RichStyle {
        protection: Some("<protection unknown=\"0\"/>".into()),
        ..RichStyle::default()
    });
    for style in bad_styles {
        let mut book = Workbook::new();
        book.active_mut()
            .set_cell_value(1, 1, CellValue::String("valid".into()))?;
        book.create_sheet(Some("Bad style"), None)?
            .set_cell_value(100, 2, CellValue::Int(1))?;
        book.sheets_mut()[1]
            .get_cell_mut(100, 2)
            .unwrap()
            .rich_style = Some(Arc::new(style));
        assert_rejected(&book, "");
    }
    Ok(())
}
