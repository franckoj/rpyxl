//! Probe only: Calamine is a pinned dev dependency, not a production backend.
//! Known representation/fidelity gaps are asserted explicitly, never skipped or
//! normalized into a passing cross-backend equivalence result. Existing native
//! fidelity/security tests remain authoritative and unchanged.

#[allow(dead_code)]
#[path = "../examples/bench_backend_read.rs"]
mod bench;

use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use calamine::{CellErrorType, Data, Reader, Xlsx};
use rpyxl_core::cell::CellValue;
use rpyxl_ooxml::XlsxPackage;
use rust_xlsxwriter::{ExcelDateTime, Format, Formula, Workbook};
use tempfile::TempDir;
use zip::write::SimpleFileOptions;

use bench::{Backend, Result};

type Values = BTreeMap<(u32, u32), Data>;

fn save(workbook: &mut Workbook) -> Result<(TempDir, std::path::PathBuf)> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("probe.xlsx");
    workbook.save(&path)?;
    Ok((dir, path))
}

fn calamine_values(path: &Path, name: &str) -> Result<Values> {
    let mut workbook = Xlsx::new(bench::validated_reader(path)?)?;
    let mut reader = workbook.worksheet_cells_reader(name)?;
    let mut values = Values::new();
    while let Some(cell) = reader.next_cell()? {
        let (row, col) = cell.get_position();
        values.insert((row + 1, col + 1), Data::from(cell.get_value().clone()));
    }
    Ok(values)
}

fn assert_native_number(value: &CellValue, expected: f64) {
    match value {
        CellValue::Int(value) => assert_eq!(*value as f64, expected),
        CellValue::Float(value) => assert_eq!(*value, expected),
        other => panic!("expected native number {expected}, got {other:?}"),
    }
}

#[test]
fn scalar_cached_values_and_all_sheets_match_without_rpyxl_conversion() -> Result<()> {
    let mut writer = Workbook::new();
    let sheet = writer.add_worksheet().set_name("Values")?;
    let strings = [
        "héllo 日本語 😀 <&>",
        "  padded  ",
        "   ",
        "\t padded\t",
        "a\nb",
    ];
    for (row, text) in strings.iter().enumerate() {
        sheet.write_string(row as u32, 0, *text)?;
    }
    sheet.write_number(5, 0, 42)?;
    sheet.write_number(6, 0, -12.5)?;
    sheet.write_boolean(7, 0, true)?;
    sheet.write_boolean(8, 0, false)?;
    // Intentionally stale: neither backend may calculate 1+2 to replace 999.
    sheet.write_formula(9, 0, Formula::new("=1+2").set_result("999"))?;
    sheet.write_formula(10, 0, Formula::new("=\" cached \"").set_result(" cached "))?;
    sheet.write_formula(11, 0, Formula::new("=TRUE()").set_result("TRUE"))?;
    sheet.write_number(14, 2, 7)?;
    writer.add_worksheet().set_name("Empty")?;
    writer
        .add_worksheet()
        .set_name("Hidden")?
        .set_hidden(true)
        .write_string(0, 0, "still visited")?;
    let (_dir, path) = save(&mut writer)?;

    let native = bench::run(Backend::Native, &path)?;
    let calamine = bench::run(Backend::Calamine, &path)?;
    assert_eq!(native.sheets.len(), 3);
    assert_eq!(calamine.sheets.len(), 3);
    for (native, calamine) in native.sheets.iter().zip(&calamine.sheets) {
        assert_eq!(native.name, calamine.name);
        assert_eq!(native.nonempty, calamine.nonempty);
        assert_eq!(native.bounds, calamine.bounds);
        assert_eq!(native.excluded_cells, 0);
        assert_eq!(calamine.excluded_cells, 0);
        assert_eq!(native.checksum_cells, calamine.checksum_cells);
        assert_eq!(native.checksum, calamine.checksum);
    }
    assert_eq!(native.sheets[0].nonempty, 13);
    assert_eq!(native.sheets[0].bounds, Some([1, 1, 15, 3]));
    assert_eq!(native.sheets[1].bounds, None);
    assert_eq!(native.sheets[2].nonempty, 1);
    assert!(native.overall >= native.prevalidate + native.backend_total);
    assert!(calamine.overall >= calamine.prevalidate + calamine.backend_total);

    let mut package = XlsxPackage::from_reader(bench::validated_reader(&path)?, true, true)?;
    let workbook = package.to_workbook()?;
    let sheet = workbook.get_sheet("Values").unwrap();
    let values = calamine_values(&path, "Values")?;
    for (row, text) in strings.iter().enumerate() {
        let pos = (row as u32 + 1, 1);
        assert_eq!(
            sheet.get_cell(pos.0, pos.1).unwrap().value,
            CellValue::String((*text).into())
        );
        assert_eq!(values[&pos], Data::String((*text).into()));
    }
    for (row, expected) in [(6, 42.0), (7, -12.5), (10, 999.0), (15, 7.0)] {
        let col = if row == 15 { 3 } else { 1 };
        assert_native_number(&sheet.get_cell(row, col).unwrap().value, expected);
        assert_eq!(values[&(row, col)], Data::Float(expected));
    }
    for (row, expected) in [(8, true), (9, false), (12, true)] {
        assert_eq!(
            sheet.get_cell(row, 1).unwrap().value,
            CellValue::Bool(expected)
        );
        assert_eq!(values[&(row, 1)], Data::Bool(expected));
    }
    assert_eq!(
        sheet.get_cell(11, 1).unwrap().value,
        CellValue::String(" cached ".into())
    );
    assert_eq!(values[&(11, 1)], Data::String(" cached ".into()));

    // Prove streaming cached values agree with Calamine's materialized API;
    // the benchmark itself never constructs this Range.
    let mut calamine_book = Xlsx::new(bench::validated_reader(&path)?)?;
    let range = calamine_book.worksheet_range("Values")?;
    for (&(row, col), value) in &values {
        assert_eq!(range.get_value((row - 1, col - 1)), Some(value));
    }
    let mut formulas = XlsxPackage::open(&path, true, false)?;
    let formulas = formulas.to_workbook()?;
    match &formulas
        .get_sheet("Values")
        .unwrap()
        .get_cell(10, 1)
        .unwrap()
        .value
    {
        CellValue::Formula {
            formula,
            cached_value,
        } => {
            assert_eq!(formula, "=1+2");
            assert_native_number(cached_value.as_deref().unwrap(), 999.0);
        }
        other => panic!("non-data-only control must retain formula, got {other:?}"),
    }
    let json = native.json(&path);
    assert!(json.contains("\"mismatches\":null"));
    assert!(json.contains("\"xml_security_equivalent\":false"));
    assert!(json.contains("cached_values_only_no_formula_evaluation"));
    Ok(())
}

#[test]
fn crlf_excel_escape_gap_is_visible_not_normalized_away() -> Result<()> {
    let mut writer = Workbook::new();
    writer
        .add_worksheet()
        .set_name("CRLF")?
        .write_string(0, 0, "line1\r\nline2")?;
    let (_dir, path) = save(&mut writer)?;
    let mut package = XlsxPackage::open(&path, true, true)?;
    let workbook = package.to_workbook()?;
    // RustXlsxWriter emits _x000D_ for CR. Current native shared-string decoding
    // does XML decoding but not Excel's _xHHHH_ decoding. Do not bless this as
    // equivalent CRLF fidelity or change the existing native XML CRLF tests.
    assert_eq!(
        workbook
            .get_sheet("CRLF")
            .unwrap()
            .get_cell(1, 1)
            .unwrap()
            .value,
        CellValue::String("line1_x000D_\nline2".into())
    );
    assert_eq!(
        calamine_values(&path, "CRLF")?[&(1, 1)],
        Data::String("line1\r\nline2".into())
    );
    let native = bench::run(Backend::Native, &path)?;
    let calamine = bench::run(Backend::Calamine, &path)?;
    assert_ne!(native.sheets[0].checksum, calamine.sheets[0].checksum);
    Ok(())
}

#[test]
fn dates_durations_and_error_caches_are_explicit_unsupported_type_gaps() -> Result<()> {
    let mut writer = Workbook::new();
    let sheet = writer.add_worksheet().set_name("Gaps")?;
    let date = ExcelDateTime::from_ymd(2024, 2, 29)?.and_hms(12, 30, 0)?;
    let date_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    sheet.write_datetime_with_format(0, 0, &date, &date_format)?;
    sheet.write_number_with_format(1, 0, 1.5, &Format::new().set_num_format("[h]:mm:ss"))?;
    sheet.write_formula(2, 0, Formula::new("=1/0").set_result("#DIV/0!"))?;
    let (_dir, path) = save(&mut writer)?;
    let mut package = XlsxPackage::open(&path, true, true)?;
    let workbook = package.to_workbook()?;
    let sheet = workbook.get_sheet("Gaps").unwrap();
    let values = calamine_values(&path, "Gaps")?;
    assert_eq!(
        sheet.get_cell(1, 1).unwrap().value,
        CellValue::DateTime("2024-02-29 12:30:00".into())
    );
    match &values[&(1, 1)] {
        Data::DateTime(value) => {
            assert!(value.is_datetime());
            assert_eq!(value.to_ymd_hms_milli(), (2024, 2, 29, 12, 30, 0, 0));
        }
        other => panic!("expected Calamine date, got {other:?}"),
    }
    assert_native_number(&sheet.get_cell(2, 1).unwrap().value, 1.5);
    match &values[&(2, 1)] {
        Data::DateTime(value) => {
            assert!(value.is_duration());
            assert_eq!(value.as_f64(), 1.5);
        }
        other => panic!("expected Calamine duration, got {other:?}"),
    }
    // Native has no error CellValue variant; never claim string/error equality.
    assert_eq!(
        sheet.get_cell(3, 1).unwrap().value,
        CellValue::String("#DIV/0!".into())
    );
    assert_eq!(values[&(3, 1)], Data::Error(CellErrorType::Div0));
    let native = bench::run(Backend::Native, &path)?;
    let calamine = bench::run(Backend::Calamine, &path)?;
    assert_eq!(native.sheets[0].excluded_cells, 1);
    assert_eq!(calamine.sheets[0].excluded_cells, 3);
    assert_eq!(calamine.sheets[0].types["error"], 1);
    assert_eq!(calamine.sheets[0].types["duration"], 1);
    Ok(())
}

/// Modify a deterministic writer fixture, preserving all other ZIP parts.
fn replace_part(path: &Path, part: &str, transform: impl FnOnce(String) -> String) -> Result<()> {
    let mut source = zip::ZipArchive::new(Cursor::new(std::fs::read(path)?))?;
    let mut target = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut transform = Some(transform);
    for index in 0..source.len() {
        let mut entry = source.by_index(index)?;
        let name = entry.name().to_owned();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        if name == part {
            bytes = transform.take().unwrap()(String::from_utf8(bytes)?).into_bytes();
        }
        target.start_file(name, SimpleFileOptions::default())?;
        target.write_all(&bytes)?;
    }
    assert!(transform.is_none(), "fixture part must exist");
    std::fs::write(path, target.finish()?.into_inner())?;
    Ok(())
}

#[test]
fn absent_formula_cache_is_empty_not_evaluated_or_defaulted_to_zero() -> Result<()> {
    let mut writer = Workbook::new();
    writer.add_worksheet().set_name("Missing")?.write_formula(
        0,
        0,
        Formula::new("=1+2").set_result("999"),
    )?;
    let (_dir, path) = save(&mut writer)?;
    replace_part(&path, "xl/worksheets/sheet1.xml", |xml| {
        assert!(xml.contains("<v>999</v>"));
        xml.replace("<v>999</v>", "")
    })?;
    let mut package = XlsxPackage::open(&path, true, true)?;
    let workbook = package.to_workbook()?;
    assert_eq!(
        workbook
            .get_sheet("Missing")
            .unwrap()
            .get_cell(1, 1)
            .unwrap()
            .value,
        CellValue::Empty
    );
    let values = calamine_values(&path, "Missing")?;
    // next_cell may omit empty cells entirely. Both storage forms mean no value.
    assert!(values
        .get(&(1, 1))
        .is_none_or(|value| *value == Data::Empty));
    for backend in [Backend::Native, Backend::Calamine] {
        let report = bench::run(backend, &path)?;
        assert_eq!(report.sheets[0].nonempty, 0);
        assert_eq!(report.sheets[0].checksum_cells, 0);
    }
    Ok(())
}

#[test]
fn both_backends_reject_the_same_native_zip_entry_limit() -> Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("too_many_entries.xlsx");
    let mut archive = zip::ZipWriter::new(std::fs::File::create(&path)?);
    for index in 0..=rpyxl_ooxml::reader::MAX_ZIP_ENTRIES {
        archive.start_file(format!("entry{index}"), SimpleFileOptions::default())?;
    }
    archive.finish()?;
    for backend in [Backend::Native, Backend::Calamine] {
        let error = bench::run(backend, &path).unwrap_err().to_string();
        assert!(error.contains("Archive contains"), "{backend:?}: {error}");
        assert!(error.contains("10000"), "{backend:?}: {error}");
    }
    Ok(())
}

#[test]
fn native_shared_strings_dtd_rejection_is_not_bypassed_by_the_benchmark() -> Result<()> {
    let mut writer = Workbook::new();
    writer.add_worksheet().write_string(0, 0, "probe")?;
    let (_dir, path) = save(&mut writer)?;
    replace_part(&path, "xl/sharedStrings.xml", |xml| {
        let end = xml.find("?>").expect("writer XML declaration") + 2;
        format!(
            "{}<!DOCTYPE sst [<!ENTITY probe 'unsafe'>]>{}",
            &xml[..end],
            &xml[end..]
        )
    })?;
    let error = bench::run(Backend::Native, &path).unwrap_err().to_string();
    assert!(
        error.contains("DTD") || error.contains("DOCTYPE"),
        "{error}"
    );
    // Deliberately no assertion of Calamine XML-security equivalence.
    Ok(())
}

#[test]
fn checksum_is_order_independent_but_value_coordinate_and_type_sensitive() {
    let mut a = bench::SheetSummary::default();
    a.native(1, 1, &CellValue::Int(42));
    a.native(2, 1, &CellValue::String("  x\r\n".into()));
    let mut b = bench::SheetSummary::default();
    b.calamine(2, 1, &calamine::DataRef::SharedString("  x\r\n"));
    b.calamine(1, 1, &calamine::DataRef::Float(42.0));
    assert_eq!(a.checksum, b.checksum);
    assert_eq!(a.checksum_cells, b.checksum_cells);
    b.calamine(3, 1, &calamine::DataRef::Empty);
    assert_eq!(a.checksum, b.checksum);
    for (row, value) in [
        (1, CellValue::String("42".into())),
        (2, CellValue::Int(42)),
        (1, CellValue::Int(43)),
    ] {
        let mut changed = bench::SheetSummary::default();
        changed.native(row, 1, &value);
        let mut original = bench::SheetSummary::default();
        original.native(1, 1, &CellValue::Int(42));
        assert_ne!(original.checksum, changed.checksum);
    }
    let mut exact = bench::SheetSummary::default();
    exact.native(1, 1, &CellValue::Int(9_007_199_254_740_993));
    let mut rounded = bench::SheetSummary::default();
    rounded.calamine(1, 1, &calamine::DataRef::Float(9_007_199_254_740_992.0));
    assert_ne!(exact.checksum, rounded.checksum);
    assert_eq!(
        bench::json_string("hé\"\\\t\r\n\u{0001}"),
        "\"hé\\\"\\\\\\t\\r\\n\\u0001\""
    );
}
