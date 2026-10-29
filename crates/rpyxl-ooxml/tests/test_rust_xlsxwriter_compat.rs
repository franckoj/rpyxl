use rpyxl_core::cell::CellValue;
use rpyxl_ooxml::reader::XlsxPackage;
use rust_xlsxwriter::Workbook as RustXlsxWorkbook;
use tempfile::NamedTempFile;

#[test]
fn test_rust_xlsxwriter_compatibility() -> Result<(), Box<dyn std::error::Error>> {
    // 1. Generate an XLSX workbook using rust_xlsxwriter
    let temp_file = NamedTempFile::new()?;
    let path = temp_file.path();

    let mut workbook = RustXlsxWorkbook::new();
    let worksheet = workbook.add_worksheet().set_name("RustSheet")?;

    worksheet.write_string(0, 0, "RustXlsxWriter Cell")?; // A1
    worksheet.write_number(0, 1, 100.5)?; // B1
    worksheet.write_boolean(0, 2, true)?; // C1
    worksheet.write_number(1, 0, 42)?; // A2
    worksheet.write_string(1, 1, "RPyXL Compatibility")?; // B2

    workbook.save(path)?;

    // 2. Open and parse with RPyXL OOXML engine
    let mut package = XlsxPackage::open(path, false, false)?;
    assert_eq!(package.sheets.len(), 1);
    assert_eq!(package.sheets[0].name, "RustSheet");

    let wb = package.to_workbook()?;
    let ws = wb.get_sheet("RustSheet").expect("Sheet should exist");

    // 3. Verify cell values in RPyXL core
    let cell_a1 = ws.get_cell(1, 1).expect("A1 should exist");
    assert_eq!(
        cell_a1.value,
        CellValue::String("RustXlsxWriter Cell".to_string())
    );

    let cell_b1 = ws.get_cell(1, 2).expect("B1 should exist");
    if let CellValue::Float(n) = cell_b1.value {
        assert!((n - 100.5).abs() < 1e-6);
    } else {
        panic!("Expected Float in B1, got {:?}", cell_b1.value);
    }

    let cell_c1 = ws.get_cell(1, 3).expect("C1 should exist");
    assert_eq!(cell_c1.value, CellValue::Bool(true));

    let cell_a2 = ws.get_cell(2, 1).expect("A2 should exist");
    match cell_a2.value {
        CellValue::Int(42) => (),
        CellValue::Float(f) if (f - 42.0).abs() < 1e-6 => (),
        ref other => panic!("Expected 42 in A2, got {:?}", other),
    }

    let cell_b2 = ws.get_cell(2, 2).expect("B2 should exist");
    assert_eq!(
        cell_b2.value,
        CellValue::String("RPyXL Compatibility".to_string())
    );

    Ok(())
}
