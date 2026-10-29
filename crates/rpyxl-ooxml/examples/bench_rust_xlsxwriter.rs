use rust_xlsxwriter::Workbook;
use std::env;
use std::fs;
use std::path::Path;
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let rows: u32 = args.get(1).and_then(|s| s.parse().ok()).unwrap_or(10_000);
    let cols: u16 = args.get(2).and_then(|s| s.parse().ok()).unwrap_or(10);
    let schema = args.get(3).map(|s| s.as_str()).unwrap_or("enterprise");

    let out_dir = Path::new("/tmp/rust_xlsxwriter_bench");
    fs::create_dir_all(out_dir)?;
    let out_path = out_dir.join(format!("bench_{}_{}_{}.xlsx", rows, cols, schema));

    let t0 = Instant::now();
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for r in 0..rows {
        if schema == "enterprise" {
            worksheet.write_string(r, 0, format!("User_{}", r))?;
            worksheet.write_number(r, 1, (r * 10) as f64)?;
            worksheet.write_number(r, 2, (r as f64) * 1.25)?;
            worksheet.write_boolean(r, 3, r % 2 == 0)?;
            worksheet.write_string(r, 4, "2023-01-01 12:00:00")?;
            worksheet.write_string(r, 5, format!("Item_{}", r % 100))?;
            worksheet.write_number(r, 6, (r * 5) as f64)?;
            worksheet.write_number(r, 7, ((r as f64) * 0.05 * 100.0).round() / 100.0)?;
            worksheet.write_boolean(r, 8, r % 3 == 0)?;
            worksheet.write_string(r, 9, "Active")?;
        } else {
            for c in 0..cols {
                worksheet.write_number(r, c, (r * 10 + c as u32) as f64)?;
            }
        }
    }

    workbook.save(&out_path)?;
    let duration = t0.elapsed();

    let total_cells = (rows as u64) * (cols as u64);
    let throughput = (total_cells as f64) / duration.as_secs_f64();

    println!(
        "{{\"rows\": {}, \"cols\": {}, \"cells\": {}, \"schema\": \"{}\", \"seconds\": {:.6}, \"throughput_cells_per_sec\": {:.0}}}",
        rows, cols, total_cells, schema, duration.as_secs_f64(), throughput
    );

    let _ = fs::remove_file(&out_path);
    Ok(())
}
