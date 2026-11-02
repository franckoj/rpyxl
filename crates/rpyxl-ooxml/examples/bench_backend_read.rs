//! Isolated, single-thread, cached-value experiment; no production routing.
//! Run each backend in a fresh release process against the SAME existing XLSX:
//! cargo run --release -p rpyxl-ooxml --example bench_backend_read -- native input.xlsx
//! cargo run --release -p rpyxl-ooxml --example bench_backend_read -- calamine input.xlsx
//! Input generation is deliberately not part of this executable.
//!
//! Native materializes the current full Workbook (including styles/templates);
//! Calamine 0.38 streams worksheet cells, retaining workbook metadata/shared strings.
//! Neither arm builds a Calamine Range plus an RPyXL copy. This is NOT an equal-
//! storage parser microbenchmark. Measure peak RSS externally in separate processes;
//! no portable RSS measurement or security equivalence is claimed here.
//!
//! Required CI gates (local Windows may lack the MSVC C++ linker):
//! cargo check -p rpyxl-ooxml --example bench_backend_read --test test_calamine_probe
//! cargo test -p rpyxl-ooxml --test test_calamine_probe

use std::collections::BTreeMap;
use std::error::Error;
use std::fs::File;
use std::io::{BufReader, Seek, SeekFrom};
use std::path::Path;
use std::time::{Duration, Instant};

use calamine::{DataRef, Reader, SheetType, Xlsx};
use rpyxl_core::cell::CellValue;
use rpyxl_ooxml::{validate_zip_archive_security, XlsxPackage};

pub(crate) type Result<T> = std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Backend {
    Native,
    Calamine,
}

impl Backend {
    fn name(self) -> &'static str {
        match self {
            Self::Native => "native",
            Self::Calamine => "calamine",
        }
    }

    fn storage(self) -> &'static str {
        match self {
            Self::Native => "full_native_to_workbook_then_borrowed_cell_iteration",
            Self::Calamine => "xlsx_cell_reader_stream_no_range_no_rpyxl_copy",
        }
    }
}

#[derive(Debug, Default, PartialEq, Eq)]
pub(crate) struct SheetSummary {
    pub(crate) name: String,
    pub(crate) cells_visited: u64,
    pub(crate) nonempty: u64,
    /// One-based nonempty [min_row, min_col, max_row, max_col], not declared dimensions.
    pub(crate) bounds: Option<[u32; 4]>,
    pub(crate) types: BTreeMap<&'static str, u64>,
    pub(crate) checksum_cells: u64,
    pub(crate) excluded_cells: u64,
    pub(crate) checksum: u64,
}

impl SheetSummary {
    fn new(name: &str) -> Self {
        Self {
            name: name.to_owned(),
            ..Self::default()
        }
    }

    fn visit(&mut self, row: u32, col: u32, kind: &'static str, nonempty: bool) {
        self.cells_visited += 1;
        *self.types.entry(kind).or_default() += 1;
        if nonempty {
            self.nonempty += 1;
            let bounds = self.bounds.get_or_insert([row, col, row, col]);
            bounds[0] = bounds[0].min(row);
            bounds[1] = bounds[1].min(col);
            bounds[2] = bounds[2].max(row);
            bounds[3] = bounds[3].max(col);
        }
    }

    // FNV-1a per framed cell, wrapping sum across cells: deterministic regardless
    // of native BTreeMap order versus worksheet XML order. Not cryptographic.
    fn hash(&mut self, row: u32, col: u32, tag: u8, bytes: &[u8]) {
        let mut hash = 0xcbf29ce484222325_u64;
        for byte in row
            .to_le_bytes()
            .iter()
            .chain(col.to_le_bytes().iter())
            .chain(std::iter::once(&tag))
            .chain(bytes.iter())
        {
            hash = (hash ^ u64::from(*byte)).wrapping_mul(0x100000001b3);
        }
        self.checksum = self.checksum.wrapping_add(hash);
        self.checksum_cells += 1;
    }

    fn integer(&mut self, row: u32, col: u32, value: i64) {
        self.hash(row, col, b'i', &value.to_le_bytes());
    }

    fn float(&mut self, row: u32, col: u32, value: f64) {
        if !value.is_finite() {
            self.excluded_cells += 1;
        } else if value.fract() == 0.0
            && (-9_223_372_036_854_775_808.0..9_223_372_036_854_775_808.0).contains(&value)
        {
            // Equal exactly representable integral floats/ints, including -0.
            // Do not cast all integers to f64: that would hide precision loss.
            self.integer(row, col, value as i64);
        } else {
            self.hash(row, col, b'f', &value.to_bits().to_le_bytes());
        }
    }

    pub(crate) fn native(&mut self, row: u32, col: u32, value: &CellValue) {
        let kind = match value {
            CellValue::Empty => "empty",
            CellValue::String(_) => "string",
            CellValue::Int(_) => "int",
            CellValue::Float(_) => "float",
            CellValue::Bool(_) => "bool",
            CellValue::Date(_) => "date",
            CellValue::DateTime(_) => "datetime",
            CellValue::Formula { .. } => "unexpected_formula",
        };
        self.visit(row, col, kind, !value.is_empty());
        match value {
            CellValue::Empty => {}
            CellValue::String(text) => self.hash(row, col, b's', text.as_bytes()),
            CellValue::Int(value) => self.integer(row, col, *value),
            CellValue::Float(value) => self.float(row, col, *value),
            CellValue::Bool(value) => self.hash(row, col, b'b', &[u8::from(*value)]),
            _ => self.excluded_cells += 1,
        }
    }

    pub(crate) fn calamine(&mut self, row: u32, col: u32, value: &DataRef<'_>) {
        let kind = match value {
            DataRef::Empty => "empty",
            DataRef::String(_) | DataRef::SharedString(_) => "string",
            DataRef::Int(_) => "int",
            DataRef::Float(_) => "float",
            DataRef::Bool(_) => "bool",
            DataRef::DateTime(value) if value.is_duration() => "duration",
            DataRef::DateTime(_) => "datetime",
            DataRef::DateTimeIso(_) => "datetime_iso",
            DataRef::DurationIso(_) => "duration_iso",
            DataRef::Error(_) => "error",
        };
        self.visit(row, col, kind, !matches!(value, DataRef::Empty));
        match value {
            DataRef::Empty => {}
            DataRef::String(text) => self.hash(row, col, b's', text.as_bytes()),
            DataRef::SharedString(text) => self.hash(row, col, b's', text.as_bytes()),
            DataRef::Int(value) => self.integer(row, col, *value),
            DataRef::Float(value) => self.float(row, col, *value),
            DataRef::Bool(value) => self.hash(row, col, b'b', &[u8::from(*value)]),
            _ => self.excluded_cells += 1,
        }
    }

    fn json(&self) -> String {
        let types = self
            .types
            .iter()
            .map(|(kind, count)| format!("{}:{count}", json_string(kind)))
            .collect::<Vec<_>>()
            .join(",");
        let bounds = match self.bounds {
            Some([a, b, c, d]) => format!("[{a},{b},{c},{d}]"),
            None => "null".to_owned(),
        };
        format!(
            "{{\"name\":{},\"cells_visited\":{},\"nonempty\":{},\"nonempty_bounds_1based\":{bounds},\"types\":{{{types}}},\"checksum_cells\":{},\"excluded_cells\":{},\"scalar_checksum\":\"{:016x}\"}}",
            json_string(&self.name), self.cells_visited, self.nonempty,
            self.checksum_cells, self.excluded_cells, self.checksum
        )
    }
}

#[derive(Debug)]
pub(crate) struct Report {
    pub(crate) backend: Backend,
    pub(crate) sheets: Vec<SheetSummary>,
    pub(crate) prevalidate: Duration,
    pub(crate) open_load: Duration,
    pub(crate) consume: Duration,
    pub(crate) backend_total: Duration,
    pub(crate) overall: Duration,
}

/// Same zip 2.x limits, same open file, for BOTH backends before their timer.
/// Native also revalidates internally; that extra work stays in its open/load time.
/// This does not add DTD/entity checks to Calamine or alter native XML safeguards.
pub(crate) fn validated_reader(path: &Path) -> Result<BufReader<File>> {
    let mut archive = zip::ZipArchive::new(BufReader::new(File::open(path)?))?;
    validate_zip_archive_security(&mut archive)?;
    let mut reader = archive.into_inner();
    reader.seek(SeekFrom::Start(0))?;
    Ok(reader)
}

pub(crate) fn run(backend: Backend, path: &Path) -> Result<Report> {
    let overall_start = Instant::now();
    let reader = validated_reader(path)?;
    let prevalidate = overall_start.elapsed();
    let start = Instant::now();
    let mut sheets = Vec::new();
    let (open_load, consume) = match backend {
        Backend::Native => {
            let mut package = XlsxPackage::from_reader(reader, true, true)?;
            let workbook = package.to_workbook()?;
            let open_load = start.elapsed();
            let consume_start = Instant::now();
            for sheet in workbook.sheets() {
                let mut summary = SheetSummary::new(&sheet.title);
                for (&(row, col), cell) in sheet.cells() {
                    summary.native(row, col, &cell.value);
                }
                sheets.push(summary);
            }
            (open_load, consume_start.elapsed())
        }
        Backend::Calamine => {
            let mut workbook = Xlsx::new(reader)?;
            // Ignore chartsheets, not hidden worksheets. Never use worksheets(),
            // which materializes ranges and can silently omit read errors.
            let names: Vec<_> = workbook
                .sheets_metadata()
                .iter()
                .filter(|sheet| sheet.typ == SheetType::WorkSheet)
                .map(|sheet| sheet.name.clone())
                .collect();
            let open_load = start.elapsed();
            let consume_start = Instant::now();
            for name in names {
                let mut summary = SheetSummary::new(&name);
                let mut cells = workbook.worksheet_cells_reader(&name)?;
                while let Some(cell) = cells.next_cell()? {
                    let (row, col) = cell.get_position();
                    summary.calamine(row + 1, col + 1, cell.get_value());
                }
                sheets.push(summary);
            }
            (open_load, consume_start.elapsed())
        }
    };
    let backend_total = start.elapsed();
    let overall = overall_start.elapsed();
    Ok(Report {
        backend,
        sheets,
        prevalidate,
        open_load,
        consume,
        backend_total,
        overall,
    })
}

pub(crate) fn json_string(text: &str) -> String {
    let mut out = String::from("\"");
    for ch in text.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            ch if ch <= '\u{1f}' => out.push_str(&format!("\\u{:04x}", u32::from(ch))),
            ch => out.push(ch),
        }
    }
    out.push('"');
    out
}

impl Report {
    pub(crate) fn json(&self, path: &Path) -> String {
        let sheets = self
            .sheets
            .iter()
            .map(SheetSummary::json)
            .collect::<Vec<_>>()
            .join(",");
        let cells: u64 = self.sheets.iter().map(|s| s.cells_visited).sum();
        let nonempty: u64 = self.sheets.iter().map(|s| s.nonempty).sum();
        format!(
            concat!(
                "{{\"schema\":\"backend_read_v1\",\"backend\":{},\"path\":{},",
                "\"mode\":\"cached_values_only_no_formula_evaluation\",\"threads\":1,",
                "\"storage\":{},\"rss_bytes\":null,\"rss_note\":\"measure_peak_externally_per_fresh_process_storage_differs\",",
                "\"zip_security\":\"same_native_prevalidation_both_native_rechecks_inside_timer\",",
                "\"xml_security_equivalent\":false,",
                "\"timing_ns\":{{\"prevalidate_open_zip\":{},\"backend_open_load\":{},\"consume\":{},\"backend_total_with_drop\":{},\"overall_total\":{}}},",
                "\"timing_note\":\"native_open_load_includes_all_sheets_calamine_consume_includes_sheet_inflate_parse_no_generation_or_output\",",
                "\"cells_visited\":{cells},\"nonempty\":{nonempty},\"sheets\":[{sheets}],",
                "\"comparison\":{{\"status\":\"not_run_single_backend\",\"mismatches\":null,",
                "\"mismatch_schema\":[\"sheet_order_name\",\"nonempty_bounds\",\"nonempty_count\",\"scalar_checksum_cells\",\"scalar_checksum\",\"type_gap\"],",
                "\"checksum\":\"wrapping_sum_fnv1a64_coordinate_tag_payload_per_sheet_nonempty_scalars\",",
                "\"scope\":\"text_bool_finite_numbers_only_integral_float_int_normalized_no_whitespace_normalization\",",
                "\"unsupported\":[\"dates_times_durations\",\"errors_native_exposes_as_strings\",\"nonfinite_numbers\",\"formula_text_or_recalculation\"],",
                                "\"known_text_gap\":\"native_does_not_decode_excel_xHHHH_escapes_including_CR_in_RustXlsxWriter_CRLF_checksum_mismatch_is_real\",",
                "\"rules\":\"compare_same_file_cached_mode_only_exclude_different_cache_modes_report_gaps_as_unsupported_not_pass_ignore_empty_storage_and_int_float_type_differences_checksums_are_not_a_fidelity_proof\"}}}}"
            ),
            json_string(self.backend.name()), json_string(&path.to_string_lossy()),
            json_string(self.backend.storage()), self.prevalidate.as_nanos(),
            self.open_load.as_nanos(), self.consume.as_nanos(),
            self.backend_total.as_nanos(), self.overall.as_nanos(),
            cells = cells, nonempty = nonempty, sheets = sheets,
        )
    }
}

fn main() -> Result<()> {
    let mut args = std::env::args_os().skip(1);
    let backend = match args.next().as_deref().and_then(|s| s.to_str()) {
        Some("native") => Backend::Native,
        Some("calamine") => Backend::Calamine,
        _ => {
            return Err(
                "usage: bench_backend_read <native|calamine> <input.xlsx> (cached values only)"
                    .into(),
            )
        }
    };
    let path = args.next().ok_or("missing XLSX path")?;
    if args.next().is_some() {
        return Err("unexpected argument; only cached-value mode is supported".into());
    }
    let path = Path::new(&path);
    let report = run(backend, path)?;
    println!("{}", report.json(path));
    Ok(())
}
