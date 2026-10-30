use std::collections::HashMap;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;

use parking_lot::RwLock;
use pyo3::exceptions::{PyAttributeError, PyIndexError, PyKeyError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDict, PyFloat, PyInt, PyList, PySlice, PyString, PyTuple};

use rpyxl_core::cell::{CellValue, RichStyle};
use rpyxl_core::coordinate::{
    coordinate_to_tuple, get_column_letter, range_boundaries, tuple_to_coordinate, MAX_COLUMN,
    MAX_ROW,
};
use rpyxl_core::error::{RpyxlError, RpyxlResult};
use rpyxl_core::workbook::Workbook;
use rpyxl_ooxml::{
    read_single_cell, save_workbook, save_write_only_workbook, write_workbook,
    write_write_only_workbook, SharedStringsTable, SheetInfo, StreamingRowReader, Styles,
    WriteOnlySheet, XlsxPackage,
};

type RichStyleInput = (
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
);

pub fn map_rpyxl_err(err: RpyxlError) -> PyErr {
    match err {
        RpyxlError::SheetNotFound(name) => {
            PyKeyError::new_err(format!("Worksheet {} does not exist.", name))
        }
        RpyxlError::DuplicateSheetTitle(name) => {
            PyValueError::new_err(format!("Worksheet {} already exists.", name))
        }
        RpyxlError::InvalidSheetTitle(msg) => PyValueError::new_err(msg),
        RpyxlError::SheetIndexOutOfBounds(idx, count) => PyIndexError::new_err(format!(
            "Worksheet index {} out of range (sheet count {})",
            idx, count
        )),
        RpyxlError::CannotRemoveLastSheet => {
            PyValueError::new_err("Cannot remove the only worksheet from workbook")
        }
        RpyxlError::InvalidCoordinate(coord) => {
            PyValueError::new_err(format!("Invalid coordinate '{}'", coord))
        }
        RpyxlError::InvalidRange(rng) => PyValueError::new_err(format!("Invalid range '{}'", rng)),
        RpyxlError::RowOutOfBounds(row) => PyValueError::new_err(format!(
            "Row {} is out of bounds (allowed: 1-1,048,576)",
            row
        )),
        RpyxlError::ColumnOutOfBounds(col) => PyValueError::new_err(format!(
            "Column {} is out of bounds (allowed: 1-16,384)",
            col
        )),
        RpyxlError::ValueError(msg) => PyValueError::new_err(msg),
        RpyxlError::TypeError(msg) => PyTypeError::new_err(msg),
        RpyxlError::IoError(msg) => pyo3::exceptions::PyIOError::new_err(msg),
        RpyxlError::ZipBombDetected(msg) => {
            PyValueError::new_err(format!("ZIP bomb detected: {}", msg))
        }
        RpyxlError::XmlSecurityViolation(msg) => {
            PyValueError::new_err(format!("XML security violation: {}", msg))
        }
        RpyxlError::InvalidFileFormat(msg) => {
            PyValueError::new_err(format!("Invalid XLSX file format: {}", msg))
        }
    }
}

trait ReadSeek: std::io::Read + std::io::Seek + Send {}
impl<T: std::io::Read + std::io::Seek + Send> ReadSeek for T {}

#[derive(Clone)]
enum WorkbookSource {
    Path(PathBuf),
    Bytes(Arc<[u8]>),
}

impl WorkbookSource {
    fn open(&self) -> RpyxlResult<Box<dyn ReadSeek>> {
        match self {
            Self::Path(path) => std::fs::File::open(path)
                .map(|file| Box::new(std::io::BufReader::new(file)) as Box<dyn ReadSeek>)
                .map_err(|e| RpyxlError::IoError(format!("Failed to open file: {}", e))),
            Self::Bytes(bytes) => Ok(Box::new(Cursor::new(Arc::clone(bytes))) as Box<dyn ReadSeek>),
        }
    }
}

fn filesystem_path(value: &Bound<'_, PyAny>) -> PyResult<Option<PathBuf>> {
    if value.is_instance_of::<PyString>()
        || value.is_instance_of::<PyBytes>()
        || value.hasattr("__fspath__")?
    {
        let os = value.py().import_bound("os")?;
        let path = os.call_method1("fspath", (value,))?;
        return os.call_method1("fsdecode", (path,))?.extract().map(Some);
    }
    Ok(None)
}

fn workbook_source(value: &Bound<'_, PyAny>) -> PyResult<WorkbookSource> {
    if let Some(path) = filesystem_path(value)? {
        let path = value
            .py()
            .allow_threads(move || -> std::io::Result<PathBuf> {
                let path = if path.is_absolute() {
                    path
                } else {
                    std::env::current_dir()?.join(path)
                };
                // Preserve Python's errno-specific exceptions for filesystem inputs.
                std::fs::File::open(&path)?;
                Ok(path)
            })?;
        return Ok(WorkbookSource::Path(path));
    }
    for method in ["read", "seek", "tell"] {
        if !value.hasattr(method)? || !value.getattr(method)?.is_callable() {
            return Err(PyTypeError::new_err(
                "Expected a filesystem path or a seekable binary file-like object with read(), seek(), and tell()",
            ));
        }
    }
    let position = value.call_method0("tell")?;
    let result = (|| -> PyResult<WorkbookSource> {
        value.call_method1("seek", (0,))?;
        let data = value.call_method0("read")?;
        let bytes = data.downcast::<PyBytes>().map_err(|_| {
            PyTypeError::new_err(
                "Workbook input must be opened in binary mode (read() must return bytes)",
            )
        })?;
        Ok(WorkbookSource::Bytes(Arc::from(bytes.as_bytes())))
    })();
    // Restore even when seeking, reading, or validating the returned data fails.
    match value.call_method1("seek", (&position,)) {
        Ok(_) => result,
        Err(restore_error) => {
            if let Err(original_error) = result {
                restore_error.set_cause(value.py(), Some(original_error));
            }
            Err(restore_error)
        }
    }
}

fn check_binary_writer(value: &Bound<'_, PyAny>) -> PyResult<()> {
    for method in ["write", "seek", "truncate"] {
        if !value.hasattr(method)? || !value.getattr(method)?.is_callable() {
            return Err(PyTypeError::new_err(
                "Workbook output must be a seekable binary file-like object with write(), seek(), and truncate()",
            ));
        }
    }
    if value.hasattr("seekable")? && !value.call_method0("seekable")?.is_truthy()? {
        return Err(PyValueError::new_err(
            "Workbook output must be seekable to replace its contents",
        ));
    }
    if value.hasattr("mode")? {
        if let Ok(mode) = value.getattr("mode")?.extract::<String>() {
            if mode.contains('a') {
                return Err(PyValueError::new_err(
                    "Workbook output must not be opened in append mode",
                ));
            }
        }
    }
    Ok(())
}

fn write_binary_output(value: &Bound<'_, PyAny>, bytes: &[u8]) -> PyResult<()> {
    // Saving replaces the whole archive, regardless of the stream's initial position.
    value.call_method1("seek", (0,))?;
    let mut offset = 0;
    while offset < bytes.len() {
        let result =
            value.call_method1("write", (PyBytes::new_bound(value.py(), &bytes[offset..]),))?;
        let written = result.extract::<usize>().map_err(|_| {
            PyTypeError::new_err("Binary write() must return the number of bytes written")
        })?;
        if written == 0 || written > bytes.len() - offset {
            return Err(pyo3::exceptions::PyIOError::new_err(
                "Binary write() returned an invalid byte count",
            ));
        }
        offset += written;
    }
    value.call_method1("truncate", (bytes.len(),))?;
    Ok(())
}

pub fn python_value_to_cell_value(val: &Bound<'_, PyAny>) -> PyResult<CellValue> {
    if val.is_none() {
        Ok(CellValue::Empty)
    } else if val.is_instance_of::<PyBool>() {
        Ok(CellValue::Bool(val.extract::<bool>()?))
    } else if val.is_instance_of::<PyInt>() {
        Ok(CellValue::Int(val.extract::<i64>()?))
    } else if val.is_instance_of::<PyFloat>() {
        Ok(CellValue::Float(val.extract::<f64>()?))
    } else if val.is_instance_of::<PyString>() {
        let s = val.extract::<String>()?;
        if s.starts_with('=') {
            Ok(CellValue::Formula {
                formula: s,
                cached_value: None,
            })
        } else {
            Ok(CellValue::String(s))
        }
    } else {
        let datetime = val.py().import_bound("datetime")?;
        let is_datetime = val.is_instance(&datetime.getattr("datetime")?)?;
        let is_time = val.is_instance(&datetime.getattr("time")?)?;
        let is_date = val.is_instance(&datetime.getattr("date")?)?;
        if is_datetime || is_time || is_date {
            if (is_datetime || is_time) && !val.call_method0("utcoffset")?.is_none() {
                return Err(PyValueError::new_err(
                    "Excel does not support timezone-aware datetime or time values",
                ));
            }
            let s = val.call_method0("isoformat")?.extract::<String>()?;
            return Ok(if is_datetime || is_time {
                // Keep time-only ISO values in the existing variant. The current
                // OOXML writer persists those as text, not numeric Excel times.
                CellValue::DateTime(s)
            } else {
                CellValue::Date(s)
            });
        }
        // Fallback: convert to string representation
        let s = val.str()?.to_string();
        Ok(CellValue::String(s))
    }
}

pub fn cell_value_to_python<'py>(py: Python<'py>, val: &CellValue) -> PyResult<Bound<'py, PyAny>> {
    match val {
        CellValue::Empty => Ok(py.None().into_bound(py)),
        CellValue::String(s) => Ok(s.to_object(py).into_bound(py)),
        CellValue::Int(i) => Ok(i.to_object(py).into_bound(py)),
        CellValue::Float(f) => Ok(f.to_object(py).into_bound(py)),
        CellValue::Bool(b) => Ok(b.to_object(py).into_bound(py)),
        CellValue::Date(d) => date_str_to_pydate(py, d),
        CellValue::DateTime(dt) => date_str_to_pydatetime(py, dt),
        CellValue::Formula { formula, .. } => Ok(formula.to_object(py).into_bound(py)),
    }
}

fn date_str_to_pydate<'py>(py: Python<'py>, s: &str) -> PyResult<Bound<'py, PyAny>> {
    parse_iso_temporal(py, "date", s)
}

fn date_str_to_pydatetime<'py>(py: Python<'py>, s: &str) -> PyResult<Bound<'py, PyAny>> {
    let kind = if s.as_bytes().get(2) == Some(&b':') {
        "time"
    } else {
        "datetime"
    };
    parse_iso_temporal(py, kind, s)
}

fn parse_iso_temporal<'py>(py: Python<'py>, kind: &str, s: &str) -> PyResult<Bound<'py, PyAny>> {
    let cls = py.import_bound("datetime")?.getattr(kind)?;
    // Python 3.9/3.10 fromisoformat does not accept the UTC Z suffix.
    let normalized = s.strip_suffix('Z').map(|s| format!("{}+00:00", s));
    match cls.call_method1("fromisoformat", (normalized.as_deref().unwrap_or(s),)) {
        Ok(value) => Ok(value),
        Err(err) if err.is_instance_of::<PyValueError>(py) => Ok(s.to_object(py).into_bound(py)),
        Err(err) => Err(err),
    }
}

#[pyclass(name = "Cell")]
pub struct PyCell {
    wb: Arc<RwLock<Workbook>>,
    sheet_title: Arc<RwLock<String>>,
    row: u32,
    column: u32,
}

#[pymethods]
impl PyCell {
    #[getter]
    pub fn row(&self) -> u32 {
        self.row
    }

    #[getter]
    pub fn column(&self) -> u32 {
        self.column
    }

    #[getter]
    pub fn coordinate(&self) -> PyResult<String> {
        tuple_to_coordinate(self.row, self.column).map_err(map_rpyxl_err)
    }

    #[getter]
    pub fn column_letter(&self) -> PyResult<String> {
        get_column_letter(self.column).map_err(map_rpyxl_err)
    }

    #[getter]
    pub fn parent(&self) -> PyWorksheet {
        PyWorksheet {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::clone(&self.sheet_title),
        }
    }

    #[getter]
    pub fn is_date(&self) -> PyResult<bool> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        let val = ws.get_cell_value(self.row, self.column);
        Ok(matches!(val, CellValue::Date(_) | CellValue::DateTime(_)))
    }

    #[getter]
    pub fn number_format(&self) -> PyResult<String> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        if let Some(c) = ws.get_cell(self.row, self.column) {
            if let Some(ref nf) = c.number_format {
                return Ok(nf.clone());
            }
            if let CellValue::Date(_) = c.value {
                return Ok("yyyy-mm-dd".to_string());
            }
            if let CellValue::DateTime(_) = c.value {
                return Ok("yyyy-mm-dd h:mm:ss".to_string());
            }
        }
        Ok("General".to_string())
    }

    #[setter]
    pub fn set_number_format(&self, fmt: Option<String>) -> PyResult<()> {
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.set_cell_number_format(self.row, self.column, fmt)
            .map_err(map_rpyxl_err)
    }

    #[getter]
    pub fn comment(&self) -> PyResult<Option<String>> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        Ok(ws
            .get_cell(self.row, self.column)
            .and_then(|c| c.comment.clone()))
    }

    #[setter]
    pub fn set_comment(&self, c: Option<String>) -> PyResult<()> {
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.set_cell_comment(self.row, self.column, c)
            .map_err(map_rpyxl_err)
    }

    #[getter]
    pub fn hyperlink(&self) -> PyResult<Option<String>> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        Ok(ws
            .get_cell(self.row, self.column)
            .and_then(|c| c.hyperlink.clone()))
    }

    #[setter]
    pub fn set_hyperlink(&self, h: Option<String>) -> PyResult<()> {
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.set_cell_hyperlink(self.row, self.column, h)
            .map_err(map_rpyxl_err)
    }

    #[getter]
    pub fn data_type(&self) -> PyResult<&'static str> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        let val = ws.get_cell_value(self.row, self.column);
        Ok(val.data_type())
    }

    #[setter]
    pub fn set_data_type(&self, dt: &str) -> PyResult<()> {
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        let cell = ws
            .get_or_create_cell(self.row, self.column)
            .map_err(map_rpyxl_err)?;
        if dt == "s" {
            let cur_str = match &cell.value {
                CellValue::String(s) => s.clone(),
                CellValue::Int(i) => i.to_string(),
                CellValue::Float(f) => f.to_string(),
                CellValue::Bool(b) => b.to_string(),
                CellValue::Date(d) => d.clone(),
                CellValue::DateTime(dt) => dt.clone(),
                CellValue::Formula { formula, .. } => formula.clone(),
                CellValue::Empty => String::new(),
            };
            cell.value = CellValue::String(cur_str);
        }
        Ok(())
    }

    #[getter]
    pub fn value<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let wb = self.wb.read();
        let data_only = wb.is_data_only();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        let val = ws.get_resolved_cell_value(self.row, self.column, data_only);
        cell_value_to_python(py, &val)
    }

    #[setter]
    pub fn set_value(&self, new_val: &Bound<'_, PyAny>) -> PyResult<()> {
        let val = python_value_to_cell_value(new_val)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.set_cell_value(self.row, self.column, val)
            .map_err(map_rpyxl_err)?;
        Ok(())
    }

    pub fn __repr__(&self) -> PyResult<String> {
        let title = self.sheet_title.read().clone();
        let coord = tuple_to_coordinate(self.row, self.column).map_err(map_rpyxl_err)?;
        Ok(format!("<Cell '{}'.{}>", title, coord))
    }
}

fn structural_edit_args(idx: i64, amount: i64) -> PyResult<(u32, u32)> {
    let idx = u32::try_from(idx)
        .ok()
        .filter(|idx| *idx > 0)
        .ok_or_else(|| PyValueError::new_err("Index must be a positive worksheet index"))?;
    let amount = u32::try_from(amount)
        .map_err(|_| PyValueError::new_err("Amount must be between 0 and 4294967295"))?;
    Ok((idx, amount))
}

#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    wb: Arc<RwLock<Workbook>>,
    sheet_title: Arc<RwLock<String>>,
}

#[pymethods]
impl PyWorksheet {
    #[getter]
    pub fn parent(&self) -> PyWorkbook {
        PyWorkbook {
            wb: Arc::clone(&self.wb),
        }
    }

    #[getter]
    pub fn title(&self) -> String {
        self.sheet_title.read().clone()
    }

    #[setter]
    pub fn set_title(&self, new_title: &str) -> PyResult<()> {
        let old_title = self.sheet_title.read().clone();
        let mut wb = self.wb.write();
        wb.rename_sheet(&old_title, new_title)
            .map_err(map_rpyxl_err)?;
        *self.sheet_title.write() = new_title.to_string();
        Ok(())
    }

    #[getter]
    pub fn min_row(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        Ok(ws.min_row())
    }

    #[getter]
    pub fn max_row(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        Ok(ws.max_row())
    }

    #[getter]
    pub fn min_column(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        Ok(ws.min_column())
    }

    #[getter]
    pub fn max_column(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        Ok(ws.max_column())
    }

    #[getter]
    pub fn dimensions(&self) -> PyResult<String> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.dimensions().map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (row, column, value=None))]
    pub fn cell(
        &self,
        row: u32,
        column: u32,
        value: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyCell> {
        if row < 1 || column < 1 {
            return Err(PyValueError::new_err(
                "Row or column values must be at least 1",
            ));
        }

        {
            let mut wb = self.wb.write();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet_mut(&title)
                .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                .map_err(map_rpyxl_err)?;

            if let Some(v) = value {
                let cell_val = python_value_to_cell_value(v)?;
                ws.set_cell_value(row, column, cell_val)
                    .map_err(map_rpyxl_err)?;
            } else {
                ws.get_or_create_cell(row, column).map_err(map_rpyxl_err)?;
            }
        }

        Ok(PyCell {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::clone(&self.sheet_title),
            row,
            column,
        })
    }

    pub fn __getitem__<'py>(
        &self,
        py: Python<'py>,
        key: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        // 1. Integer row: ws[1]
        if let Ok(r) = key.extract::<u32>() {
            if r < 1 {
                return Err(PyValueError::new_err("Row values must be at least 1"));
            }
            let (min_c, max_c) = {
                let wb = self.wb.read();
                let title = self.sheet_title.read().clone();
                let ws = wb
                    .get_sheet(&title)
                    .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                    .map_err(map_rpyxl_err)?;
                (ws.min_column().max(1), ws.max_column().max(1))
            };
            let mut inner_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
            for c in min_c..=max_c {
                let cell = Bound::new(
                    py,
                    PyCell {
                        wb: Arc::clone(&self.wb),
                        sheet_title: Arc::clone(&self.sheet_title),
                        row: r,
                        column: c,
                    },
                )?;
                inner_cols.push(cell.into_any());
            }
            return Ok(PyTuple::new_bound(py, inner_cols).into_any());
        }

        // 2. PySlice: ws[1:5], ws["A":"C"], ws["A1":"B2"]
        if let Ok(slice) = key.downcast::<PySlice>() {
            let start = slice.getattr("start")?;
            let stop = slice.getattr("stop")?;

            if let (Ok(r1), Ok(r2)) = (start.extract::<u32>(), stop.extract::<u32>()) {
                let min_r = r1.min(r2).max(1);
                let max_r = r1.max(r2);
                let (min_c, max_c) = {
                    let wb = self.wb.read();
                    let title = self.sheet_title.read().clone();
                    let ws = wb
                        .get_sheet(&title)
                        .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                        .map_err(map_rpyxl_err)?;
                    (ws.min_column().max(1), ws.max_column().max(1))
                };
                let mut outer_rows = Vec::with_capacity((max_r - min_r + 1) as usize);
                for r in min_r..=max_r {
                    let mut inner_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                    for c in min_c..=max_c {
                        let cell = Bound::new(
                            py,
                            PyCell {
                                wb: Arc::clone(&self.wb),
                                sheet_title: Arc::clone(&self.sheet_title),
                                row: r,
                                column: c,
                            },
                        )?;
                        inner_cols.push(cell.into_any());
                    }
                    outer_rows.push(PyTuple::new_bound(py, inner_cols));
                }
                return Ok(PyTuple::new_bound(py, outer_rows).into_any());
            }

            if let (Ok(s1), Ok(s2)) = (start.extract::<String>(), stop.extract::<String>()) {
                if s1.chars().all(|c| c.is_ascii_alphabetic())
                    && s2.chars().all(|c| c.is_ascii_alphabetic())
                {
                    let c1 = rpyxl_core::coordinate::column_index_from_string(&s1)
                        .map_err(map_rpyxl_err)?;
                    let c2 = rpyxl_core::coordinate::column_index_from_string(&s2)
                        .map_err(map_rpyxl_err)?;
                    let min_c = c1.min(c2);
                    let max_c = c1.max(c2);
                    let (min_r, max_r) = {
                        let wb = self.wb.read();
                        let title = self.sheet_title.read().clone();
                        let ws = wb
                            .get_sheet(&title)
                            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                            .map_err(map_rpyxl_err)?;
                        (ws.min_row().max(1), ws.max_row().max(1))
                    };
                    let mut outer_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                    for c in min_c..=max_c {
                        let mut inner_rows = Vec::with_capacity((max_r - min_r + 1) as usize);
                        for r in min_r..=max_r {
                            let cell = Bound::new(
                                py,
                                PyCell {
                                    wb: Arc::clone(&self.wb),
                                    sheet_title: Arc::clone(&self.sheet_title),
                                    row: r,
                                    column: c,
                                },
                            )?;
                            inner_rows.push(cell.into_any());
                        }
                        outer_cols.push(PyTuple::new_bound(py, inner_rows));
                    }
                    return Ok(PyTuple::new_bound(py, outer_cols).into_any());
                } else {
                    let (r1, c1) = coordinate_to_tuple(&s1).map_err(map_rpyxl_err)?;
                    let (r2, c2) = coordinate_to_tuple(&s2).map_err(map_rpyxl_err)?;
                    let min_r = r1.min(r2);
                    let max_r = r1.max(r2);
                    let min_c = c1.min(c2);
                    let max_c = c1.max(c2);
                    let mut outer_rows = Vec::with_capacity((max_r - min_r + 1) as usize);
                    for r in min_r..=max_r {
                        let mut inner_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                        for c in min_c..=max_c {
                            let cell = Bound::new(
                                py,
                                PyCell {
                                    wb: Arc::clone(&self.wb),
                                    sheet_title: Arc::clone(&self.sheet_title),
                                    row: r,
                                    column: c,
                                },
                            )?;
                            inner_cols.push(cell.into_any());
                        }
                        outer_rows.push(PyTuple::new_bound(py, inner_cols));
                    }
                    return Ok(PyTuple::new_bound(py, outer_rows).into_any());
                }
            }
        }

        // 3. String key
        let key_str = key.extract::<String>()?;
        if key_str.contains(':') {
            let parts: Vec<&str> = key_str.split(':').collect();
            if parts.len() == 2
                && parts[0].chars().all(|c| c.is_ascii_digit())
                && parts[1].chars().all(|c| c.is_ascii_digit())
            {
                let r1: u32 = parts[0].parse().unwrap_or(1);
                let r2: u32 = parts[1].parse().unwrap_or(1);
                let min_r = r1.min(r2).max(1);
                let max_r = r1.max(r2);
                let (min_c, max_c) = {
                    let wb = self.wb.read();
                    let title = self.sheet_title.read().clone();
                    let ws = wb
                        .get_sheet(&title)
                        .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                        .map_err(map_rpyxl_err)?;
                    (ws.min_column().max(1), ws.max_column().max(1))
                };
                let mut outer_rows = Vec::with_capacity((max_r - min_r + 1) as usize);
                for r in min_r..=max_r {
                    let mut inner_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                    for c in min_c..=max_c {
                        let cell = Bound::new(
                            py,
                            PyCell {
                                wb: Arc::clone(&self.wb),
                                sheet_title: Arc::clone(&self.sheet_title),
                                row: r,
                                column: c,
                            },
                        )?;
                        inner_cols.push(cell.into_any());
                    }
                    outer_rows.push(PyTuple::new_bound(py, inner_cols));
                }
                Ok(PyTuple::new_bound(py, outer_rows).into_any())
            } else if parts.len() == 2
                && parts[0].chars().all(|c| c.is_ascii_alphabetic())
                && parts[1].chars().all(|c| c.is_ascii_alphabetic())
            {
                let c1 = rpyxl_core::coordinate::column_index_from_string(parts[0])
                    .map_err(map_rpyxl_err)?;
                let c2 = rpyxl_core::coordinate::column_index_from_string(parts[1])
                    .map_err(map_rpyxl_err)?;
                let min_c = c1.min(c2);
                let max_c = c1.max(c2);
                let (min_r, max_r) = {
                    let wb = self.wb.read();
                    let title = self.sheet_title.read().clone();
                    let ws = wb
                        .get_sheet(&title)
                        .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                        .map_err(map_rpyxl_err)?;
                    (ws.min_row().max(1), ws.max_row().max(1))
                };
                let mut outer_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                for c in min_c..=max_c {
                    let mut inner_rows = Vec::with_capacity((max_r - min_r + 1) as usize);
                    for r in min_r..=max_r {
                        let cell = Bound::new(
                            py,
                            PyCell {
                                wb: Arc::clone(&self.wb),
                                sheet_title: Arc::clone(&self.sheet_title),
                                row: r,
                                column: c,
                            },
                        )?;
                        inner_rows.push(cell.into_any());
                    }
                    outer_cols.push(PyTuple::new_bound(py, inner_rows));
                }
                Ok(PyTuple::new_bound(py, outer_cols).into_any())
            } else {
                let (min_r, min_c, max_r, max_c) =
                    range_boundaries(&key_str).map_err(map_rpyxl_err)?;
                let mut outer_rows = Vec::with_capacity((max_r - min_r + 1) as usize);

                for r in min_r..=max_r {
                    let mut inner_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                    for c in min_c..=max_c {
                        let cell = Bound::new(
                            py,
                            PyCell {
                                wb: Arc::clone(&self.wb),
                                sheet_title: Arc::clone(&self.sheet_title),
                                row: r,
                                column: c,
                            },
                        )?;
                        inner_cols.push(cell.into_any());
                    }
                    outer_rows.push(PyTuple::new_bound(py, inner_cols));
                }
                Ok(PyTuple::new_bound(py, outer_rows).into_any())
            }
        } else if key_str.chars().all(|c| c.is_ascii_alphabetic()) {
            let col = rpyxl_core::coordinate::column_index_from_string(&key_str)
                .map_err(map_rpyxl_err)?;
            let (min_r, max_r) = {
                let wb = self.wb.read();
                let title = self.sheet_title.read().clone();
                let ws = wb
                    .get_sheet(&title)
                    .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                    .map_err(map_rpyxl_err)?;
                (ws.min_row().max(1), ws.max_row().max(1))
            };
            let mut col_cells = Vec::with_capacity((max_r - min_r + 1) as usize);
            for r in min_r..=max_r {
                let cell = Bound::new(
                    py,
                    PyCell {
                        wb: Arc::clone(&self.wb),
                        sheet_title: Arc::clone(&self.sheet_title),
                        row: r,
                        column: col,
                    },
                )?;
                col_cells.push(cell.into_any());
            }
            Ok(PyTuple::new_bound(py, col_cells).into_any())
        } else {
            let (r, c) = coordinate_to_tuple(&key_str).map_err(map_rpyxl_err)?;
            {
                let mut wb = self.wb.write();
                let title = self.sheet_title.read().clone();
                let ws = wb
                    .get_sheet_mut(&title)
                    .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                    .map_err(map_rpyxl_err)?;
                ws.get_or_create_cell(r, c).map_err(map_rpyxl_err)?;
            }
            let cell = Bound::new(
                py,
                PyCell {
                    wb: Arc::clone(&self.wb),
                    sheet_title: Arc::clone(&self.sheet_title),
                    row: r,
                    column: c,
                },
            )?;
            Ok(cell.into_any())
        }
    }

    pub fn __setitem__(&self, key: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let (r, c) = coordinate_to_tuple(key).map_err(map_rpyxl_err)?;
        let cell_val = python_value_to_cell_value(value)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.set_cell_value(r, c, cell_val).map_err(map_rpyxl_err)?;
        Ok(())
    }

    pub fn append(&self, iterable: &Bound<'_, PyAny>) -> PyResult<u32> {
        if let Ok(dict) = iterable.downcast::<PyDict>() {
            let mut col_vals: Vec<(u32, CellValue)> = Vec::new();
            for (k, v) in dict.iter() {
                let col_idx: u32 = if let Ok(col_num) = k.extract::<u32>() {
                    if col_num < 1 {
                        return Err(PyValueError::new_err("Column values must be at least 1"));
                    }
                    col_num
                } else if let Ok(col_str) = k.extract::<String>() {
                    rpyxl_core::coordinate::column_index_from_string(&col_str)
                        .map_err(map_rpyxl_err)?
                } else {
                    return Err(PyTypeError::new_err(
                        "Dict keys in append must be column letters or integer column numbers",
                    ));
                };
                let val = python_value_to_cell_value(&v)?;
                col_vals.push((col_idx, val));
            }
            col_vals.sort_by_key(|(col, _)| *col);
            let max_c = col_vals.iter().map(|(c, _)| *c).max().unwrap_or(0);
            let mut full_row = vec![CellValue::Empty; max_c as usize];
            for (c, val) in col_vals {
                if c >= 1 && (c as usize) <= full_row.len() {
                    full_row[(c - 1) as usize] = val;
                }
            }
            let mut wb = self.wb.write();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet_mut(&title)
                .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                .map_err(map_rpyxl_err)?;
            return ws.append(full_row).map_err(map_rpyxl_err);
        }

        let mut values = Vec::new();
        for item in iterable.iter()? {
            let item_val = item?;
            values.push(python_value_to_cell_value(&item_val)?);
        }

        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
            .map_err(map_rpyxl_err)?;
        ws.append(values).map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (idx, amount=1))]
    pub fn insert_rows(&self, idx: i64, amount: i64) -> PyResult<()> {
        let (idx, amount) = structural_edit_args(idx, amount)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| map_rpyxl_err(RpyxlError::SheetNotFound(title)))?;
        ws.insert_rows(idx, amount).map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (idx, amount=1))]
    pub fn delete_rows(&self, idx: i64, amount: i64) -> PyResult<()> {
        let (idx, amount) = structural_edit_args(idx, amount)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| map_rpyxl_err(RpyxlError::SheetNotFound(title)))?;
        ws.delete_rows(idx, amount).map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (idx, amount=1))]
    pub fn insert_cols(&self, idx: i64, amount: i64) -> PyResult<()> {
        let (idx, amount) = structural_edit_args(idx, amount)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| map_rpyxl_err(RpyxlError::SheetNotFound(title)))?;
        ws.insert_cols(idx, amount).map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (idx, amount=1))]
    pub fn delete_cols(&self, idx: i64, amount: i64) -> PyResult<()> {
        let (idx, amount) = structural_edit_args(idx, amount)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| map_rpyxl_err(RpyxlError::SheetNotFound(title)))?;
        ws.delete_cols(idx, amount).map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (cell_range, rows=0, cols=0, translate=false))]
    pub fn move_range(
        &self,
        cell_range: &str,
        rows: i64,
        cols: i64,
        translate: bool,
    ) -> PyResult<()> {
        if translate {
            return Err(pyo3::exceptions::PyNotImplementedError::new_err(
                "Formula translation in move_range is not implemented",
            ));
        }
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| map_rpyxl_err(RpyxlError::SheetNotFound(title)))?;
        ws.move_range(cell_range, rows, cols).map_err(map_rpyxl_err)
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    pub fn iter_rows<'py>(
        &self,
        py: Python<'py>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let (is_empty, (r_start, r_end, c_start, c_end)) = {
            let wb = self.wb.read();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                .map_err(map_rpyxl_err)?;
            (
                ws.is_empty(),
                ws.resolve_bounds(min_row, max_row, min_col, max_col),
            )
        };

        if (is_empty
            && min_row.is_none()
            && max_row.is_none()
            && min_col.is_none()
            && max_col.is_none())
            || r_start > r_end
            || c_start > c_end
        {
            return Ok(PyList::empty_bound(py).into_any());
        }

        let mut rows = Vec::with_capacity((r_end - r_start + 1) as usize);

        if values_only {
            let wb = self.wb.read();
            let data_only = wb.is_data_only();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                .map_err(map_rpyxl_err)?;

            for r in r_start..=r_end {
                let mut row_vals = Vec::with_capacity((c_end - c_start + 1) as usize);
                for c in c_start..=c_end {
                    let val = ws.get_resolved_cell_value(r, c, data_only);
                    row_vals.push(cell_value_to_python(py, &val)?);
                }
                rows.push(PyTuple::new_bound(py, row_vals));
            }
        } else {
            for r in r_start..=r_end {
                let mut row_cells = Vec::with_capacity((c_end - c_start + 1) as usize);
                for c in c_start..=c_end {
                    let cell = Bound::new(
                        py,
                        PyCell {
                            wb: Arc::clone(&self.wb),
                            sheet_title: Arc::clone(&self.sheet_title),
                            row: r,
                            column: c,
                        },
                    )?;
                    row_cells.push(cell.into_any());
                }
                rows.push(PyTuple::new_bound(py, row_cells));
            }
        }

        let py_list = PyList::new_bound(py, rows);
        Ok(py_list.into_any())
    }

    #[pyo3(signature = (min_col=None, max_col=None, min_row=None, max_row=None, values_only=false))]
    pub fn iter_cols<'py>(
        &self,
        py: Python<'py>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        values_only: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let (is_empty, (r_start, r_end, c_start, c_end)) = {
            let wb = self.wb.read();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                .map_err(map_rpyxl_err)?;
            (
                ws.is_empty(),
                ws.resolve_bounds(min_row, max_row, min_col, max_col),
            )
        };

        if (is_empty
            && min_row.is_none()
            && max_row.is_none()
            && min_col.is_none()
            && max_col.is_none())
            || r_start > r_end
            || c_start > c_end
        {
            return Ok(PyList::empty_bound(py).into_any());
        }

        let mut cols = Vec::with_capacity((c_end - c_start + 1) as usize);

        if values_only {
            let wb = self.wb.read();
            let data_only = wb.is_data_only();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| RpyxlError::SheetNotFound(title.clone()))
                .map_err(map_rpyxl_err)?;

            for c in c_start..=c_end {
                let mut col_vals = Vec::with_capacity((r_end - r_start + 1) as usize);
                for r in r_start..=r_end {
                    let val = ws.get_resolved_cell_value(r, c, data_only);
                    col_vals.push(cell_value_to_python(py, &val)?);
                }
                cols.push(PyTuple::new_bound(py, col_vals));
            }
        } else {
            for c in c_start..=c_end {
                let mut col_cells = Vec::with_capacity((r_end - r_start + 1) as usize);
                for r in r_start..=r_end {
                    let cell = Bound::new(
                        py,
                        PyCell {
                            wb: Arc::clone(&self.wb),
                            sheet_title: Arc::clone(&self.sheet_title),
                            row: r,
                            column: c,
                        },
                    )?;
                    col_cells.push(cell.into_any());
                }
                cols.push(PyTuple::new_bound(py, col_cells));
            }
        }

        let py_list = PyList::new_bound(py, cols);
        Ok(py_list.into_any())
    }

    #[getter]
    pub fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.iter_rows(py, None, None, None, None, true)
    }

    #[getter]
    pub fn rows<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.iter_rows(py, None, None, None, None, false)
    }

    #[getter]
    pub fn columns<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.iter_cols(py, None, None, None, None, false)
    }

    pub fn __repr__(&self) -> String {
        format!("<Worksheet \"{}\">", self.sheet_title.read())
    }
}

#[pyclass(name = "Workbook")]
pub struct PyWorkbook {
    wb: Arc<RwLock<Workbook>>,
}

#[pymethods]
impl PyWorkbook {
    #[new]
    pub fn new() -> Self {
        Self {
            wb: Arc::new(RwLock::new(Workbook::new())),
        }
    }

    #[getter]
    pub fn read_only(&self) -> bool {
        self.wb.read().is_read_only()
    }

    #[getter]
    pub fn write_only(&self) -> bool {
        false
    }

    #[getter]
    pub fn data_only(&self) -> bool {
        self.wb.read().is_data_only()
    }

    #[getter]
    pub fn active(&self) -> PyResult<PyWorksheet> {
        let wb = self.wb.read();
        let title = wb.active().title.clone();
        Ok(PyWorksheet {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::new(RwLock::new(title)),
        })
    }

    #[setter]
    pub fn set_active(&self, target: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.wb.write();
        if let Ok(idx) = target.extract::<usize>() {
            wb.set_active_index(idx).map_err(map_rpyxl_err)?;
        } else if let Ok(ws) = target.extract::<PyRef<'_, PyWorksheet>>() {
            let title = ws.sheet_title.read().clone();
            wb.set_active_by_title(&title).map_err(map_rpyxl_err)?;
        } else if let Ok(title) = target.extract::<String>() {
            wb.set_active_by_title(&title).map_err(map_rpyxl_err)?;
        } else {
            return Err(PyTypeError::new_err(
                "active sheet must be set with an index, sheet name, or Worksheet object",
            ));
        }
        Ok(())
    }

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.wb.read().sheetnames()
    }

    #[getter]
    pub fn worksheets(&self) -> Vec<PyWorksheet> {
        let wb = self.wb.read();
        wb.sheetnames()
            .into_iter()
            .map(|title| PyWorksheet {
                wb: Arc::clone(&self.wb),
                sheet_title: Arc::new(RwLock::new(title)),
            })
            .collect()
    }

    #[pyo3(signature = (title=None, index=None))]
    pub fn create_sheet(&self, title: Option<&str>, index: Option<usize>) -> PyResult<PyWorksheet> {
        let mut wb = self.wb.write();
        let ws = wb.create_sheet(title, index).map_err(map_rpyxl_err)?;
        let sheet_title = ws.title.clone();
        Ok(PyWorksheet {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::new(RwLock::new(sheet_title)),
        })
    }

    pub fn copy_worksheet(&self, worksheet: &Bound<'_, PyWorksheet>) -> PyResult<PyWorksheet> {
        let from_title = worksheet.borrow().sheet_title.read().clone();
        let mut wb = self.wb.write();
        let ws = wb.copy_worksheet(&from_title).map_err(map_rpyxl_err)?;
        let sheet_title = ws.title.clone();
        Ok(PyWorksheet {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::new(RwLock::new(sheet_title)),
        })
    }

    pub fn index(&self, worksheet: &Bound<'_, PyWorksheet>) -> PyResult<usize> {
        let ws = worksheet.borrow();
        if !Arc::ptr_eq(&self.wb, &ws.wb) {
            return Err(PyValueError::new_err(format!(
                "Worksheet '{}' is not in workbook",
                ws.sheet_title.read()
            )));
        }
        let title = ws.sheet_title.read().clone();
        let wb = self.wb.read();
        wb.index_of(&title).ok_or_else(|| {
            PyValueError::new_err(format!("Worksheet '{}' is not in workbook", title))
        })
    }

    pub fn remove(&self, worksheet: &Bound<'_, PyWorksheet>) -> PyResult<()> {
        let ws = worksheet.borrow();
        let title = ws.sheet_title.read().clone();
        let mut wb = self.wb.write();
        wb.remove_sheet(&title).map_err(map_rpyxl_err)?;
        Ok(())
    }

    pub fn remove_sheet(&self, title: &str) -> PyResult<()> {
        let mut wb = self.wb.write();
        wb.remove_sheet(title).map_err(map_rpyxl_err)?;
        Ok(())
    }

    pub fn __getitem__(&self, title: &str) -> PyResult<PyWorksheet> {
        let wb = self.wb.read();
        if wb.contains_sheet(title) {
            Ok(PyWorksheet {
                wb: Arc::clone(&self.wb),
                sheet_title: Arc::new(RwLock::new(title.to_string())),
            })
        } else {
            Err(PyKeyError::new_err(format!(
                "Worksheet {} does not exist.",
                title
            )))
        }
    }

    pub fn __contains__(&self, title: &str) -> bool {
        self.wb.read().contains_sheet(title)
    }

    pub fn __delitem__(&self, title: &str) -> PyResult<()> {
        let mut wb = self.wb.write();
        wb.remove_sheet(title).map_err(map_rpyxl_err)?;
        Ok(())
    }

    pub fn __len__(&self) -> usize {
        self.wb.read().sheet_count()
    }

    pub fn __iter__(slf: PyRef<'_, Self>) -> PyWorkbookIter {
        let wb = slf.wb.read();
        let titles = wb.sheetnames();
        PyWorkbookIter {
            wb: Arc::clone(&slf.wb),
            sheet_titles: titles,
            index: 0,
        }
    }

    pub fn close(&self) -> PyResult<()> {
        Ok(())
    }

    pub fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (_exc_type=None, _exc_val=None, _exc_tb=None))]
    pub fn __exit__(
        &self,
        _exc_type: Option<&Bound<'_, PyAny>>,
        _exc_val: Option<&Bound<'_, PyAny>>,
        _exc_tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }

    fn _replace_cell_styles(
        &self,
        py: Python<'_>,
        records: Vec<RichStyleInput>,
        assignments: Vec<(String, u32, u32, usize)>,
    ) -> PyResult<()> {
        let styles: Vec<_> = records
            .into_iter()
            .map(|(font, fill, border, alignment, protection)| {
                Arc::new(RichStyle {
                    font,
                    fill,
                    border,
                    alignment,
                    protection,
                })
            })
            .collect();
        let workbook = Arc::clone(&self.wb);
        py.allow_threads(move || -> RpyxlResult<()> {
            let mut wb = workbook.write();
            // Validate the complete replacement before clearing or creating cells.
            for (title, row, column, index) in &assignments {
                if wb.get_sheet(title).is_none() {
                    return Err(RpyxlError::SheetNotFound(title.clone()));
                }
                if !(1..=MAX_ROW).contains(row) {
                    return Err(RpyxlError::RowOutOfBounds(*row));
                }
                if !(1..=MAX_COLUMN).contains(column) {
                    return Err(RpyxlError::ColumnOutOfBounds(*column));
                }
                if *index >= styles.len() {
                    return Err(RpyxlError::ValueError(
                        "Invalid rich-style record index".into(),
                    ));
                }
            }
            for sheet in wb.sheets_mut() {
                sheet.clear_rich_styles();
            }
            for (title, row, column, index) in assignments {
                let sheet = wb.get_sheet_mut(&title).expect("validated worksheet");
                sheet.get_or_create_cell(row, column)?.rich_style =
                    Some(Arc::clone(&styles[index]));
            }
            Ok(())
        })
        .map_err(map_rpyxl_err)
    }

    pub fn save(&self, path: &Bound<'_, PyAny>) -> PyResult<()> {
        let py = path.py();
        let wb = Arc::clone(&self.wb);
        if let Some(path) = filesystem_path(path)? {
            return py
                .allow_threads(move || save_workbook(path, &wb.read()))
                .map_err(map_rpyxl_err);
        }
        check_binary_writer(path)?;
        let bytes = py
            .allow_threads(move || {
                let mut buffer = Cursor::new(Vec::new());
                write_workbook(&mut buffer, &wb.read())?;
                Ok::<_, RpyxlError>(buffer.into_inner())
            })
            .map_err(map_rpyxl_err)?;
        write_binary_output(path, &bytes)
    }

    pub fn __repr__(&self) -> String {
        format!("<Workbook sheetnames={:?}>", self.wb.read().sheetnames())
    }
}

impl Default for PyWorkbook {
    fn default() -> Self {
        Self::new()
    }
}

#[pyclass]
pub struct PyWorkbookIter {
    wb: Arc<RwLock<Workbook>>,
    sheet_titles: Vec<String>,
    index: usize,
}

#[pymethods]
impl PyWorkbookIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<PyWorksheet> {
        if self.index < self.sheet_titles.len() {
            let title = self.sheet_titles[self.index].clone();
            self.index += 1;
            Some(PyWorksheet {
                wb: Arc::clone(&self.wb),
                sheet_title: Arc::new(RwLock::new(title)),
            })
        } else {
            None
        }
    }
}

// -------------------------------------------------------------
// Read-Only Streaming Classes
// -------------------------------------------------------------

#[pyclass(name = "ReadOnlyCell")]
pub struct PyReadOnlyCell {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub column: u32,
    coordinate: String,
    column_letter: String,
    value: PyObject,
    data_type: &'static str,
}

#[pymethods]
impl PyReadOnlyCell {
    #[getter]
    pub fn coordinate(&self) -> &str {
        &self.coordinate
    }

    #[getter]
    pub fn column_letter(&self) -> &str {
        &self.column_letter
    }

    #[getter]
    pub fn value<'py>(&self, py: Python<'py>) -> PyObject {
        self.value.clone_ref(py)
    }

    #[setter]
    pub fn set_value(&self, _val: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyTypeError::new_err("Cannot assign to read-only cell"))
    }

    #[getter]
    pub fn data_type(&self) -> &'static str {
        self.data_type
    }

    pub fn __repr__(&self) -> String {
        format!("<ReadOnlyCell {}.{}>", self.coordinate, self.row)
    }
}

pub type StreamingRowMessage = RpyxlResult<Option<Vec<(u32, CellValue)>>>;

#[pyclass(name = "RowIterator")]
pub struct PyRowIterator {
    rx: std::sync::mpsc::Receiver<StreamingRowMessage>,
    min_col: u32,
    max_col: Option<u32>,
    values_only: bool,
    data_only: bool,
    current_row_idx: u32,
}

#[pymethods]
impl PyRowIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        let rx = &mut self.rx;
        match py.allow_threads(move || rx.recv()) {
            Ok(Ok(Some(cells))) => {
                let row_num = self.current_row_idx;
                self.current_row_idx += 1;

                let effective_max_col = self
                    .max_col
                    .unwrap_or_else(|| cells.iter().map(|(c, _)| *c).max().unwrap_or(self.min_col));
                let width = (effective_max_col.saturating_sub(self.min_col) + 1) as usize;
                let mut cell_map = HashMap::with_capacity(cells.len());
                for (col, val) in cells {
                    cell_map.insert(col, val);
                }

                let mut row_items = Vec::with_capacity(width);
                if self.values_only {
                    for col in self.min_col..=effective_max_col {
                        if let Some(val) = cell_map.remove(&col) {
                            let resolved = match val {
                                CellValue::Formula {
                                    cached_value,
                                    formula,
                                } => {
                                    if self.data_only {
                                        cached_value.map(|b| *b).unwrap_or(CellValue::Empty)
                                    } else {
                                        CellValue::Formula {
                                            cached_value,
                                            formula,
                                        }
                                    }
                                }
                                other => other,
                            };
                            row_items.push(cell_value_to_python(py, &resolved)?);
                        } else {
                            row_items.push(py.None().into_bound(py));
                        }
                    }
                } else {
                    for col in self.min_col..=effective_max_col {
                        let coord = tuple_to_coordinate(row_num, col).map_err(map_rpyxl_err)?;
                        let col_letter = get_column_letter(col).map_err(map_rpyxl_err)?;
                        if let Some(val) = cell_map.remove(&col) {
                            let dt = val.data_type();
                            let resolved = match val {
                                CellValue::Formula {
                                    cached_value,
                                    formula,
                                } => {
                                    if self.data_only {
                                        cached_value.map(|b| *b).unwrap_or(CellValue::Empty)
                                    } else {
                                        CellValue::Formula {
                                            cached_value,
                                            formula,
                                        }
                                    }
                                }
                                other => other,
                            };
                            let py_val = cell_value_to_python(py, &resolved)?.unbind();
                            let cell = Bound::new(
                                py,
                                PyReadOnlyCell {
                                    row: row_num,
                                    column: col,
                                    coordinate: coord,
                                    column_letter: col_letter,
                                    value: py_val,
                                    data_type: dt,
                                },
                            )?;
                            row_items.push(cell.into_any());
                        } else {
                            let cell = Bound::new(
                                py,
                                PyReadOnlyCell {
                                    row: row_num,
                                    column: col,
                                    coordinate: coord,
                                    column_letter: col_letter,
                                    value: py.None(),
                                    data_type: "n",
                                },
                            )?;
                            row_items.push(cell.into_any());
                        }
                    }
                }
                Ok(Some(PyTuple::new_bound(py, row_items)))
            }
            Ok(Ok(None)) | Err(_) => Ok(None),
            Ok(Err(e)) => Err(map_rpyxl_err(e)),
        }
    }
}

#[pyclass(name = "ReadOnlyWorksheet")]
pub struct PyReadOnlyWorksheet {
    source: WorkbookSource,
    title: String,
    target_path: String,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    min_row: Option<u32>,
    max_row: Option<u32>,
    min_col: Option<u32>,
    max_col: Option<u32>,
    data_only: bool,
}

#[pymethods]
impl PyReadOnlyWorksheet {
    #[getter]
    pub fn title(&self) -> &str {
        &self.title
    }

    #[getter]
    pub fn min_row(&self) -> Option<u32> {
        self.min_row
    }

    #[getter]
    pub fn max_row(&self) -> Option<u32> {
        self.max_row
    }

    #[getter]
    pub fn min_column(&self) -> Option<u32> {
        self.min_col
    }

    #[getter]
    pub fn max_column(&self) -> Option<u32> {
        self.max_col
    }

    #[getter]
    pub fn dimensions(&self) -> PyResult<String> {
        let start = tuple_to_coordinate(self.min_row.unwrap_or(1), self.min_col.unwrap_or(1))
            .map_err(map_rpyxl_err)?;
        let end = tuple_to_coordinate(self.max_row.unwrap_or(1), self.max_col.unwrap_or(1))
            .map_err(map_rpyxl_err)?;
        Ok(format!("{}:{}", start, end))
    }

    pub fn calculate_dimension(&self) -> PyResult<String> {
        self.dimensions()
    }

    pub fn cell<'py>(&self, py: Python<'py>, row: u32, column: u32) -> PyResult<PyReadOnlyCell> {
        let cell_val = py
            .allow_threads(|| {
                let mut archive = zip::ZipArchive::new(self.source.open()?)
                    .map_err(|e| RpyxlError::IoError(format!("Invalid zip: {}", e)))?;
                rpyxl_ooxml::validate_zip_archive_security(&mut archive)?;
                let sheet_file = archive
                    .by_name(&self.target_path)
                    .map_err(|e| RpyxlError::IoError(format!("Cannot open sheet XML: {}", e)))?;
                read_single_cell(
                    sheet_file,
                    row,
                    column,
                    &self.shared_strings,
                    &self.styles,
                    self.data_only,
                )
            })
            .map_err(map_rpyxl_err)?;

        let dt = cell_val.data_type();
        let py_val = cell_value_to_python(py, &cell_val)?.unbind();
        let coord = tuple_to_coordinate(row, column).map_err(map_rpyxl_err)?;
        let col_letter = get_column_letter(column).map_err(map_rpyxl_err)?;

        Ok(PyReadOnlyCell {
            row,
            column,
            coordinate: coord,
            column_letter: col_letter,
            value: py_val,
            data_type: dt,
        })
    }

    pub fn __getitem__<'py>(
        &self,
        py: Python<'py>,
        key: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let key_str = key.extract::<String>()?;
        if key_str.contains(':') {
            let (min_r, min_c, max_r, max_c) = range_boundaries(&key_str).map_err(map_rpyxl_err)?;
            let mut outer = Vec::with_capacity((max_r - min_r + 1) as usize);
            for r in min_r..=max_r {
                let mut inner = Vec::with_capacity((max_c - min_c + 1) as usize);
                for c in min_c..=max_c {
                    let cell = self.cell(py, r, c)?;
                    inner.push(Bound::new(py, cell)?.into_any());
                }
                outer.push(PyTuple::new_bound(py, inner));
            }
            Ok(PyTuple::new_bound(py, outer).into_any())
        } else {
            let (r, c) = coordinate_to_tuple(&key_str).map_err(map_rpyxl_err)?;
            let cell = self.cell(py, r, c)?;
            Ok(Bound::new(py, cell)?.into_any())
        }
    }

    pub fn __setitem__(&self, _key: &str, _value: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyTypeError::new_err("Worksheet is read-only"))
    }

    pub fn append(&self, _iterable: &Bound<'_, PyAny>) -> PyResult<u32> {
        Err(PyTypeError::new_err("Worksheet is read-only"))
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    pub fn iter_rows(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<PyRowIterator> {
        let r_start = min_row.or(self.min_row).unwrap_or(1);
        let r_end = max_row.or(self.max_row).unwrap_or(1_048_576);
        let c_start = min_col.or(self.min_col).unwrap_or(1);
        let c_end_filter = max_col.or(self.max_col).unwrap_or(16_384);
        let c_end_spec = max_col.or(self.max_col);

        let (tx, rx) =
            std::sync::mpsc::sync_channel::<RpyxlResult<Option<Vec<(u32, CellValue)>>>>(64);

        let source = self.source.clone();
        let target_path = self.target_path.clone();
        let shared_strings = Arc::clone(&self.shared_strings);
        let styles = Arc::clone(&self.styles);
        let data_only = self.data_only;

        std::thread::spawn(move || {
            let reader = match source.open() {
                Ok(reader) => reader,
                Err(e) => {
                    let _ = tx.send(Err(e));
                    return;
                }
            };
            let mut archive = match zip::ZipArchive::new(reader) {
                Ok(a) => a,
                Err(e) => {
                    let _ = tx.send(Err(RpyxlError::IoError(format!("Invalid zip: {}", e))));
                    return;
                }
            };
            if let Err(e) = rpyxl_ooxml::validate_zip_archive_security(&mut archive) {
                let _ = tx.send(Err(e));
                return;
            }
            let sheet_file = match archive.by_name(&target_path) {
                Ok(f) => f,
                Err(e) => {
                    let _ = tx.send(Err(RpyxlError::IoError(format!(
                        "Cannot open {}: {}",
                        target_path, e
                    ))));
                    return;
                }
            };

            let mut reader = StreamingRowReader::new(
                sheet_file,
                shared_strings,
                styles,
                r_start,
                r_end,
                c_start,
                c_end_filter,
                data_only,
            );

            let mut next_expected = r_start;
            loop {
                match reader.next_xml_row() {
                    Ok(Some((xml_row, cells))) => {
                        while next_expected < xml_row {
                            if tx.send(Ok(Some(Vec::new()))).is_err() {
                                return;
                            }
                            next_expected += 1;
                        }
                        if tx.send(Ok(Some(cells))).is_err() {
                            return;
                        }
                        next_expected = xml_row + 1;
                    }
                    Ok(None) => break,
                    Err(e) => {
                        let _ = tx.send(Err(e));
                        return;
                    }
                }
            }
            let _ = tx.send(Ok(None));
        });

        Ok(PyRowIterator {
            rx,
            min_col: c_start,
            max_col: c_end_spec,
            values_only,
            data_only: self.data_only,
            current_row_idx: r_start,
        })
    }

    #[getter]
    pub fn values(&self) -> PyResult<PyRowIterator> {
        self.iter_rows(None, None, None, None, true)
    }

    #[getter]
    pub fn rows(&self) -> PyResult<PyRowIterator> {
        self.iter_rows(None, None, None, None, false)
    }

    pub fn __repr__(&self) -> String {
        format!("<ReadOnlyWorksheet \"{}\">", self.title)
    }
}

#[pyclass(name = "ReadOnlyWorkbook")]
pub struct PyReadOnlyWorkbook {
    source: WorkbookSource,
    sheet_infos: Vec<SheetInfo>,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    active_index: usize,
    data_only: bool,
    is_closed: bool,
}

#[pymethods]
impl PyReadOnlyWorkbook {
    #[getter]
    pub fn read_only(&self) -> bool {
        true
    }

    #[getter]
    pub fn data_only(&self) -> bool {
        self.data_only
    }

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_infos.iter().map(|s| s.name.clone()).collect()
    }

    #[getter]
    pub fn active(&self) -> PyResult<PyReadOnlyWorksheet> {
        if self.sheet_infos.is_empty() {
            return Err(PyKeyError::new_err("Workbook has no sheets"));
        }
        let idx = if self.active_index < self.sheet_infos.len() {
            self.active_index
        } else {
            0
        };
        let info = &self.sheet_infos[idx];
        Ok(PyReadOnlyWorksheet {
            source: self.source.clone(),
            title: info.name.clone(),
            target_path: info.target_path.clone(),
            shared_strings: Arc::clone(&self.shared_strings),
            styles: Arc::clone(&self.styles),
            min_row: info.min_row,
            max_row: info.max_row,
            min_col: info.min_col,
            max_col: info.max_col,
            data_only: self.data_only,
        })
    }

    #[getter]
    pub fn worksheets(&self) -> Vec<PyReadOnlyWorksheet> {
        self.sheet_infos
            .iter()
            .map(|info| PyReadOnlyWorksheet {
                source: self.source.clone(),
                title: info.name.clone(),
                target_path: info.target_path.clone(),
                shared_strings: Arc::clone(&self.shared_strings),
                styles: Arc::clone(&self.styles),
                min_row: info.min_row,
                max_row: info.max_row,
                min_col: info.min_col,
                max_col: info.max_col,
                data_only: self.data_only,
            })
            .collect()
    }

    pub fn __getitem__(&self, title: &str) -> PyResult<PyReadOnlyWorksheet> {
        let info = self
            .sheet_infos
            .iter()
            .find(|s| s.name == title)
            .ok_or_else(|| PyKeyError::new_err(format!("Worksheet {} does not exist.", title)))?;

        Ok(PyReadOnlyWorksheet {
            source: self.source.clone(),
            title: info.name.clone(),
            target_path: info.target_path.clone(),
            shared_strings: Arc::clone(&self.shared_strings),
            styles: Arc::clone(&self.styles),
            min_row: info.min_row,
            max_row: info.max_row,
            min_col: info.min_col,
            max_col: info.max_col,
            data_only: self.data_only,
        })
    }

    pub fn __contains__(&self, title: &str) -> bool {
        self.sheet_infos.iter().any(|s| s.name == title)
    }

    pub fn __len__(&self) -> usize {
        self.sheet_infos.len()
    }

    pub fn __iter__(&self) -> PyReadOnlyWorkbookIter {
        PyReadOnlyWorkbookIter {
            source: self.source.clone(),
            sheet_infos: self.sheet_infos.clone(),
            shared_strings: Arc::clone(&self.shared_strings),
            styles: Arc::clone(&self.styles),
            data_only: self.data_only,
            index: 0,
        }
    }

    pub fn close(&mut self) -> PyResult<()> {
        self.is_closed = true;
        Ok(())
    }

    pub fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (_exc_type=None, _exc_val=None, _exc_tb=None))]
    pub fn __exit__(
        &mut self,
        _exc_type: Option<&Bound<'_, PyAny>>,
        _exc_val: Option<&Bound<'_, PyAny>>,
        _exc_tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }

    pub fn __repr__(&self) -> String {
        format!("<ReadOnlyWorkbook sheetnames={:?}>", self.sheetnames())
    }
}

#[pyclass]
pub struct PyReadOnlyWorkbookIter {
    source: WorkbookSource,
    sheet_infos: Vec<SheetInfo>,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    data_only: bool,
    index: usize,
}

#[pymethods]
impl PyReadOnlyWorkbookIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<PyReadOnlyWorksheet> {
        if self.index < self.sheet_infos.len() {
            let info = &self.sheet_infos[self.index];
            self.index += 1;
            Some(PyReadOnlyWorksheet {
                source: self.source.clone(),
                title: info.name.clone(),
                target_path: info.target_path.clone(),
                shared_strings: Arc::clone(&self.shared_strings),
                styles: Arc::clone(&self.styles),
                min_row: info.min_row,
                max_row: info.max_row,
                min_col: info.min_col,
                max_col: info.max_col,
                data_only: self.data_only,
            })
        } else {
            None
        }
    }
}

// -------------------------------------------------------------
// Public load_workbook binding
// -------------------------------------------------------------

#[pyfunction]
#[pyo3(signature = (path, read_only=false, data_only=false))]
pub fn load_workbook<'py>(
    py: Python<'py>,
    path: &Bound<'_, PyAny>,
    read_only: bool,
    data_only: bool,
) -> PyResult<Bound<'py, PyAny>> {
    let source = workbook_source(path)?;
    let mut package = py
        .allow_threads(|| XlsxPackage::from_reader(source.open()?, read_only, data_only))
        .map_err(map_rpyxl_err)?;

    if read_only {
        let ro_wb = PyReadOnlyWorkbook {
            source,
            sheet_infos: package.sheets,
            shared_strings: package.shared_strings,
            styles: package.styles,
            active_index: package.active_sheet_index,
            data_only,
            is_closed: false,
        };
        Ok(Bound::new(py, ro_wb)?.into_any())
    } else {
        let wb = py
            .allow_threads(move || package.to_workbook())
            .map_err(map_rpyxl_err)?;
        let py_wb = PyWorkbook {
            wb: Arc::new(RwLock::new(wb)),
        };
        Ok(Bound::new(py, py_wb)?.into_any())
    }
}

// -----------------------------------------------------------------------------
// Write-Only Streaming Implementation
// -----------------------------------------------------------------------------

#[pyclass(name = "WriteOnlyWorksheet")]
pub struct PyWriteOnlyWorksheet {
    sheet: Arc<RwLock<WriteOnlySheet>>,
}

#[pymethods]
impl PyWriteOnlyWorksheet {
    #[getter]
    pub fn title(&self) -> String {
        self.sheet.read().title().to_string()
    }

    #[setter]
    pub fn set_title(&self, title: String) {
        self.sheet.write().set_title(&title);
    }

    pub fn append(&self, iterable: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut values = Vec::new();
        for item in iterable.iter()? {
            let item_val = item?;
            values.push(python_value_to_cell_value(&item_val)?);
        }

        let mut sheet = self.sheet.write();
        sheet.append_row(&values).map_err(map_rpyxl_err)
    }

    pub fn close(&self) -> PyResult<()> {
        Ok(())
    }

    pub fn __getitem__(&self, _key: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyTypeError::new_err(
            "'WriteOnlyWorksheet' object is not subscriptable",
        ))
    }

    #[pyo3(signature = (_row=1, _column=1, _value=None))]
    pub fn cell(&self, _row: u32, _column: u32, _value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        Err(PyAttributeError::new_err(
            "'WriteOnlyWorksheet' object has no attribute 'cell'",
        ))
    }

    pub fn iter_rows(&self) -> PyResult<()> {
        Err(PyAttributeError::new_err(
            "'WriteOnlyWorksheet' object has no attribute 'iter_rows'",
        ))
    }

    pub fn __repr__(&self) -> String {
        format!("<WriteOnlyWorksheet {:?}>", self.sheet.read().title())
    }
}

#[pyclass(name = "WriteOnlyWorkbook")]
pub struct PyWriteOnlyWorkbook {
    sheets: Vec<Arc<RwLock<WriteOnlySheet>>>,
    closed: bool,
}

#[pymethods]
impl PyWriteOnlyWorkbook {
    #[new]
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            closed: false,
        }
    }

    #[getter]
    pub fn write_only(&self) -> bool {
        true
    }

    #[getter]
    pub fn read_only(&self) -> bool {
        false
    }

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheets
            .iter()
            .map(|s| s.read().title().to_string())
            .collect()
    }

    #[getter]
    pub fn worksheets(&self) -> Vec<PyWriteOnlyWorksheet> {
        self.sheets
            .iter()
            .map(|s| PyWriteOnlyWorksheet {
                sheet: Arc::clone(s),
            })
            .collect()
    }

    #[pyo3(signature = (title=None, index=None))]
    pub fn create_sheet(
        &mut self,
        title: Option<String>,
        index: Option<usize>,
    ) -> PyResult<PyWriteOnlyWorksheet> {
        if self.closed {
            return Err(PyValueError::new_err("Workbook is closed"));
        }

        let sheet_title = title.unwrap_or_else(|| format!("Sheet{}", self.sheets.len() + 1));

        let sheet = WriteOnlySheet::new(&sheet_title).map_err(map_rpyxl_err)?;
        let sheet_arc = Arc::new(RwLock::new(sheet));

        if let Some(idx) = index {
            if idx < self.sheets.len() {
                self.sheets.insert(idx, Arc::clone(&sheet_arc));
            } else {
                self.sheets.push(Arc::clone(&sheet_arc));
            }
        } else {
            self.sheets.push(Arc::clone(&sheet_arc));
        }

        Ok(PyWriteOnlyWorksheet { sheet: sheet_arc })
    }

    pub fn __getitem__(&self, key: &Bound<'_, PyAny>) -> PyResult<PyWriteOnlyWorksheet> {
        if let Ok(name) = key.extract::<String>() {
            for sheet in &self.sheets {
                if sheet.read().title() == name {
                    return Ok(PyWriteOnlyWorksheet {
                        sheet: Arc::clone(sheet),
                    });
                }
            }
            Err(PyKeyError::new_err(format!(
                "Worksheet {} does not exist.",
                name
            )))
        } else if let Ok(idx) = key.extract::<isize>() {
            let len = self.sheets.len() as isize;
            let norm_idx = if idx < 0 { len + idx } else { idx };
            if norm_idx >= 0 && (norm_idx as usize) < self.sheets.len() {
                Ok(PyWriteOnlyWorksheet {
                    sheet: Arc::clone(&self.sheets[norm_idx as usize]),
                })
            } else {
                Err(PyIndexError::new_err("Worksheet index out of range"))
            }
        } else {
            Err(PyTypeError::new_err("Sheet key must be str or int"))
        }
    }

    pub fn index(&self, worksheet: &Bound<'_, PyWriteOnlyWorksheet>) -> PyResult<usize> {
        let title = worksheet.borrow().title();
        self.sheets
            .iter()
            .position(|s| s.read().title() == title)
            .ok_or_else(|| {
                PyValueError::new_err(format!("Worksheet '{}' is not in workbook", title))
            })
    }

    pub fn save(&mut self, path: &Bound<'_, PyAny>) -> PyResult<()> {
        if self.closed {
            return Err(PyValueError::new_err("Workbook is closed"));
        }
        if self.sheets.is_empty() {
            return Err(PyValueError::new_err(
                "Cannot save empty write-only workbook",
            ));
        }

        let py = path.py();
        let destination = filesystem_path(path)?;
        if destination.is_none() {
            check_binary_writer(path)?;
        }
        let sheets = self.sheets.clone();
        let bytes = py
            .allow_threads(move || {
                // Acquire and drop every native lock while detached from Python.
                let mut sheet_guards: Vec<_> = sheets.iter().map(|s| s.write()).collect();
                let mut refs: Vec<&mut WriteOnlySheet> =
                    sheet_guards.iter_mut().map(|g| &mut **g).collect();
                if let Some(destination) = destination {
                    save_write_only_workbook(destination, &mut refs)?;
                    Ok::<_, RpyxlError>(None)
                } else {
                    let mut buffer = Cursor::new(Vec::new());
                    write_write_only_workbook(&mut buffer, &mut refs)?;
                    Ok(Some(buffer.into_inner()))
                }
            })
            .map_err(map_rpyxl_err)?;
        if let Some(bytes) = bytes {
            write_binary_output(path, &bytes)?;
        }
        Ok(())
    }

    pub fn close(&mut self) -> PyResult<()> {
        self.closed = true;
        self.sheets.clear();
        Ok(())
    }

    pub fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (_exc_type=None, _exc_val=None, _exc_tb=None))]
    pub fn __exit__(
        &mut self,
        _exc_type: Option<&Bound<'_, PyAny>>,
        _exc_val: Option<&Bound<'_, PyAny>>,
        _exc_tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }

    pub fn __repr__(&self) -> String {
        let titles: Vec<String> = self
            .sheets
            .iter()
            .map(|s| s.read().title().to_string())
            .collect();
        format!("<WriteOnlyWorkbook sheetnames={:?}>", titles)
    }
}

impl Default for PyWriteOnlyWorkbook {
    fn default() -> Self {
        Self::new()
    }
}
