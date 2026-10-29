#![allow(clippy::useless_conversion)]

use pyo3::prelude::*;

mod bindings;

pub use bindings::{
    load_workbook, PyCell, PyReadOnlyCell, PyReadOnlyWorkbook, PyReadOnlyWorksheet, PyRowIterator,
    PyWorkbook, PyWorksheet, PyWriteOnlyWorkbook, PyWriteOnlyWorksheet,
};

#[pymodule]
fn _core(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCell>()?;
    m.add_class::<PyReadOnlyWorkbook>()?;
    m.add_class::<PyReadOnlyWorksheet>()?;
    m.add_class::<PyReadOnlyCell>()?;
    m.add_class::<PyRowIterator>()?;
    m.add_class::<PyWriteOnlyWorkbook>()?;
    m.add_class::<PyWriteOnlyWorksheet>()?;
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    Ok(())
}
