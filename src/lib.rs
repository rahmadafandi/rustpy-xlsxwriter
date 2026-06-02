mod arrow_ffi;
mod arrow_writer;
mod csv_writer;
mod data_types;
mod format;
mod helpers;
mod utils;
mod worksheet;

use pyo3::prelude::*;

#[pymodule]
fn rustpy_xlsxwriter(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(worksheet::write_worksheet, m)?)?;
    m.add_function(wrap_pyfunction!(worksheet::write_worksheets, m)?)?;
    m.add_function(wrap_pyfunction!(utils::validate_sheet_name, m)?)?;
    m.add_function(wrap_pyfunction!(csv_writer::write_csv, m)?)?;
    m.add_class::<format::Format>()?;
    Ok(())
}
