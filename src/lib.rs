mod arrow_ffi;
mod arrow_writer;
mod cell;
mod charts;
mod conditional_format;
mod csv_writer;
mod data_types;
mod data_validation;
mod format;
mod helpers;
mod notes;
mod outline;
mod ignore_errors;
mod images;
mod page_setup;
mod sheet_view;
mod sparklines;
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
