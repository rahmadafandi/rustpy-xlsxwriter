mod data_types;
mod metadata;
mod utils;
mod worksheet;

use pyo3::prelude::*;

/// A Python module implemented in Rust.
#[pymodule]
fn rustpy_xlsxwriter(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(worksheet::write_worksheet, m)?)?;
    m.add_function(wrap_pyfunction!(worksheet::write_worksheets, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_version, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_name, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_authors, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_description, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_repository, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_homepage, m)?)?;
    m.add_function(wrap_pyfunction!(metadata::get_license, m)?)?;
    m.add_function(wrap_pyfunction!(utils::validate_sheet_name, m)?)?;
    Ok(())
}
