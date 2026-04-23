use pyo3::prelude::*;

/// Validates that the sheet name meets Excel's requirements:
/// - Must be 1..=31 characters
/// - Cannot contain [ ] : * ? / \
#[pyfunction]
pub fn validate_sheet_name(name: &str) -> bool {
    if name.is_empty() || name.chars().count() > 31 {
        return false;
    }
    !name.contains(&['[', ']', ':', '*', '?', '/', '\\'][..])
}

/// Validate a sheet name or return a `PyValueError` with a consistent message.
pub fn ensure_valid_sheet_name(name: &str) -> PyResult<()> {
    if validate_sheet_name(name) {
        Ok(())
    } else {
        Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\",
            name
        )))
    }
}
