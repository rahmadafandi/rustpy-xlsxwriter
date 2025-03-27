use pyo3::prelude::*;

/// Validates that the sheet name meets Excel's requirements:
/// - Must be <= 31 characters
/// - Cannot contain [ ] : * ? / \
/// Returns true if valid, false if invalid
#[pyfunction]
pub fn validate_sheet_name(name: &str) -> bool {
    if name.len() > 31 {
        return false;
    }
    !name.contains(&['[', ']', ':', '*', '?', '/', '\\'][..])
}