use pyo3::prelude::*;

/// Returns the version of the library
#[pyfunction]
pub fn get_version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

/// Returns the name of the library
#[pyfunction]
pub fn get_name() -> String {
    env!("CARGO_PKG_NAME").to_string()
}

/// Returns the authors of the library
#[pyfunction]
pub fn get_authors() -> String {
    env!("CARGO_PKG_AUTHORS").to_string()
}

/// Returns the description of the library
#[pyfunction]
pub fn get_description() -> String {
    env!("CARGO_PKG_DESCRIPTION").to_string()
}

/// Returns the repository URL of the library
#[pyfunction]
pub fn get_repository() -> String {
    env!("CARGO_PKG_REPOSITORY").to_string()
}

/// Returns the homepage URL of the library
#[pyfunction]
pub fn get_homepage() -> String {
    env!("CARGO_PKG_HOMEPAGE").to_string()
}

/// Returns the license of the library
#[pyfunction]
pub fn get_license() -> String {
    env!("CARGO_PKG_LICENSE").to_string()
}