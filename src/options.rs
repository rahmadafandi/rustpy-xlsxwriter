//! The vocabulary every option mapping shares.
//!
//! Each feature parses its own mapping — `page_setup`, `charts`,
//! `data_validations` and the rest — but they all do the same three things:
//! raise a `ValueError` the same way, read an optional key out of a dict,
//! reject the ones they do not know, and
//! turn a column name into an index or warn that it is not there. Written out
//! per module that was about 120 lines of identical code, and twelve places to
//! forget when any of it changed.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict, PyDictMethods};

/// The `ValueError` every option parser raises.
pub fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

/// Read an optional key, extracted to `T`.
///
/// The higher-ranked bounds are what make this generic at all: `extract`
/// borrows from the `Bound` it is called on, and its error type has to be
/// convertible for `?` to work across every `T`.
pub fn opt<'py, T>(map: &Bound<'py, PyDict>, key: &str) -> PyResult<Option<T>>
where
    T: for<'a> FromPyObject<'a, 'py>,
    for<'a> <T as FromPyObject<'a, 'py>>::Error: Into<PyErr>,
{
    match map.get_item(key)? {
        Some(v) => Ok(Some(v.extract().map_err(Into::into)?)),
        None => Ok(None),
    }
}

/// Reject any key the caller's mapping does not recognise.
///
/// A key comes from the programmer, so a misspelled one that quietly does
/// nothing gives no other signal — unlike a column name, which comes from
/// data that varies and only warns.
pub fn reject_unknown_keys(map: &Bound<'_, PyDict>, what: &str, known: &[&str]) -> PyResult<()> {
    for key in map.keys().iter() {
        let name: String = key.extract()?;
        if !known.contains(&name.as_str()) {
            return Err(value_err(format!(
                "{what}: unknown key '{name}' (expected one of {})",
                known.join(", ")
            )));
        }
    }
    Ok(())
}

/// Position of `column` in `headers`, or `None` after warning that it is
/// missing.
///
/// Every column-keyed option warns rather than raising here: a rule that
/// cannot be placed costs some shading or a link, not the export.
pub fn column_index(
    headers: &[String],
    column: &str,
    what: &str,
    py: Python,
) -> PyResult<Option<usize>> {
    match headers.iter().position(|h| h == column) {
        Some(idx) => Ok(Some(idx)),
        None => {
            py.import("warnings")?.call_method1(
                "warn",
                (format!("{what}: unknown column '{column}', skipped"),),
            )?;
            Ok(None)
        }
    }
}
