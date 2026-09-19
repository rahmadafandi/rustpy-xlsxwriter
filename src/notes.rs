//! Cell notes on header cells, parsed from a `notes` mapping.
//!
//! Notes are keyed by column name and land on that column's header, which is
//! what makes them useful here: a note is where you explain what a column
//! means without widening it or adding a legend sheet. Per-row notes would
//! need row indices and a different shape, and nothing has asked for them.
//!
//! Unlike a row group, a note survives constant-memory mode. It is stored
//! beside the cell data because it feeds the `<row>` span attributes, and
//! that writer omits spans anyway — the note itself, its anchor and the
//! legacy drawing all come out identical either way, so no sheet is taken out
//! of constant memory for one.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{Note, Worksheet};

use crate::format::parse_color;
use crate::worksheet::xlsx_err;

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

const KEYS: [&str; 6] = [
    "text",
    "author",
    "width",
    "height",
    "visible",
    "background_color",
];

/// `(column, note)` pairs.
#[derive(Default)]
pub struct Notes(Vec<(String, Note)>);

fn build(value: &Bound<'_, PyAny>, column: &str) -> PyResult<Note> {
    // The short form is just the text, which is what almost every caller
    // wants; the dict form is there for the rest.
    if let Ok(text) = value.extract::<String>() {
        return Ok(Note::new(text));
    }
    let map = value.cast::<PyDict>().map_err(|_| {
        value_err(format!(
            "notes: '{column}' must be a string or a dict with 'text'"
        ))
    })?;
    for key in map.keys().iter() {
        let name: String = key.extract()?;
        if !KEYS.contains(&name.as_str()) {
            return Err(value_err(format!(
                "notes: unknown key '{name}' for '{column}' (expected one of {})",
                KEYS.join(", ")
            )));
        }
    }
    let text: String = map
        .get_item("text")?
        .ok_or_else(|| value_err(format!("notes: '{column}' needs 'text'")))?
        .extract()?;

    let mut note = Note::new(text);
    if let Some(v) = map.get_item("author")? {
        note = note.set_author(v.extract::<String>()?);
    }
    if let Some(v) = map.get_item("width")? {
        note = note.set_width(v.extract()?);
    }
    if let Some(v) = map.get_item("height")? {
        note = note.set_height(v.extract()?);
    }
    if let Some(v) = map.get_item("visible")? {
        note = note.set_visible(v.extract()?);
    }
    if let Some(v) = map.get_item("background_color")? {
        note = note.set_background_color(parse_color(&v.extract::<String>()?)?);
    }
    Ok(note)
}

impl Notes {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(Notes(out));
        };
        let map = spec
            .cast::<PyDict>()
            .map_err(|_| value_err("notes must be a dict keyed by column name".to_string()))?;

        for (key, value) in map.iter() {
            let column: String = key.extract()?;
            let note = build(&value, &column)?;
            out.push((column, note));
        }
        Ok(Notes(out))
    }

    /// An unknown column warns and is skipped, like the other column-keyed
    /// options.
    pub fn apply(
        &self,
        worksheet: &mut Worksheet,
        headers: &[String],
        header_row: u32,
        py: Python,
    ) -> PyResult<()> {
        if self.0.is_empty() {
            return Ok(());
        }
        let warnings = py.import("warnings")?;
        for (column, note) in &self.0 {
            let Some(idx) = headers.iter().position(|h| h == column) else {
                warnings.call_method1(
                    "warn",
                    (format!("notes: unknown column '{column}', skipped"),),
                )?;
                continue;
            };
            worksheet
                .insert_note(header_row, idx as u16, note)
                .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
