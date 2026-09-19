//! Outline grouping — the collapsible +/- brackets beside rows and columns.
//!
//! Applied at two different points, which the split methods below reflect.
//! Row groups set row options, so they go in before the first data row like
//! the heights do: constant-memory mode cannot revisit a flushed row. Column
//! groups are named by header, and the headers are only known once the first
//! record has been read, so those go in after the data with the other
//! column-keyed options.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::Worksheet;

use crate::worksheet::xlsx_err;

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

const KEYS: [&str; 4] = ["rows", "columns", "symbols_above", "symbols_to_left"];

#[derive(Default)]
pub struct Outline {
    /// `(first_row, last_row, collapsed)`, by sheet row index.
    rows: Vec<(u32, u32, bool)>,
    /// `(first_header, last_header, collapsed)`, by column name.
    columns: Vec<(String, String, bool)>,
    symbols_above: Option<bool>,
    symbols_to_left: Option<bool>,
}

/// One group entry: `{"from": ..., "to": ..., "collapsed": bool}`.
///
/// The bounds come back unextracted — rows are indices and columns are header
/// names, so each caller narrows them itself. A generic here would need its
/// extraction error to convert to `PyErr`, which is more bound than the two
/// call sites are worth.
fn group_parts<'py>(
    item: &Bound<'py, PyAny>,
    what: &str,
) -> PyResult<(Bound<'py, PyAny>, Bound<'py, PyAny>, bool)> {
    let map = item.cast::<PyDict>().map_err(|_| {
        value_err(format!(
            "outline: every '{what}' group must be a dict with 'from' and 'to'"
        ))
    })?;
    for key in map.keys().iter() {
        let name: String = key.extract()?;
        if !["from", "to", "collapsed"].contains(&name.as_str()) {
            return Err(value_err(format!(
                "outline: unknown key '{name}' in a '{what}' group \
                 (expected from, to, collapsed)"
            )));
        }
    }
    let first = map
        .get_item("from")?
        .ok_or_else(|| value_err(format!("outline: a '{what}' group needs 'from'")))?;
    let last = map
        .get_item("to")?
        .ok_or_else(|| value_err(format!("outline: a '{what}' group needs 'to'")))?;
    let collapsed = match map.get_item("collapsed")? {
        Some(v) => v.extract()?,
        None => false,
    };
    Ok((first, last, collapsed))
}

impl Outline {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut outline = Outline::default();
        let Some(spec) = spec else {
            return Ok(outline);
        };
        let map = spec
            .cast::<PyDict>()
            .map_err(|_| value_err("outline must be a dict".to_string()))?;

        for (key, value) in map.iter() {
            let name: String = key.extract()?;
            match name.as_str() {
                "rows" => {
                    for item in value.try_iter()? {
                        let (first, last, collapsed) = group_parts(&item?, "rows")?;
                        outline
                            .rows
                            .push((first.extract()?, last.extract()?, collapsed));
                    }
                }
                "columns" => {
                    for item in value.try_iter()? {
                        let (first, last, collapsed) = group_parts(&item?, "columns")?;
                        outline
                            .columns
                            .push((first.extract()?, last.extract()?, collapsed));
                    }
                }
                "symbols_above" => outline.symbols_above = Some(value.extract()?),
                "symbols_to_left" => outline.symbols_to_left = Some(value.extract()?),
                _ => {
                    return Err(value_err(format!(
                        "outline: unknown key '{name}' (expected one of {})",
                        KEYS.join(", ")
                    )))
                }
            }
        }

        for (first, last, _) in &outline.rows {
            if first > last {
                return Err(value_err(format!(
                    "outline: row group 'from' ({first}) is after 'to' ({last})"
                )));
            }
        }
        Ok(outline)
    }

    /// `true` when any row group was asked for.
    ///
    /// Callers use this to keep the sheet out of constant-memory mode:
    /// `write_constant_table_row` writes `hidden` but not `outlineLevel`, so a
    /// collapsed group there produces hidden rows with no bracket to reopen
    /// them — worse than not offering the feature.
    pub fn needs_buffered_rows(&self) -> bool {
        !self.rows.is_empty()
    }

    /// Row groups and the symbol positions, before the first data row.
    pub fn apply_rows(&self, worksheet: &mut Worksheet) -> PyResult<()> {
        if let Some(on) = self.symbols_above {
            worksheet.group_symbols_above(on);
        }
        if let Some(on) = self.symbols_to_left {
            worksheet.group_symbols_to_left(on);
        }
        for (first, last, collapsed) in &self.rows {
            if *collapsed {
                worksheet.group_rows_collapsed(*first, *last)
            } else {
                worksheet.group_rows(*first, *last)
            }
            .map_err(xlsx_err)?;
        }
        Ok(())
    }

    /// Column groups, once the headers that name them are known.
    ///
    /// An unknown name warns and is skipped, like the other column-keyed
    /// options — a missing bracket is not worth failing an export over.
    pub fn apply_columns(
        &self,
        worksheet: &mut Worksheet,
        headers: &[String],
        py: Python,
    ) -> PyResult<()> {
        if self.columns.is_empty() {
            return Ok(());
        }
        let warnings = py.import("warnings")?;
        for (from, to, collapsed) in &self.columns {
            let (Some(first), Some(last)) = (
                headers.iter().position(|h| h == from),
                headers.iter().position(|h| h == to),
            ) else {
                warnings.call_method1(
                    "warn",
                    (format!(
                        "outline: unknown column in group '{from}'..'{to}', skipped"
                    ),),
                )?;
                continue;
            };
            if first > last {
                return Err(value_err(format!(
                    "outline: column group '{from}' comes after '{to}' in the data"
                )));
            }
            if *collapsed {
                worksheet.group_columns_collapsed(first as u16, last as u16)
            } else {
                worksheet.group_columns(first as u16, last as u16)
            }
            .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
