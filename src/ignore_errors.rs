//! Suppressing Excel's green error triangles, per column.
//!
//! The one that matters in practice is `number_stored_as_text`: an ID, SKU or
//! postcode column is digits stored as text on purpose, and Excel flags every
//! cell of it. Applied after the data, like the conditional formats, because
//! the range depends on how many rows there turned out to be.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{IgnoreError, Worksheet};

use crate::options::{column_index, value_err};
use crate::worksheet::xlsx_err;

const NAMES: [&str; 9] = [
    "number_stored_as_text",
    "formula_error",
    "formula_differs",
    "formula_refers_to_empty_cells",
    "formula_omits_cells",
    "data_validation_error",
    "two_digit_text_year",
    "unlocked_cells_with_formula",
    "inconsistent_column_formula",
];

fn error_of(name: &str) -> PyResult<IgnoreError> {
    Ok(match name {
        "number_stored_as_text" => IgnoreError::NumberStoredAsText,
        "formula_error" => IgnoreError::FormulaError,
        "formula_differs" => IgnoreError::FormulaDiffers,
        "formula_refers_to_empty_cells" => IgnoreError::FormulaRefersToEmptyCells,
        "formula_omits_cells" => IgnoreError::FormulaOmitsCells,
        "data_validation_error" => IgnoreError::DataValidationError,
        "two_digit_text_year" => IgnoreError::TwoDigitTextYear,
        "unlocked_cells_with_formula" => IgnoreError::UnlockedCellsWithFormula,
        "inconsistent_column_formula" => IgnoreError::InconsistentColumnFormula,
        other => {
            return Err(value_err(format!(
                "ignore_errors: unknown error '{other}' (expected one of {})",
                NAMES.join(", ")
            )))
        }
    })
}

/// `(column, error)` pairs.
#[derive(Default)]
pub struct IgnoreErrors(Vec<(String, IgnoreError)>);

impl IgnoreErrors {
    /// Accepts a list of column names — which means `number_stored_as_text`,
    /// the reason anyone reaches for this — or a dict mapping a column to one
    /// error name.
    ///
    /// One name, not a list: Excel allows a single ignore rule per cell, so a
    /// second one on the same column is rejected by the file format itself.
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(IgnoreErrors(out));
        };

        if let Ok(map) = spec.cast::<PyDict>() {
            for (key, value) in map.iter() {
                let column: String = key.extract()?;
                let name: String = value.extract().map_err(|_| {
                    value_err(format!(
                        "ignore_errors: '{column}' must be a single error name — \
                         Excel allows one ignore rule per cell, so a column \
                         cannot carry two"
                    ))
                })?;
                out.push((column.clone(), error_of(&name)?));
            }
        } else {
            let columns: Vec<String> = spec.extract().map_err(|_| {
                value_err(
                    "ignore_errors must be a list of column names or a dict of \
                     {column: error}"
                        .to_string(),
                )
            })?;
            for column in columns {
                out.push((column, IgnoreError::NumberStoredAsText));
            }
        }
        Ok(IgnoreErrors(out))
    }

    /// An unknown column warns and is skipped, like the other column-keyed
    /// options: a triangle left showing is not worth failing an export over.
    pub fn apply(
        &self,
        worksheet: &mut Worksheet,
        headers: &[String],
        header_row: u32,
        data_rows: u32,
        py: Python,
    ) -> PyResult<()> {
        if self.0.is_empty() || data_rows == 0 {
            return Ok(());
        }
        let first = header_row + 1;
        let last = header_row + data_rows;

        for (column, error) in &self.0 {
            let Some(idx) = column_index(headers, column, "ignore_errors", py)? else {
                continue;
            };
            let col = idx as u16;
            worksheet
                .ignore_error_range(first, col, last, col, *error)
                .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
