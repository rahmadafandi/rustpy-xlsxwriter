//! Formula columns and the totals row.
//!
//! Both write Excel formulas derived from the data's shape rather than from
//! its values: a computed column appends one formula per row, a totals row
//! appends one per column beneath it. They share the aggregate vocabulary and
//! the structural check, so they live together.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{Format, Worksheet};

use crate::worksheet::xlsx_err;

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

/// A computed column: a header and a formula template appended after the data
/// columns. `{row}` is replaced with the current row's 1-based sheet row and
/// `{first}` with the first data row.
///
/// There is deliberately no `{last}`: these cells are written while the data is
/// still streaming, so the final row is not yet known. Use `totals_row` for
/// anything spanning the whole column — that runs after the last row.
///
/// The formula text is passed to `rust_xlsxwriter` verbatim, so anything Excel
/// accepts works — nested calls, `SUMIFS`, cross-sheet references, and the 161
/// functions the crate rewrites with an `_xlfn.` prefix or as a dynamic array.
/// Only the structure is checked, by [`formula_problem`]; function names and
/// semantics are Excel's business.
pub struct FormulaColumn {
    pub header: String,
    pub template: String,
}

impl FormulaColumn {
    /// Substitute the row placeholders for one data row.
    pub fn render(&self, row: u32, first: u32) -> String {
        let mut out = self.template.clone();
        if out.contains("{row}") {
            out = out.replace("{row}", &row.to_string());
        }
        if out.contains("{first}") {
            out = out.replace("{first}", &first.to_string());
        }
        out
    }
}

/// Structural check on a formula: balanced parentheses and quotes, and some
/// content after the `=`.
///
/// Deliberately *not* a function-name check. A malformed formula does not
/// corrupt the file — every case tested opens fine and shows `#NAME?`/`#VALUE!`
/// in the cell — so the only thing validation buys is catching a typo earlier.
/// That makes a false positive strictly worse than a false negative: rejecting
/// a formula Excel would have accepted blocks real work, while letting one
/// through costs an error value in one cell. Function names cannot be checked
/// safely — `LAMBDA`/`LET` bind their own names, workbooks carry user-defined
/// functions, and Excel keeps adding to the list.
///
/// Returns the problem description, or `None` when the formula looks sound.
pub fn formula_problem(formula: &str) -> Option<String> {
    let body = formula.strip_prefix('=').unwrap_or(formula);
    if body.trim().is_empty() {
        return Some("it is empty".to_string());
    }

    let mut depth: i32 = 0;
    // Excel quotes strings with `"` (a literal quote inside is doubled) and
    // wraps sheet names containing spaces or punctuation in `'`. Parentheses
    // inside either are text, not structure — `='Sheet (1)'!A1` is valid.
    let mut in_double = false;
    let mut in_single = false;
    let chars: Vec<char> = body.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '"' if !in_single => {
                if in_double && chars.get(i + 1) == Some(&'"') {
                    i += 1; // an escaped quote inside a string
                } else {
                    in_double = !in_double;
                }
            }
            '\'' if !in_double => in_single = !in_single,
            '(' if !in_double && !in_single => depth += 1,
            ')' if !in_double && !in_single => {
                depth -= 1;
                if depth < 0 {
                    return Some("it closes a parenthesis that was never opened".to_string());
                }
            }
            _ => {}
        }
        i += 1;
    }

    if in_double {
        return Some("a double quote is left open".to_string());
    }
    if in_single {
        return Some("a single quote is left open".to_string());
    }
    if depth > 0 {
        return Some(format!(
            "{depth} parenthesis{} left unclosed",
            if depth == 1 { " is" } else { "es are" }
        ));
    }
    None
}

/// Read `formula_columns` — an ordered `{header: formula}` mapping.
pub fn resolve_formula_columns(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Vec<FormulaColumn>> {
    let Some(spec) = spec else {
        return Ok(Vec::new());
    };
    let dict = spec
        .cast::<PyDict>()
        .map_err(|_| value_err("formula_columns must be a dict of {header: formula}".into()))?;
    let mut out = Vec::with_capacity(dict.len());
    for (key, val) in dict.iter() {
        let header: String = key
            .extract()
            .map_err(|_| value_err("formula_columns keys must be header names".into()))?;
        let template: String = val.extract().map_err(|_| {
            value_err(format!(
                "formula_columns['{header}'] must be a formula string"
            ))
        })?;
        if template.trim().is_empty() {
            return Err(value_err(format!("formula_columns['{header}'] is empty")));
        }
        // Placeholders expand to digits, so the structure is already final.
        if let Some(problem) = formula_problem(&template) {
            return Err(value_err(format!(
                "formula_columns['{header}'] looks malformed: {problem}. \
Formula: {template}"
            )));
        }
        if template.contains("{last}") {
            return Err(value_err(format!(
                "formula_columns['{header}'] uses {{last}}, but the last data row is not \
known while rows are still streaming. Use totals_row for a whole-column formula, \
or {{row}} for a per-row one."
            )));
        }
        out.push(FormulaColumn { header, template });
    }
    Ok(out)
}

/// Write the formula cells for one data row, starting at `first_col`.
pub fn write_formula_row(
    worksheet: &mut Worksheet,
    columns: &[FormulaColumn],
    first_col: u16,
    row: u32,
    first_data_row: u32,
    fmt: Option<&Format>,
) -> PyResult<()> {
    for (offset, column) in columns.iter().enumerate() {
        // A1 notation is 1-based.
        let formula = column.render(row + 1, first_data_row + 1);
        let col = first_col + offset as u16;
        match fmt {
            Some(f) => worksheet
                .write_formula_with_format(row, col, formula.as_str(), f)
                .map_err(xlsx_err)?,
            None => worksheet
                .write_formula(row, col, formula.as_str())
                .map_err(xlsx_err)?,
        };
    }
    Ok(())
}

/// What to write in one totals cell.
pub enum TotalsCell {
    /// A named aggregate, expanded to `=FUNC(range)` over the column.
    Aggregate(&'static str),
    /// A caller-supplied formula. `{col}` is the column letter, `{first}` and
    /// `{last}` the first and last data rows — all known by the time the totals
    /// row is written.
    Formula(String),
}

/// Map an aggregate name to its Excel function.
pub fn excel_function(name: &str) -> Option<&'static str> {
    match name.to_ascii_lowercase().as_str() {
        "sum" => Some("SUM"),
        "average" | "avg" | "mean" => Some("AVERAGE"),
        "count" => Some("COUNT"),
        "min" => Some("MIN"),
        "max" => Some("MAX"),
        "product" => Some("PRODUCT"),
        "stdev" => Some("STDEV"),
        _ => None,
    }
}
