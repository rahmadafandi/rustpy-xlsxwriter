//! Per-sheet settings, resolved once and applied around the data.
//!
//! [`SheetLayout`] is the bag every write path carries: what the caller asked
//! for, already parsed and validated, so the row loops never touch a Python
//! object. Its parts land at two different moments — merges, heights, page
//! setup and images before the first data row, because constant-memory mode
//! cannot revisit a flushed one; the autofilter and totals row after it,
//! because their ranges depend on how many rows there turned out to be.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict, PyList};
use rust_xlsxwriter::{Format, Worksheet};

use crate::formula::{excel_function, formula_problem, TotalsCell};
use crate::helpers::NumReps;
use crate::worksheet::xlsx_err;

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

/// Row-level layout for one sheet, resolved from Python before any cell is
/// written.
///
/// Constant-memory mode flushes each row as soon as the next one starts and
/// cannot revisit it — a `merge_range` or `set_row_height` aimed at a row that
/// has already gone out is *silently* dropped (`rust_xlsxwriter` prints to
/// stderr and carries on). So every one of these is applied up front, in
/// [`SheetLayout::apply`], before the first data cell.
#[derive(Default)]
pub struct SheetLayout {
    /// Row index the header is written on. Data starts at `header_row + 1`.
    pub header_row: u32,
    /// `(first_row, first_col, last_row, last_col, value, format)`.
    pub merges: Vec<(u32, u16, u32, u16, String, Option<Format>)>,
    pub row_heights: Vec<(u32, f64)>,
    pub row_formats: Vec<(u32, Format)>,
    /// Background colour for alternating data rows, as given by the caller.
    pub band_color: Option<String>,
    /// Add a filter dropdown over the header row and its data. Unlike the rest
    /// of this struct this is applied *after* the data, in
    /// [`SheetLayout::apply_autofilter`] — the range needs the final row count,
    /// which is only known once the last row has been written.
    pub autofilter: bool,
    /// `(column name, aggregate or raw formula)` for the totals row. Applied
    /// after the data: the row sits below it and the ranges depend on how many
    /// rows there turned out to be.
    pub totals: Vec<(String, TotalsCell)>,
    pub totals_label: Option<String>,
    pub totals_format: Option<Format>,
    /// Text for missing values and the infinities. Not geometry like the rest of this
    /// struct, but it is resolved per sheet and every write path already
    /// carries the layout, which beats a parameter on four more functions.
    pub na_text: Option<String>,
    pub inf_text: Option<String>,
    /// Page and print setup; see [`crate::page_setup`].
    pub page: crate::page_setup::PageSetup,
    /// Screen presentation; see [`crate::sheet_view`].
    pub view: crate::sheet_view::SheetView,
    /// Per-column conditional formats. Applied after the data, like the
    /// autofilter, since the range depends on the final row count.
    pub conditional: crate::conditional_format::ConditionalFormats,
    /// Error indicators to suppress; see [`crate::ignore_errors`].
    pub ignore: crate::ignore_errors::IgnoreErrors,
    /// Per-column data validation; see [`crate::data_validation`].
    pub validations: crate::data_validation::DataValidations,
    /// Collapsible row/column groups; see [`crate::outline`].
    pub outline: crate::outline::Outline,
    /// Notes on header cells; see [`crate::notes`].
    pub notes: crate::notes::Notes,
    /// Images anchored to cells; see [`crate::images`].
    pub images: crate::images::Images,
    /// Per-row trend charts; see [`crate::sparklines`].
    pub sparklines: crate::sparklines::Sparklines,
    /// Charts anchored to a cell; see [`crate::charts`].
    pub charts: crate::charts::Charts,
}

impl SheetLayout {
    /// Borrow the non-finite representations for the write paths.
    pub fn reps(&self) -> NumReps<'_> {
        NumReps {
            na: self.na_text.as_deref(),
            inf: self.inf_text.as_deref(),
        }
    }
}

impl SheetLayout {
    /// First data row.
    pub fn first_data_row(&self) -> u32 {
        self.header_row + 1
    }

    /// True when `row` (an absolute sheet row) is a shaded band row. The first
    /// data row is left unshaded so banding starts on the second one.
    pub fn is_banded(&self, row: u32) -> bool {
        self.band_color.is_some() && (row - self.first_data_row()) % 2 == 1
    }

    /// Emit merges, row heights and row formats. Must run before data rows.
    pub fn apply(&self, worksheet: &mut Worksheet) -> PyResult<()> {
        self.page.apply(worksheet)?;
        self.view.apply(worksheet);
        self.outline.apply_rows(worksheet)?;
        self.images.apply(worksheet)?;
        for (r1, c1, r2, c2, value, fmt) in &self.merges {
            let blank = Format::new();
            worksheet
                .merge_range(*r1, *c1, *r2, *c2, value, fmt.as_ref().unwrap_or(&blank))
                .map_err(xlsx_err)?;
        }
        // Height before format: `set_row_format` on a row with no stored
        // options would otherwise reset the height back to the default.
        for (row, height) in &self.row_heights {
            worksheet.set_row_height(*row, *height).map_err(xlsx_err)?;
        }
        for (row, fmt) in &self.row_formats {
            worksheet.set_row_format(*row, fmt).map_err(xlsx_err)?;
        }
        Ok(())
    }

    /// Add the filter over `header_row..=header_row + data_rows`. Safe to call
    /// after every row is flushed: the `autoFilter` element lives in the
    /// worksheet footer, not in the row data.
    pub fn apply_autofilter(
        &self,
        worksheet: &mut Worksheet,
        data_rows: u32,
        num_columns: usize,
    ) -> PyResult<()> {
        if !self.autofilter || num_columns == 0 {
            return Ok(());
        }
        worksheet
            .autofilter(
                self.header_row,
                0,
                self.header_row + data_rows,
                (num_columns - 1) as u16,
            )
            .map_err(xlsx_err)?;
        Ok(())
    }

    /// Write the totals row below the data. Skipped when there is no data —
    /// a formula over an empty range (`=SUM(B2:B1)`) is not valid.
    pub fn apply_totals(
        &self,
        worksheet: &mut Worksheet,
        headers: &[String],
        data_rows: u32,
        py: Python,
    ) -> PyResult<()> {
        if (self.totals.is_empty() && self.totals_label.is_none()) || data_rows == 0 {
            return Ok(());
        }
        let row = self.first_data_row() + data_rows;
        // A1 notation is 1-based, and the data starts one row below the header.
        let first = self.first_data_row() + 1;
        let last = self.first_data_row() + data_rows;

        let warnings = py.import("warnings")?;
        let mut used_first_column = false;

        for (name, function) in &self.totals {
            let Some(col) = headers.iter().position(|h| h == name) else {
                warnings.call_method1(
                    "warn",
                    (format!("totals_row: unknown column '{name}', skipped"),),
                )?;
                continue;
            };
            if col == 0 {
                used_first_column = true;
            }
            let letter = rust_xlsxwriter::utility::column_number_to_name(col as u16);
            let formula = match function {
                TotalsCell::Aggregate(f) => format!("={f}({letter}{first}:{letter}{last})"),
                TotalsCell::Formula(t) => t
                    .replace("{col}", &letter)
                    .replace("{first}", &first.to_string())
                    .replace("{last}", &last.to_string()),
            };
            match &self.totals_format {
                Some(fmt) => worksheet
                    .write_formula_with_format(row, col as u16, formula.as_str(), fmt)
                    .map_err(xlsx_err)?,
                None => worksheet
                    .write_formula(row, col as u16, formula.as_str())
                    .map_err(xlsx_err)?,
            };
        }

        if let Some(label) = &self.totals_label {
            if used_first_column {
                return Err(value_err(
                    "totals_label would overwrite the totals formula in the first column; \
drop one of them or move the aggregate to another column"
                        .into(),
                ));
            }
            match &self.totals_format {
                Some(fmt) => worksheet
                    .write_string_with_format(row, 0, label, fmt)
                    .map_err(xlsx_err)?,
                None => worksheet.write_string(row, 0, label).map_err(xlsx_err)?,
            };
        }
        Ok(())
    }
}

/// Read a `{row_index: value}` mapping into a sorted `Vec<(u32, T)>`.
fn row_keyed<T>(
    spec: Option<&Bound<'_, PyAny>>,
    what: &str,
    mut convert: impl FnMut(&Bound<'_, PyAny>) -> PyResult<T>,
) -> PyResult<Vec<(u32, T)>> {
    let Some(spec) = spec else {
        return Ok(Vec::new());
    };
    let dict = spec
        .cast::<PyDict>()
        .map_err(|_| value_err(format!("{what} must be a dict keyed by row index")))?;
    let mut out = Vec::with_capacity(dict.len());
    for (key, val) in dict.iter() {
        let row: u32 = key
            .extract()
            .map_err(|_| value_err(format!("{what}: row index must be a non-negative int")))?;
        out.push((row, convert(&val)?));
    }
    out.sort_by_key(|(row, _)| *row);
    Ok(out)
}

/// Build a [`SheetLayout`] from the Python-side arguments, rejecting anything
/// constant-memory mode would otherwise drop without raising.
#[allow(clippy::too_many_arguments)]
pub fn resolve_layout(
    header_row: u32,
    merge_ranges: Option<&Bound<'_, PyAny>>,
    row_heights: Option<&Bound<'_, PyAny>>,
    row_formats: Option<&Bound<'_, PyAny>>,
    banded_rows: Option<String>,
    autofilter: bool,
    totals_row: Option<&Bound<'_, PyAny>>,
    totals_label: Option<String>,
    totals_format: Option<Format>,
    na_text: Option<String>,
    inf_text: Option<String>,
    page_setup: Option<&Bound<'_, PyAny>>,
    conditional_formats: Option<&Bound<'_, PyAny>>,
    sheet_view: Option<&Bound<'_, PyAny>>,
    ignore_errors: Option<&Bound<'_, PyAny>>,
    data_validations: Option<&Bound<'_, PyAny>>,
    outline: Option<&Bound<'_, PyAny>>,
    notes: Option<&Bound<'_, PyAny>>,
    images: Option<&Bound<'_, PyAny>>,
    sparklines: Option<&Bound<'_, PyAny>>,
    charts: Option<&Bound<'_, PyAny>>,
) -> PyResult<SheetLayout> {
    let mut totals = Vec::new();
    if let Some(spec) = totals_row {
        let dict = spec.cast::<PyDict>().map_err(|_| {
            value_err("totals_row must be a dict of {column name: aggregate}".into())
        })?;
        for (key, val) in dict.iter() {
            let column: String = key
                .extract()
                .map_err(|_| value_err("totals_row keys must be column names".into()))?;
            let name: String = val.extract().map_err(|_| {
                value_err("totals_row values must be aggregate names or formulas".into())
            })?;
            // A leading '=' marks a raw formula; anything else must name a
            // known aggregate, so a typo raises instead of writing a literal.
            let cell = if name.trim_start().starts_with('=') {
                if let Some(problem) = formula_problem(&name) {
                    return Err(value_err(format!(
                        "totals_row['{column}'] looks malformed: {problem}. Formula: {name}"
                    )));
                }
                TotalsCell::Formula(name)
            } else {
                TotalsCell::Aggregate(excel_function(&name).ok_or_else(|| {
                    value_err(format!(
                        "totals_row: unknown aggregate '{name}' for column '{column}' \
(valid: sum, average, count, min, max, product, stdev; or a formula starting with '=')"
                    ))
                })?)
            };
            totals.push((column, cell));
        }
    }

    let mut merges = Vec::new();
    if let Some(spec) = merge_ranges {
        let list = spec.cast::<PyList>().map_err(|_| {
            value_err("merge_ranges must be a list of (first_row, first_col, last_row, last_col, value[, format]) tuples".into())
        })?;
        for item in list.iter() {
            let parts: Vec<Bound<'_, PyAny>> = item.extract().map_err(|_| {
                value_err("each merge range must be a tuple/list of 5 or 6 items".into())
            })?;
            if parts.len() < 5 || parts.len() > 6 {
                return Err(value_err(format!(
                    "each merge range needs 5 or 6 items (first_row, first_col, last_row, last_col, value[, format]), got {}",
                    parts.len()
                )));
            }
            let r1: u32 = parts[0].extract()?;
            let c1: u16 = parts[1].extract()?;
            let r2: u32 = parts[2].extract()?;
            let c2: u16 = parts[3].extract()?;
            if r2 < r1 || c2 < c1 {
                return Err(value_err(format!(
                    "merge range ({r1}, {c1}, {r2}, {c2}) is inverted: last_row/last_col must not precede first_row/first_col"
                )));
            }
            // A merge can only be written before the rows it covers are
            // flushed, and headers/data start at `header_row`.
            if r2 >= header_row {
                return Err(value_err(format!(
                    "merge range ({r1}, {c1}, {r2}, {c2}) reaches row {r2}, but the header row is {header_row} and data follows it. \
Merged ranges must sit strictly above the header row — raise header_row to at least {} to make room.",
                    r2 + 1
                )));
            }
            let value = parts[4].str()?.to_string();
            let fmt = match parts.get(5) {
                Some(f) if !f.is_none() => Some(
                    f.extract::<crate::format::Format>()
                        .map_err(|_| {
                            value_err("merge range format must be a Format object".into())
                        })?
                        .inner,
                ),
                _ => None,
            };
            merges.push((r1, c1, r2, c2, value, fmt));
        }
    }

    let heights = row_keyed(row_heights, "row_heights", |v| {
        let h: f64 = v
            .extract()
            .map_err(|_| value_err("row_heights values must be numbers".into()))?;
        if h < 0.0 {
            return Err(value_err("row_heights values must not be negative".into()));
        }
        Ok(h)
    })?;

    let formats = row_keyed(row_formats, "row_formats", |v| {
        Ok(v.extract::<crate::format::Format>()
            .map_err(|_| value_err("row_formats values must be Format objects".into()))?
            .inner)
    })?;

    Ok(SheetLayout {
        header_row,
        merges,
        row_heights: heights,
        row_formats: formats,
        band_color: banded_rows,
        autofilter,
        totals,
        totals_label,
        totals_format,
        na_text,
        inf_text,
        page: crate::page_setup::PageSetup::from_py(page_setup)?,
        view: crate::sheet_view::SheetView::from_py(sheet_view)?,
        ignore: crate::ignore_errors::IgnoreErrors::from_py(ignore_errors)?,
        validations: crate::data_validation::DataValidations::from_py(data_validations)?,
        outline: crate::outline::Outline::from_py(outline)?,
        notes: crate::notes::Notes::from_py(notes)?,
        images: crate::images::Images::from_py(images)?,
        sparklines: crate::sparklines::Sparklines::from_py(sparklines)?,
        charts: crate::charts::Charts::from_py(charts)?,
        conditional: crate::conditional_format::ConditionalFormats::from_py(conditional_formats)?,
    })
}
