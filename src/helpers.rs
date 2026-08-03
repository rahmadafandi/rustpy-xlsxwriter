//! Shared helpers for Excel/CSV writing paths.

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyDict, PyList, PyTimeAccess};
use pyo3::Py;
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, Worksheet};

use crate::worksheet::xlsx_err;

/// Column type used for first-row caching in Records path and
/// as the return tag from [`write_py_any`].
#[repr(u8)]
#[derive(Copy, Clone, PartialEq, Eq, Default, Debug)]
pub enum ColType {
    #[default]
    Unknown = 0,
    String = 1,
    Float = 2,
    Bool = 3,
    Int = 4,
    DateTime = 5,
    Date = 6,
}

/// Convert a Python `datetime` to `ExcelDateTime`.
pub fn py_datetime_to_excel(dt: &Bound<PyDateTime>) -> PyResult<ExcelDateTime> {
    ExcelDateTime::from_ymd(
        dt.get_year() as u16,
        dt.get_month() as u8,
        dt.get_day() as u8,
    )
    .map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Failed to create datetime: {}",
            e
        ))
    })?
    .and_hms(
        dt.get_hour() as u16,
        dt.get_minute() as u8,
        dt.get_second() as u8,
    )
    .map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Failed to create timestamp: {}",
            e
        ))
    })
}

/// Convert a Python `date` to `ExcelDateTime`.
pub fn py_date_to_excel(d: &Bound<PyDate>) -> PyResult<ExcelDateTime> {
    ExcelDateTime::from_ymd(
        d.get_year() as u16,
        d.get_month() as u8,
        d.get_day() as u8,
    )
    .map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Failed to create date: {}",
            e
        ))
    })
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
}

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

/// Read a `{row_index: value}` mapping into a sorted `Vec<(u32, T)>`.
fn row_keyed<T>(
    spec: Option<&Bound<'_, PyAny>>,
    what: &str,
    mut convert: impl FnMut(&Bound<'_, PyAny>) -> PyResult<T>,
) -> PyResult<Vec<(u32, T)>> {
    let Some(spec) = spec else { return Ok(Vec::new()) };
    let dict = spec.cast::<PyDict>().map_err(|_| {
        value_err(format!("{what} must be a dict keyed by row index"))
    })?;
    let mut out = Vec::with_capacity(dict.len());
    for (key, val) in dict.iter() {
        let row: u32 = key.extract().map_err(|_| {
            value_err(format!("{what}: row index must be a non-negative int"))
        })?;
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
) -> PyResult<SheetLayout> {
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
        let h: f64 = v.extract().map_err(|_| {
            value_err("row_heights values must be numbers".into())
        })?;
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
    })
}

/// Write a header cell on `row`, optionally bold, and mark the column
/// as an index (bold) column if listed in `index_columns`.
/// When `header_fmt` is `Some`, it wins over `bold_headers` for the cell itself.
#[allow(clippy::too_many_arguments)]
pub fn write_header(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    header: &str,
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
    header_fmt: Option<&Format>,
) -> PyResult<()> {
    if let Some(fmt) = header_fmt {
        worksheet
            .write_string_with_format(row, col, header, fmt)
            .map_err(xlsx_err)?;
        return Ok(());
    }
    if bold_headers {
        worksheet
            .write_string_with_format(row, col, header, bold_fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet
            .write_string(row, col, header)
            .map_err(xlsx_err)?;
    }
    if let Some(cols) = index_columns {
        if cols.iter().any(|c| c == header) {
            worksheet.set_column_format(col, bold_fmt).map_err(xlsx_err)?;
        }
    }
    Ok(())
}

/// Write every header cell for a sheet on `row` via [`write_header`].
#[allow(clippy::too_many_arguments)]
pub fn write_all_headers(
    worksheet: &mut Worksheet,
    row: u32,
    headers: &[String],
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
    header_fmt: Option<&Format>,
) -> PyResult<()> {
    for (col, header) in headers.iter().enumerate() {
        write_header(
            worksheet,
            row,
            col as u16,
            header,
            bold_headers,
            bold_fmt,
            index_columns,
            header_fmt,
        )?;
    }
    Ok(())
}

/// Write a numeric cell with optional float format. NaN/Inf → empty string.
pub fn write_num(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: f64,
    float_fmt: Option<&Format>,
) -> PyResult<()> {
    if val.is_nan() || val.is_infinite() {
        // Keep the format on the blank so a banded row has no unshaded hole.
        write_string_opt(worksheet, row, col, "", float_fmt)?;
    } else if let Some(fmt) = float_fmt {
        worksheet
            .write_number_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
    }
    Ok(())
}

/// Write a numeric cell, with an optional explicit format. Unlike [`write_num`]
/// this does NOT guard NaN/Inf — callers use it for integer values (and Arrow
/// integral/float16 columns) that can never be NaN/Inf.
pub fn write_number_opt(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: f64,
    fmt: Option<&Format>,
) -> PyResult<()> {
    if let Some(fmt) = fmt {
        worksheet
            .write_number_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
    }
    Ok(())
}

/// Write a string cell, with an optional explicit format.
pub fn write_string_opt(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: &str,
    fmt: Option<&Format>,
) -> PyResult<()> {
    if let Some(fmt) = fmt {
        worksheet
            .write_string_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_string(row, col, val).map_err(xlsx_err)?;
    }
    Ok(())
}

/// Write a string as a clickable link, falling back to plain text when it is
/// not one.
///
/// `write_url` rejects anything it cannot classify — ordinary text, an empty
/// string, or a URL past Excel's 2083-character limit — so calling it blindly
/// over a column would abort the whole export on the first blank cell. A value
/// that is not a link is therefore written as its literal text: visible and
/// correct, just not clickable.
pub fn write_url_or_text(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: &str,
    fmt: Option<&Format>,
) -> PyResult<()> {
    let wrote = match fmt {
        Some(f) => worksheet.write_url_with_format(row, col, val, f).is_ok(),
        None => worksheet.write_url(row, col, val).is_ok(),
    };
    if wrote {
        return Ok(());
    }
    write_string_opt(worksheet, row, col, val, fmt)
}

/// Resolve `url_columns` (column names) to their positions in `headers`.
/// Unknown names warn and are skipped, matching `column_formats`.
pub fn resolve_url_columns(
    url_columns: Option<&Vec<String>>,
    headers: &[String],
    py: Python,
) -> PyResult<Vec<bool>> {
    let mut flags = vec![false; headers.len()];
    let Some(names) = url_columns else {
        return Ok(flags);
    };
    let warnings = py.import("warnings")?;
    for name in names {
        match headers.iter().position(|h| h == name) {
            Some(idx) => flags[idx] = true,
            None => {
                warnings.call_method1(
                    "warn",
                    (format!("url_columns: unknown column '{name}', skipped"),),
                )?;
            }
        }
    }
    Ok(flags)
}

/// Write a boolean cell, with an optional explicit format.
pub fn write_bool_opt(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: bool,
    fmt: Option<&Format>,
) -> PyResult<()> {
    if let Some(fmt) = fmt {
        worksheet
            .write_boolean_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_boolean(row, col, val).map_err(xlsx_err)?;
    }
    Ok(())
}

/// Write a datetime cell, with an optional explicit format. `None` relies on
/// the column format having been set already.
pub fn write_datetime_opt(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: &ExcelDateTime,
    fmt: Option<&Format>,
) -> PyResult<()> {
    if let Some(fmt) = fmt {
        worksheet
            .write_datetime_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_datetime(row, col, val).map_err(xlsx_err)?;
    }
    Ok(())
}

/// `true` if `val` begins with a character a spreadsheet may interpret as a
/// formula (`=`, `+`, `-`, `@`) — the classic CSV-injection vector.
fn needs_formula_guard(val: &str) -> bool {
    matches!(val.as_bytes().first(), Some(b'=' | b'+' | b'-' | b'@'))
}

/// RFC-4180 escape with optional formula-injection guard. When `guard` is set
/// and `val` starts with `= + - @`, a leading `'` is emitted so spreadsheet
/// apps treat the cell as text instead of a formula. The guard byte is placed
/// inside the quotes when the field is quoted.
pub fn write_csv_escaped_guarded(output: &mut Vec<u8>, val: &str, guard: bool) {
    let prefix = guard && needs_formula_guard(val);
    if val.contains(',') || val.contains('\n') || val.contains('\r') || val.contains('"') {
        output.push(b'"');
        if prefix {
            output.push(b'\'');
        }
        for b in val.bytes() {
            if b == b'"' {
                output.push(b'"');
            }
            output.push(b);
        }
        output.push(b'"');
    } else {
        if prefix {
            output.push(b'\'');
        }
        output.extend_from_slice(val.as_bytes());
    }
}

/// Save a workbook to a file path or writable buffer.
pub fn save_workbook(
    py: Python,
    workbook: &mut Workbook,
    file_or_buffer: Py<PyAny>,
) -> PyResult<()> {
    if let Ok(file_name) = file_or_buffer.extract::<String>(py) {
        workbook.save(&file_name).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!(
                "Failed to save workbook: {}",
                e
            ))
        })?;
        return Ok(());
    }

    let buffer = workbook.save_to_buffer().map_err(|e| {
        // Match the file-save path (PyIOError) so a save failure surfaces as
        // OSError regardless of whether the target is a path or a buffer.
        PyErr::new::<pyo3::exceptions::PyIOError, _>(format!(
            "Failed to save workbook to buffer: {}",
            e
        ))
    })?;
    write_bytes_to_target(py, &buffer, file_or_buffer)
}

/// Write raw bytes to a file path or writable buffer.
pub fn write_bytes_to_target(
    py: Python,
    bytes: &[u8],
    file_or_buffer: Py<PyAny>,
) -> PyResult<()> {
    if let Ok(file_name) = file_or_buffer.extract::<String>(py) {
        std::fs::write(&file_name, bytes).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to write file: {}", e))
        })?;
        return Ok(());
    }

    if let Ok(write_method) = file_or_buffer.getattr(py, "write") {
        let py_bytes = pyo3::types::PyBytes::new(py, bytes);
        write_method.call1(py, (py_bytes,))?;
        return Ok(());
    }

    Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
        "Argument must be a string path or a file-like object with a 'write' method",
    ))
}

/// `true` if `w` is a usable Excel column width (finite, non-negative).
fn is_valid_width(w: f64) -> bool {
    w.is_finite() && w >= 0.0
}

/// Emit a Python `UserWarning` from Rust.
fn warn_py(py: Python, msg: &str) -> PyResult<()> {
    py.import("warnings")?.call_method1("warn", (msg,))?;
    Ok(())
}

/// Apply explicit column widths AFTER `autofit()` so they override it.
///
/// `uniform` (from `column_width`) sets every column as a base layer;
/// `spec` (from `column_widths`) then overrides individual columns —
/// a dict keyed by header name, or a positional list. Unknown names,
/// out-of-range list indices, and invalid widths emit a `UserWarning`
/// and are skipped. An unsupported `spec` type raises `ValueError`.
pub fn apply_column_widths(
    worksheet: &mut Worksheet,
    headers: &[String],
    uniform: Option<f64>,
    spec: Option<&Bound<'_, PyAny>>,
    py: Python,
) -> PyResult<()> {
    let ncols = headers.len() as u16;

    if let Some(w) = uniform {
        if is_valid_width(w) {
            if ncols > 0 {
                worksheet
                    .set_column_range_width(0, ncols - 1, w)
                    .map_err(xlsx_err)?;
            }
        } else {
            warn_py(py, &format!("column_width: invalid width {w}, skipped"))?;
        }
    }

    let Some(spec) = spec else {
        return Ok(());
    };

    if let Ok(dict) = spec.cast::<PyDict>() {
        for (key, val) in dict.iter() {
            let name: String = key.extract()?;
            let width: f64 = val.extract()?;
            match headers.iter().position(|h| h == &name) {
                Some(idx) if is_valid_width(width) => {
                    worksheet
                        .set_column_width(idx as u16, width)
                        .map_err(xlsx_err)?;
                }
                Some(_) => warn_py(
                    py,
                    &format!("column_widths: invalid width {width} for '{name}', skipped"),
                )?,
                None => warn_py(
                    py,
                    &format!("column_widths: unknown column '{name}', skipped"),
                )?,
            }
        }
    } else if let Ok(list) = spec.cast::<PyList>() {
        for (idx, item) in list.iter().enumerate() {
            let width: f64 = item.extract()?;
            if idx as u16 >= ncols {
                warn_py(
                    py,
                    &format!(
                        "column_widths: index {idx} out of range ({ncols} columns), skipped"
                    ),
                )?;
                continue;
            }
            if is_valid_width(width) {
                worksheet
                    .set_column_width(idx as u16, width)
                    .map_err(xlsx_err)?;
            } else {
                warn_py(
                    py,
                    &format!("column_widths: invalid width {width} at index {idx}, skipped"),
                )?;
            }
        }
    } else {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "column_widths must be a dict (by column name) or a list (positional)",
        ));
    }

    Ok(())
}
