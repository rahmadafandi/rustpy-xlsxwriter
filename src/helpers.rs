//! Shared helpers for Excel/CSV writing paths.

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyDict, PyList, PyTimeAccess};
use pyo3::Py;
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, Worksheet};
use std::borrow::Cow;
use std::path::PathBuf;

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
    ExcelDateTime::from_ymd(dt.get_year() as u16, dt.get_month(), dt.get_day())
        .map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "Failed to create datetime: {}",
                e
            ))
        })?
        .and_hms(dt.get_hour() as u16, dt.get_minute(), dt.get_second())
        .map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "Failed to create timestamp: {}",
                e
            ))
        })
}

/// Convert a Python `date` to `ExcelDateTime`.
pub fn py_date_to_excel(d: &Bound<PyDate>) -> PyResult<ExcelDateTime> {
    ExcelDateTime::from_ymd(d.get_year() as u16, d.get_month(), d.get_day()).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create date: {}", e))
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
        worksheet.write_string(row, col, header).map_err(xlsx_err)?;
    }
    if let Some(cols) = index_columns {
        if cols.iter().any(|c| c == header) {
            worksheet
                .set_column_format(col, bold_fmt)
                .map_err(xlsx_err)?;
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

/// How missing values and the infinities are rendered.
///
/// `None` writes an empty cell, which is what every version before this did,
/// so the default keeps existing files byte-identical. The reason to set one
/// is that a blank and a missing value are indistinguishable once written.
///
/// `na` deliberately covers `None`, an Arrow null *and* a float NaN together,
/// the way `pandas.to_csv(na_rep=...)` does. Keeping them apart would be a
/// trap: pandas turns NaN in a float column into an Arrow null, so a knob that
/// only caught true NaN would do nothing on the most common input of all.
#[derive(Clone, Copy, Default)]
pub struct NumReps<'a> {
    pub na: Option<&'a str>,
    pub inf: Option<&'a str>,
}

impl<'a> NumReps<'a> {
    /// Text for a missing value, or `""` when none was set.
    ///
    /// Takes `self` by value — the struct is `Copy`, and the borrow must be of
    /// the caller's strings rather than of `self`, or a sink holding a
    /// `NumReps` could not pass the text to its own `&mut self` method.
    pub fn na_text(self) -> &'a str {
        self.na.unwrap_or("")
    }

    /// Text for a non-finite `val`, or `None` to leave the cell empty.
    ///
    /// Negative infinity takes `inf` with a `-` in front, matching what
    /// `rust_xlsxwriter` and Excel use themselves ("INF" / "-INF").
    pub fn text_for(self, val: f64) -> Option<Cow<'a, str>> {
        if val.is_nan() {
            self.na.map(Cow::Borrowed)
        } else if val.is_infinite() {
            self.inf.map(|t| {
                if val.is_sign_negative() {
                    Cow::Owned(format!("-{t}"))
                } else {
                    Cow::Borrowed(t)
                }
            })
        } else {
            None
        }
    }
}

/// Write a numeric cell with optional float format. NaN/Inf follow `reps`.
pub fn write_num(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: f64,
    float_fmt: Option<&Format>,
    reps: NumReps<'_>,
) -> PyResult<()> {
    if val.is_nan() || val.is_infinite() {
        // Keep the format on the text so a banded row has no unshaded hole.
        let text = reps.text_for(val).unwrap_or(Cow::Borrowed(""));
        write_string_opt(worksheet, row, col, &text, float_fmt)?;
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
    text: Option<&str>,
) -> PyResult<()> {
    // Display text rides on the Url itself rather than a second write, which
    // would overwrite the format the first one applied.
    let link = match text {
        Some(t) => rust_xlsxwriter::Url::new(val).set_text(t),
        None => rust_xlsxwriter::Url::new(val),
    };
    let wrote = match fmt {
        Some(f) => worksheet.write_url_with_format(row, col, link, f).is_ok(),
        None => worksheet.write_url(row, col, link).is_ok(),
    };
    if wrote {
        return Ok(());
    }
    write_string_opt(worksheet, row, col, val, fmt)
}

/// One column's link settings.
#[derive(Clone, Copy, Default)]
pub struct UrlCol {
    /// The column was listed in `url_columns`, so its text becomes a link.
    pub link: bool,
    /// Index of the column supplying the display text, when the caller passed
    /// the mapping form. `None` shows the URL itself.
    pub text_col: Option<usize>,
}

/// Resolve `url_columns` against `headers`.
///
/// Two shapes, because showing the URL itself is rarely what a report wants:
/// a list names the link columns, a dict maps each link column to the column
/// holding its display text (`{"url": "product_name"}`).
///
/// Unknown names warn and are skipped, matching `column_formats` — a stray
/// name costs a link, not the export.
pub fn resolve_url_columns(
    url_columns: Option<&Bound<'_, PyAny>>,
    headers: &[String],
    py: Python,
) -> PyResult<Vec<UrlCol>> {
    let mut cols = vec![UrlCol::default(); headers.len()];
    let Some(spec) = url_columns else {
        return Ok(cols);
    };
    let warnings = py.import("warnings")?;
    let warn = |msg: String| -> PyResult<()> {
        warnings.call_method1("warn", (msg,))?;
        Ok(())
    };

    let pairs: Vec<(String, Option<String>)> = if let Ok(map) = spec.cast::<PyDict>() {
        map.iter()
            .map(|(k, v)| Ok((k.extract::<String>()?, Some(v.extract::<String>()?))))
            .collect::<PyResult<_>>()?
    } else {
        spec.extract::<Vec<String>>()?
            .into_iter()
            .map(|name| (name, None))
            .collect()
    };

    for (name, text_name) in pairs {
        let Some(idx) = headers.iter().position(|h| h == &name) else {
            warn(format!("url_columns: unknown column '{name}', skipped"))?;
            continue;
        };
        let text_col = match text_name {
            Some(t) => match headers.iter().position(|h| h == &t) {
                Some(ti) => Some(ti),
                None => {
                    warn(format!(
                        "url_columns: unknown display-text column '{t}' for '{name}', \
                         showing the URL instead"
                    ))?;
                    None
                }
            },
            None => None,
        };
        cols[idx] = UrlCol {
            link: true,
            text_col,
        };
    }
    Ok(cols)
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
    // The save is where the XML is assembled and the zip is deflated — about
    // two thirds of a write, and none of it touches a Python object. Holding
    // the GIL through it means concurrent writers cannot overlap at all,
    // which is why more threads never made a standard build any faster.
    // Resolving the target is the only part that needs the GIL, so it happens
    // first and the rest runs detached.
    if let Ok(path) = file_or_buffer.extract::<PathBuf>(py) {
        return py.detach(|| {
            workbook.save(&path).map_err(|e| {
                PyErr::new::<pyo3::exceptions::PyIOError, _>(format!(
                    "Failed to save workbook: {}",
                    e
                ))
            })?;
            Ok(())
        });
    }

    let buffer = py.detach(|| {
        workbook.save_to_buffer().map_err(|e| {
            // Match the file-save path (PyIOError) so a save failure surfaces
            // as OSError regardless of whether the target is a path or a
            // buffer.
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!(
                "Failed to save workbook to buffer: {}",
                e
            ))
        })
    })?;
    write_bytes_to_target(py, &buffer, file_or_buffer)
}

/// Write raw bytes to a file path or writable buffer.
pub fn write_bytes_to_target(py: Python, bytes: &[u8], file_or_buffer: Py<PyAny>) -> PyResult<()> {
    if let Ok(path) = file_or_buffer.extract::<PathBuf>(py) {
        std::fs::write(&path, bytes).map_err(|e| {
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
        "Argument must be a path (str or os.PathLike) or a file-like object with a 'write' method",
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
                    &format!("column_widths: index {idx} out of range ({ncols} columns), skipped"),
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
