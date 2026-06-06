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

/// Write a header cell at row 0, optionally bold, and mark the column
/// as an index (bold) column if listed in `index_columns`.
/// When `header_fmt` is `Some`, it wins over `bold_headers` for the cell itself.
pub fn write_header(
    worksheet: &mut Worksheet,
    col: u16,
    header: &str,
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
    header_fmt: Option<&Format>,
) -> PyResult<()> {
    if let Some(fmt) = header_fmt {
        worksheet
            .write_string_with_format(0, col, header, fmt)
            .map_err(xlsx_err)?;
        return Ok(());
    }
    if bold_headers {
        worksheet
            .write_string_with_format(0, col, header, bold_fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet
            .write_string(0, col, header)
            .map_err(xlsx_err)?;
    }
    if let Some(cols) = index_columns {
        if cols.iter().any(|c| c == header) {
            worksheet.set_column_format(col, bold_fmt).map_err(xlsx_err)?;
        }
    }
    Ok(())
}

/// Write every header cell for a sheet (row 0) via [`write_header`].
pub fn write_all_headers(
    worksheet: &mut Worksheet,
    headers: &[String],
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
    header_fmt: Option<&Format>,
) -> PyResult<()> {
    for (col, header) in headers.iter().enumerate() {
        write_header(
            worksheet,
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
        worksheet.write_string(row, col, "").map_err(xlsx_err)?;
    } else if let Some(fmt) = float_fmt {
        worksheet
            .write_number_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
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
        PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
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
