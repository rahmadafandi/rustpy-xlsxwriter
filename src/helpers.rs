//! Shared helpers for Excel/CSV writing paths.

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyTimeAccess};
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
pub fn write_header(
    worksheet: &mut Worksheet,
    col: u16,
    header: &str,
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
) -> PyResult<()> {
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
