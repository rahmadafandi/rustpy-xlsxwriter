use pyo3::{Py, PyAny, Python};
use pyo3::prelude::*;
use pyo3::types::{PyDateAccess, PyDateTime, PyTimeAccess};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};
use std::collections::HashSet;

use crate::data_types::{FreezePaneConfig, WorksheetData};
use crate::utils::validate_sheet_name;

/// Helper to convert rust_xlsxwriter errors to PyErr
fn xlsx_err(e: impl std::fmt::Display) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("Excel write error: {}", e))
}

fn save_workbook(py: Python, workbook: &mut Workbook, file_or_buffer: Py<PyAny>) -> PyResult<()> {
    if let Ok(file_name) = file_or_buffer.extract::<String>(py) {
        workbook.save(&file_name).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
        })?;
    } else {
        let buffer = workbook.save_to_buffer().map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to save workbook to buffer: {}",
                e
            ))
        })?;

        if let Ok(write_method) = file_or_buffer.getattr(py, "write") {
            let py_bytes = pyo3::types::PyBytes::new(py, &buffer);
            write_method.call1(py, (py_bytes,))?;
        } else {
            return Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                "Argument 'file_name' must be a string path or a file-like object with a 'write' method",
            ));
        }
    }
    Ok(())
}


fn write_worksheet_content(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    records: &WorksheetData,
    password: Option<&String>,
    freeze_row: Option<u32>,
    freeze_col: Option<u16>,
    float_format: Option<&String>,
    index_columns: Option<&Vec<String>>,
    autofit: bool,
    py: Python,
) -> PyResult<()> {
    // Pre-create Format objects once instead of per-cell
    let float_fmt = float_format.map(|s| Format::new().set_num_format(s));
    let datetime_fmt = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
    let bold_fmt = Format::new().set_bold();
    let mut datetime_cols_set: HashSet<u16> = HashSet::new();

    match records {
        WorksheetData::Records(records_list) => {
            if let Ok(rows) = records_list.bind(py).try_iter() {
                let mut headers: Vec<String> = Vec::new();
                let mut headers_written = false;

                for (row_idx, row_res) in rows.enumerate() {
                    let row_obj = row_res?;
                    let row_dict = row_obj.cast::<pyo3::types::PyDict>()
                        .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>("Items in records must be dictionaries"))?;

                    if !headers_written {
                        let keys = row_dict.keys();
                        for key in keys.iter() {
                            let key_str = key.extract::<String>()?;
                            headers.push(key_str);
                        }

                        for (col, header) in headers.iter().enumerate() {
                            worksheet
                                .write_string(0, col as u16, header)
                                .map_err(xlsx_err)?;

                            if let Some(index_cols) = index_columns {
                                if index_cols.contains(header) {
                                    worksheet
                                        .set_column_format(col as u16, &bold_fmt)
                                        .map_err(xlsx_err)?;
                                }
                            }
                        }
                        headers_written = true;
                    }

                    for (col, header) in headers.iter().enumerate() {
                        match row_dict.get_item(header)? {
                            Some(value) => {
                                write_py_any_bound(
                                    worksheet,
                                    (row_idx + 1) as u32,
                                    col as u16,
                                    &value,
                                    float_fmt.as_ref(),
                                    Some(&datetime_fmt),
                                    &mut datetime_cols_set,
                                )?;
                            }
                            None => {
                                worksheet
                                    .write_string((row_idx + 1) as u32, col as u16, "")
                                    .map_err(xlsx_err)?;
                            }
                        }
                    }
                }
            }
        }
        WorksheetData::DataFrame(df) => {
            let columns = df.getattr(py, "columns")?;
            let headers: Vec<String> = columns.extract(py)?;

            for (col, header) in headers.iter().enumerate() {
                worksheet
                    .write_string(0, col as u16, header)
                    .map_err(xlsx_err)?;

                if let Some(index_cols) = index_columns {
                    if index_cols.contains(header) {
                        worksheet
                            .set_column_format(col as u16, &bold_fmt)
                            .map_err(xlsx_err)?;
                    }
                }
            }

            let values = df.getattr(py, "values")?;
            if let Ok(rows) = values.bind(py).try_iter() {
                for (row_idx, row_res) in rows.enumerate() {
                    let row = row_res?;
                    if let Ok(items) = row.try_iter() {
                        for (col_idx, item_res) in items.enumerate() {
                            let item = item_res?;
                            write_py_any_bound(
                                worksheet,
                                (row_idx + 1) as u32,
                                col_idx as u16,
                                &item,
                                float_fmt.as_ref(),
                                Some(&datetime_fmt),
                                &mut datetime_cols_set,
                            )?;
                        }
                    }
                }
            }
        }
    }

    // Set freeze panes if specified
    if freeze_row.is_some() || freeze_col.is_some() {
        worksheet
            .set_freeze_panes(freeze_row.unwrap_or(0), freeze_col.unwrap_or(0))
            .map_err(xlsx_err)?;
    }

    if autofit {
        worksheet.autofit();
    }
    if let Some(password) = password {
        worksheet.protect_with_password(password);
    }

    Ok(())
}

fn write_py_any_bound(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &Bound<PyAny>,
    float_fmt: Option<&Format>,
    datetime_fmt: Option<&Format>,
    datetime_cols_set: &mut HashSet<u16>,
) -> PyResult<()> {
    if value.is_none() {
        worksheet.write_string(row, col, "").map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        worksheet.write_string(row, col, s.to_str()?).map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        let val = f.value();
        if let Some(fmt) = float_fmt {
            worksheet.write_number_with_format(row, col, val, fmt).map_err(xlsx_err)?;
        } else {
            worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        }
        return Ok(());
    }

    // Check Bool BEFORE Int (Python bool is subclass of int)
    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        worksheet.write_boolean(row, col, b.is_true()).map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        let val: f64 = i.extract()?;
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(datetime) = value.cast::<PyDateTime>() {
        let year = datetime.get_year() as u16;
        let month = datetime.get_month() as u8;
        let day = datetime.get_day() as u8;
        let hour = datetime.get_hour() as u16;
        let minute = datetime.get_minute() as u8;
        let second = datetime.get_second() as u8;

        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }

        let excel_datetime = ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Failed to create datetime: {}",
                    e
                ))
            })?
            .and_hms(hour, minute, second)
            .map_err(|e| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Failed to create timestamp: {}",
                    e
                ))
            })?;

        worksheet.write_datetime(row, col, &excel_datetime).map_err(xlsx_err)?;
        return Ok(());
    }

    // Fallback: try bool extraction for numpy.bool_ (must be before f64, since bool extracts as f64 too)
    if let Ok(val) = value.extract::<bool>() {
        worksheet.write_boolean(row, col, val).map_err(xlsx_err)?;
        return Ok(());
    }

    // Fallback: try numeric extraction for numpy scalar types (numpy.int64, numpy.float64, etc.)
    if let Ok(val) = value.extract::<f64>() {
        if let Some(fmt) = float_fmt {
            worksheet.write_number_with_format(row, col, val, fmt).map_err(xlsx_err)?;
        } else {
            worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        }
        return Ok(());
    }

    // Final fallback to string representation
    worksheet.write_string(row, col, value.to_string()).map_err(xlsx_err)?;
    Ok(())
}


#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None, float_format = None, index_columns = None, autofit = true))]
pub fn write_worksheets(
    py: Python,
    records_with_sheet_name: Vec<indexmap::IndexMap<String, WorksheetData>>,
    file_name: Py<PyAny>,
    password: Option<String>,
    freeze_panes: Option<FreezePaneConfig>,
    float_format: Option<String>,
    index_columns: Option<Vec<String>>,
    autofit: bool,
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    for record_map in records_with_sheet_name {
        for (sheet_name, records) in record_map {
            if !validate_sheet_name(&sheet_name) {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\",
                    sheet_name
                )));
            }

            let mut worksheet = workbook.add_worksheet_with_constant_memory();
            worksheet.set_name(&sheet_name).map_err(xlsx_err)?;

            let (mut freeze_row, mut freeze_col) = (None, None);
            if let Some(ref freeze_panes_config) = freeze_panes {
                if let Some(general_config) = freeze_panes_config.config.get("general") {
                    if let Some(&row) = general_config.get("row") {
                        freeze_row = Some(row as u32);
                    }
                    if let Some(&col) = general_config.get("col") {
                        freeze_col = Some(col as u16);
                    }
                }

                if let Some(sheet_config) = freeze_panes_config.config.get(&sheet_name) {
                    if let Some(&row) = sheet_config.get("row") {
                        freeze_row = Some(row as u32);
                    }
                    if let Some(&col) = sheet_config.get("col") {
                        freeze_col = Some(col as u16);
                    }
                }
            }

            write_worksheet_content(
                &mut worksheet,
                &records,
                password.as_ref(),
                freeze_row,
                freeze_col,
                float_format.as_ref(),
                index_columns.as_ref(),
                autofit,
                py,
            )?;
        }
    }

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None, freeze_row = None, freeze_col = None, float_format = None, index_columns = None, autofit = true))]
pub fn write_worksheet(
    py: Python,
    records: WorksheetData,
    file_name: Py<PyAny>,
    sheet_name: Option<String>,
    password: Option<String>,
    freeze_row: Option<u32>,
    freeze_col: Option<u16>,
    float_format: Option<String>,
    index_columns: Option<Vec<String>>,
    autofit: bool,
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    let mut worksheet = workbook.add_worksheet_with_constant_memory();

    if let Some(sheet_name) = sheet_name {
        if !validate_sheet_name(&sheet_name) {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\",
                sheet_name
            )));
        }
        worksheet.set_name(sheet_name).map_err(xlsx_err)?;
    }

    write_worksheet_content(
        &mut worksheet,
        &records,
        password.as_ref(),
        freeze_row,
        freeze_col,
        float_format.as_ref(),
        index_columns.as_ref(),
        autofit,
        py,
    )?;

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}
