use pyo3::{IntoPyObjectExt, Py, PyAny, Python};
use pyo3::prelude::*;
use pyo3::types::{PyDateAccess, PyDateTime, PyTimeAccess};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};
use std::collections::HashMap;

use crate::data_types::{FreezePaneConfig, WorksheetData};
use crate::utils::validate_sheet_name;

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
        
        // Check if object has a 'write' method
        if let Ok(write_method) = file_or_buffer.getattr(py, "write") {
             // Convert Vec<u8> to &[u8] for PyBytes
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
    py: Python,
) -> PyResult<()> {
    match records {
        WorksheetData::Records(records_vec) => {
             if let Some(first_record) = records_vec.first() {
                let headers: Vec<String> = first_record.hash.keys().cloned().collect();
                // Write headers
                for (col, header) in headers.iter().enumerate() {
                    worksheet
                        .write_string(0, col as u16, header.to_string())
                        .map_err(|e| {
                            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                                "Failed to write header: {}",
                                e
                            ))
                        })?;
                    
                     // Apply bold format if it's an index column
                    if let Some(index_cols) = index_columns {
                        if index_cols.contains(header) {
                            let bold = Format::new().set_bold();
                             let _ = worksheet.set_column_format(col as u16, &bold);
                        }
                    }
                }

                // Write data
                for (row, record) in records_vec.iter().enumerate() {
                    for (col, header) in headers.iter().enumerate() {
                        match record.hash.get(header) {
                            Some(Some(value)) => {
                                write_py_any(py, worksheet, (row + 1) as u32, col as u16, value, float_format)?;
                            }
                            Some(None) | None => {
                                let _ = worksheet
                                    .write_string((row + 1) as u32, col as u16, "")
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            }
                        }
                    }
                }
            }
        }
        WorksheetData::DataFrame(df) => {
            // Get columns
            let columns = df.getattr(py, "columns")?;
            let headers: Vec<String> = columns.extract(py)?;
            
            // Write headers
            for (col, header) in headers.iter().enumerate() {
                 worksheet
                    .write_string(0, col as u16, header.to_string())
                    .map_err(|e| {
                        PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                            "Failed to write header: {}",
                            e
                        ))
                    })?;

                    // Apply bold format if it's an index column
                    if let Some(index_cols) = index_columns {
                        if index_cols.contains(header) {
                            let bold = Format::new().set_bold();
                             let _ = worksheet.set_column_format(col as u16, &bold);
                        }
                    }
            }

            // Get values (numpy array or similar iterable)
            let values = df.getattr(py, "values")?;
            // Iterate over rows using bind(py) to get Bound<PyAny> which implements try_iter
            if let Ok(rows) = values.bind(py).try_iter() {
                 for (row_idx, row_res) in rows.enumerate() {
                      let row = row_res?;
                      // If row is array-like, iterate its items
                      if let Ok(items) = row.try_iter() {
                           for (col_idx, item_res) in items.enumerate() {
                                let item = item_res?;
                                // Unwrap item extraction
                                let py_item: Py<PyAny> = item.into_py_any(py)?;
                                write_py_any(py, worksheet, (row_idx + 1) as u32, col_idx as u16, &py_item, float_format)?;
                           }
                      }
                 }
            }
        }
    }

    // Set freeze panes if specified
    if let (Some(row), Some(col)) = (freeze_row, freeze_col) {
        worksheet.set_freeze_panes(row, col).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to set freeze panes: {}",
                e
            ))
        })?;
    } else if let Some(row) = freeze_row {
        worksheet.set_freeze_panes(row, 0).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to set freeze panes: {}",
                e
            ))
        })?;
    } else if let Some(col) = freeze_col {
        worksheet.set_freeze_panes(0, col).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to set freeze panes: {}",
                e
            ))
        })?;
    }

    worksheet.autofit();
    if let Some(password) = password {
        worksheet.protect_with_password(password);
    }

    Ok(())
}

fn write_py_any(py: Python, worksheet: &mut rust_xlsxwriter::Worksheet, row: u32, col: u16, value: &Py<PyAny>, float_format: Option<&String>) -> PyResult<()> {
    if value.is_none(py) {
        let _ = worksheet
            .write_string(row, col, "")
            .map_err(|e| format!("Failed to write data: {}", e));
    } else if let Ok(datetime) = value.bind(py).cast::<PyDateTime>() {
        let year = datetime.get_year() as u16;
        let month = datetime.get_month() as u8;
        let day = datetime.get_day() as u8;
        let hour = datetime.get_hour() as u16;
        let minute = datetime.get_minute() as u8;
        let second = datetime.get_second() as u8;
        let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
        let _ = worksheet.set_column_format(col, &format3);

        let excel_datetime = ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| format!("Failed to create datetime: {}", e))
            .and_then(|dt| {
                dt.and_hms(hour, minute, second).map_err(|e| {
                    format!("Failed to create timestamp: {}", e)
                })
            })
            .unwrap();

        let _ = worksheet
            .write_datetime(row, col, &excel_datetime)
            .map_err(|e| format!("Failed to write datetime: {}", e));
    } else if let Ok(int_val) = value.extract::<i64>(py) {
        let _ = worksheet
            .write_number(row, col, int_val as f64)
            .map_err(|e| format!("Failed to write data: {}", e));
    } else if let Ok(float_val) = value.extract::<f64>(py) {
        if let Some(fmt_str) = float_format {
            let format = Format::new().set_num_format(fmt_str);
            let _ = worksheet
                .write_number_with_format(row, col, float_val, &format)
                .map_err(|e| format!("Failed to write data: {}", e));
        } else {
            let _ = worksheet
                .write_number(row, col, float_val)
                .map_err(|e| format!("Failed to write data: {}", e));
        }
    } else if let Ok(bool_val) = value.extract::<bool>(py) {
        let _ = worksheet
            .write_boolean(row, col, bool_val)
            .map_err(|e| format!("Failed to write data: {}", e));
    } else if let Ok(str_val) = value.extract::<String>(py) {
        let _ = worksheet
            .write_string(row, col, str_val)
            .map_err(|e| format!("Failed to write data: {}", e));
    } else {
        let _ = worksheet
            .write_string(row, col, value.to_string())
            .map_err(|e| format!("Failed to write data: {}", e));
    }
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None, float_format = None, index_columns = None))]
pub fn write_worksheets(
    py: Python,
    records_with_sheet_name: Vec<HashMap<String, WorksheetData>>,
    file_name: Py<PyAny>,
    password: Option<String>,
    freeze_panes: Option<FreezePaneConfig>,
    float_format: Option<String>,
    index_columns: Option<Vec<String>>,
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
            let _ = worksheet.set_name(&sheet_name);

            let (mut freeze_row, mut freeze_col) = (None, None);
            if let Some(freeze_panes_config) = freeze_panes.clone() {
                // First check for general settings
                if let Some(general_config) = freeze_panes_config.clone().config.get("general") {
                    if let Some(row) = general_config.clone().get("row") {
                        freeze_row = Some(row.clone() as u32);
                    }
                    if let Some(col) = general_config.clone().get("col") {
                        freeze_col = Some(col.clone() as u16);
                    }
                }

                // Then check for sheet-specific settings which override general settings
                if let Some(sheet_config) = freeze_panes_config.clone().config.get(&sheet_name) {
                    if let Some(row) = sheet_config.clone().get("row") {
                        freeze_row = Some(row.clone() as u32);
                    }
                    if let Some(col) = sheet_config.clone().get("col") {
                        freeze_col = Some(col.clone() as u16);
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
                py,
            )?;
        }
    }

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None, freeze_row = None, freeze_col = None, float_format = None, index_columns = None))]
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
        let _ = worksheet.set_name(sheet_name);
    }

    write_worksheet_content(
        &mut worksheet,
        &records,
        password.as_ref(),
        freeze_row,
        freeze_col,
        float_format.as_ref(),
        index_columns.as_ref(),
        py,
    )?;

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}
