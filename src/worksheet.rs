use pyo3::{Py, PyAny, Python};
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
        WorksheetData::Records(records_list) => {
            // Get iterator over the list
            if let Ok(rows) = records_list.bind(py).try_iter() {
                let mut headers: Vec<String> = Vec::new();
                let mut headers_written = false;

                for (row_idx, row_res) in rows.enumerate() {
                    let row_obj = row_res?;
                    let row_dict = row_obj.cast::<pyo3::types::PyDict>()
                        .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>("Items in records must be dictionaries"))?;

                    if !headers_written {
                        // Extract headers from the first row keys
                        // We use keys() to get a list of keys
                        let keys = row_dict.keys();
                         for key in keys.iter() {
                            let key_str = key.extract::<String>()?;
                            headers.push(key_str);
                        }
                        
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
                        headers_written = true;
                    }

                     // Write data
                    for (col, header) in headers.iter().enumerate() {
                        // Use get_item which doesn't clone if we just check for None
                        match row_dict.get_item(header)? {
                            Some(value) => {
                                write_py_any_bound(worksheet, (row_idx + 1) as u32, col as u16, &value, float_format)?;
                            }
                            None => {
                                // Key missing, write empty string or nothing?
                                // Original behavior was effectively empty string for None
                                let _ = worksheet.write_string((row_idx + 1) as u32, col as u16, "");
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
                                write_py_any_bound(worksheet, (row_idx + 1) as u32, col_idx as u16, &item, float_format)?;
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

fn write_py_any_bound(worksheet: &mut rust_xlsxwriter::Worksheet, row: u32, col: u16, value: &Bound<PyAny>, float_format: Option<&String>) -> PyResult<()> {
    // Check None first
    if value.is_none() {
        let _ = worksheet.write_string(row, col, "");
        return Ok(());
    }

    // Check String next (most common)
    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        let _ = worksheet.write_string(row, col, s.to_str()?);
        return Ok(());
    }

    // Check Float
    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        let val = f.value();
         if let Some(fmt_str) = float_format {
            let format = Format::new().set_num_format(fmt_str);
            let _ = worksheet.write_number_with_format(row, col, val, &format);
        } else {
            let _ = worksheet.write_number(row, col, val);
        }
        return Ok(());
    }

    // Check Int
    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        // rust_xlsxwriter write_number takes f64, so safe to cast unless huge int
        let val: f64 = i.extract()?;
        let _ = worksheet.write_number(row, col, val);
        return Ok(());
    }

    // Check Bool
    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        let _ = worksheet.write_boolean(row, col, b.is_true());
        return Ok(());
    }

    // Check DateTime (less common, usually)
    if let Ok(datetime) = value.cast::<PyDateTime>() {
         let year = datetime.get_year() as u16;
        let month = datetime.get_month() as u8;
        let day = datetime.get_day() as u8;
        let hour = datetime.get_hour() as u16;
        let minute = datetime.get_minute() as u8;
        let second = datetime.get_second() as u8;

        // Note: Creating Format every time might be slow? But rust_xlsxwriter handles it efficiently usually.
        // Optimization: Could cache format if possible, but difficult here without context.
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

         let _ = worksheet.write_datetime(row, col, &excel_datetime);
         return Ok(());
    }

    // Fallback to string representation
    let _ = worksheet.write_string(row, col, value.to_string());
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
