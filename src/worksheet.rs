use pyo3::prelude::*;
use pyo3::types::{PyDateAccess, PyDateTime, PyTimeAccess};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};
use std::collections::HashMap;

use crate::data_types::{FreezePaneConfig, WorksheetData};
use crate::utils::validate_sheet_name;

fn write_worksheet_content(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    records: &WorksheetData,
    password: Option<&String>,
    freeze_row: Option<u32>,
    freeze_col: Option<u16>,
) -> PyResult<()> {
    if let Some(first_record) = records.records.first() {
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
        }

        // Write data
        for (row, record) in records.records.iter().enumerate() {
            for (col, header) in headers.iter().enumerate() {
                match record.hash.get(header) {
                    Some(Some(value)) => {
                        Python::with_gil(|py| {
                            if value.is_none(py) {
                                let _ = worksheet
                                    .write_string((row + 1) as u32, col as u16, "")
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            } else if let Ok(datetime) = value.downcast_bound::<PyDateTime>(py) {
                                let year = datetime.get_year() as u16;
                                let month = datetime.get_month() as u8;
                                let day = datetime.get_day() as u8;
                                let hour = datetime.get_hour() as u16;
                                let minute = datetime.get_minute() as u8;
                                let second = datetime.get_second() as u8;
                                let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
                                let _ = worksheet.set_column_format(col as u16, &format3);

                                let excel_datetime = ExcelDateTime::from_ymd(year, month, day)
                                    .map_err(|e| format!("Failed to create datetime: {}", e))
                                    .and_then(|dt| {
                                        dt.and_hms(hour, minute, second).map_err(|e| {
                                            format!("Failed to create timestamp: {}", e)
                                        })
                                    })
                                    .unwrap();

                                let _ = worksheet
                                    .write_datetime((row + 1) as u32, col as u16, &excel_datetime)
                                    .map_err(|e| format!("Failed to write datetime: {}", e));
                            } else if let Ok(int_val) = value.extract::<i64>(py) {
                                let _ = worksheet
                                    .write_number((row + 1) as u32, col as u16, int_val as f64)
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            } else if let Ok(float_val) = value.extract::<f64>(py) {
                                let _ = worksheet
                                    .write_number((row + 1) as u32, col as u16, float_val)
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            } else if let Ok(bool_val) = value.extract::<bool>(py) {
                                let _ = worksheet
                                    .write_boolean((row + 1) as u32, col as u16, bool_val)
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            } else if let Ok(str_val) = value.extract::<String>(py) {
                                let _ = worksheet
                                    .write_string((row + 1) as u32, col as u16, str_val)
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            } else {
                                let _ = worksheet
                                    .write_string((row + 1) as u32, col as u16, value.to_string())
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            }
                        });
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

#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None))]
pub fn write_worksheets(
    records_with_sheet_name: Vec<HashMap<String, WorksheetData>>,
    file_name: String,
    password: Option<String>,
    freeze_panes: Option<FreezePaneConfig>,
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
            )?;
        }
    }

    workbook.save(&file_name).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
    })?;
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None, freeze_row = None, freeze_col = None))]
pub fn write_worksheet(
    records: WorksheetData,
    file_name: String,
    sheet_name: Option<String>,
    password: Option<String>,
    freeze_row: Option<u32>,
    freeze_col: Option<u16>,
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
    )?;

    workbook.save(&file_name).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
    })?;
    Ok(())
}
