use pyo3::prelude::*;
use pyo3::types::{PyDateAccess, PyDateTime, PyTimeAccess};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};
use std::collections::HashMap;

use crate::data_types::WorksheetData;
use crate::utils::validate_sheet_name;

fn write_worksheet_content(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    records: &WorksheetData,
    password: Option<&String>,
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

    worksheet.autofit();
    if let Some(password) = password {
        worksheet.protect_with_password(password);
    }

    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None))]
pub fn write_worksheets(
    records_with_sheet_name: Vec<HashMap<String, WorksheetData>>,
    file_name: String,
    password: Option<String>,
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
            let _ = worksheet.set_name(sheet_name);
            write_worksheet_content(&mut worksheet, &records, password.as_ref())?;
        }
    }

    workbook.save(&file_name).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
    })?;
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None))]
pub fn write_worksheet(
    records: WorksheetData,
    file_name: String,
    sheet_name: Option<String>,
    password: Option<String>,
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

    write_worksheet_content(&mut worksheet, &records, password.as_ref())?;

    workbook.save(&file_name).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
    })?;
    Ok(())
}