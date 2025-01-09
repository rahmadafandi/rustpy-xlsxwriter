use pyo3::prelude::*;
use rust_xlsxwriter::Workbook;
use std::collections::HashMap;
// TODO: Add this back in when we have a better solution
// use std::sync::mpsc;
// use std::thread;
// use num_cpus;

/// Returns the version of the library
#[pyfunction]
fn get_version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

/// Returns the name of the library
#[pyfunction]
fn get_name() -> String {
    env!("CARGO_PKG_NAME").to_string()
}

/// Returns the authors of the library
#[pyfunction]
fn get_authors() -> String {
    env!("CARGO_PKG_AUTHORS").to_string()
}

/// Returns the description of the library
#[pyfunction]
fn get_description() -> String {
    env!("CARGO_PKG_DESCRIPTION").to_string()
}

/// Returns the repository URL of the library
#[pyfunction]
fn get_repository() -> String {
    env!("CARGO_PKG_REPOSITORY").to_string()
}

/// Returns the homepage URL of the library
#[pyfunction]
fn get_homepage() -> String {
    env!("CARGO_PKG_HOMEPAGE").to_string()
}

/// Returns the license of the library
#[pyfunction]
fn get_license() -> String {
    env!("CARGO_PKG_LICENSE").to_string()
}

/// Validates that the sheet name meets Excel's requirements:
/// - Must be <= 31 characters
/// - Cannot contain [ ] : * ? / \
/// Returns true if valid, false if invalid
#[pyfunction]
fn validate_sheet_name(name: &str) -> bool {
    if name.len() > 31 {
        return false;
    }
    !name.contains(&['[', ']', ':', '*', '?', '/', '\\'][..])
}

#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None))]
fn save_records_multiple_sheets(
    records_with_sheet_name: Vec<HashMap<String, Vec<HashMap<String, Option<PyObject>>>>>,
    file_name: String,
    password: Option<String>
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    for record_map in records_with_sheet_name {
        for (sheet_name, records) in record_map {
            // Validate sheet name
            if !validate_sheet_name(&sheet_name) {
                return Err(
                    PyErr::new::<pyo3::exceptions::PyValueError, _>(
                        format!("Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\", sheet_name)
                    )
                );
            }

            let worksheet = workbook.add_worksheet();
            let _ = worksheet.set_name(sheet_name);
            if let Some(first_record) = records.first() {
                let headers: Vec<String> = first_record.keys().cloned().collect();
                println!("headers: {:?}", headers);
                for (col, header) in headers.iter().enumerate() {
                    let _ = worksheet
                        .write_string(0, col as u16, header.to_string())
                        .map_err(|e| {
                            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(
                                format!("Failed to write header: {}", e)
                            )
                        });
                }

                for (row, record) in records.iter().enumerate() {
                    for (col, header) in headers.iter().enumerate() {
                        match record.get(header) {
                            Some(Some(value)) => {
                                Python::with_gil(|py| {
                                    if value.is_none(py) {
                                        let _ = worksheet
                                            .write_string((row + 1) as u32, col as u16, "")
                                            .map_err(|e| format!("Failed to write data: {}", e));
                                    } else if let Ok(str_val) = value.extract::<String>(py) {
                                        let _ = worksheet
                                            .write_string((row + 1) as u32, col as u16, str_val)
                                            .map_err(|e| format!("Failed to write data: {}", e));
                                    } else if let Ok(int_val) = value.extract::<i64>(py) {
                                        let _ = worksheet
                                            .write_number(
                                                (row + 1) as u32,
                                                col as u16,
                                                int_val as f64
                                            )
                                            .map_err(|e| format!("Failed to write data: {}", e));
                                    } else if let Ok(float_val) = value.extract::<f64>(py) {
                                        let _ = worksheet
                                            .write_number((row + 1) as u32, col as u16, float_val)
                                            .map_err(|e| format!("Failed to write data: {}", e));
                                    } else if let Ok(bool_val) = value.extract::<bool>(py) {
                                        let _ = worksheet
                                            .write_boolean((row + 1) as u32, col as u16, bool_val)
                                            .map_err(|e| format!("Failed to write data: {}", e));
                                    } else {
                                        // For any other type, convert to string
                                        if let Ok(str_val) = value.extract::<String>(py) {
                                            let _ = worksheet
                                                .write_string((row + 1) as u32, col as u16, str_val)
                                                .map_err(|e|
                                                    format!("Failed to write data: {}", e)
                                                );
                                        }
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
            if let Some(password) = &password {
                worksheet.protect_with_password(password);
            }
        }
    }
    workbook
        .save(&file_name)
        .map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
        })?;
    Ok(())
}

/// Save records to an Excel file and return a result with error handling.
#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None))]
fn save_records(
    records: Vec<HashMap<String, Option<PyObject>>>,
    file_name: String,
    sheet_name: Option<String>,
    password: Option<String>
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    if let Some(sheet_name) = sheet_name {
        // Validate sheet name if provided
        if !validate_sheet_name(&sheet_name) {
            return Err(
                PyErr::new::<pyo3::exceptions::PyValueError, _>(
                    format!("Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\", sheet_name)
                )
            );
        }
        let _ = worksheet.set_name(sheet_name);
    }

    // Write headers
    if let Some(first_record) = records.first() {
        let headers: Vec<String> = first_record.keys().cloned().collect();
        for (col, header) in headers.iter().enumerate() {
            let _ = worksheet
                .write_string(0, col as u16, header.to_string())
                .map_err(|e| {
                    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(
                        format!("Failed to write header: {}", e)
                    )
                });
        }

        for (row, record) in records.iter().enumerate() {
            for (col, header) in headers.iter().enumerate() {
                match record.get(header) {
                    Some(Some(value)) => {
                        Python::with_gil(|py| {
                            if value.is_none(py) {
                                let _ = worksheet
                                    .write_string((row + 1) as u32, col as u16, "")
                                    .map_err(|e| format!("Failed to write data: {}", e));
                            } else if let Ok(str_val) = value.extract::<String>(py) {
                                let _ = worksheet
                                    .write_string((row + 1) as u32, col as u16, str_val)
                                    .map_err(|e| format!("Failed to write data: {}", e));
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
                            } else {
                                // For any other type, convert to string
                                if let Ok(str_val) = value.extract::<String>(py) {
                                    let _ = worksheet
                                        .write_string((row + 1) as u32, col as u16, str_val)
                                        .map_err(|e| format!("Failed to write data: {}", e));
                                }
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
        worksheet.protect_with_password(&password);
    }

    // Save the workbook
    let _ = workbook
        .save(&file_name)
        .map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
        });

    Ok(())
}

// TODO: Add this back in when we have a better solution
// #[pyfunction]
// #[pyo3(signature = (records, file_name, sheet_name = None, password = None))]
// fn save_records_multithread(
//     records: Vec<HashMap<String, String>>,
//     file_name: String,
//     sheet_name: Option<String>,
//     password: Option<String>
// ) -> PyResult<()> {
//     let mut workbook = Workbook::new();
//     let worksheet = workbook.add_worksheet();
//     if let Some(sheet_name) = sheet_name {
//         let _ = worksheet.set_name(sheet_name);
//     }

//     // Write headers
//     let headers = match records.first() {
//         Some(first_record) => {
//             let headers: Vec<String> = first_record.keys().cloned().collect();
//             // Write headers to worksheet
//             for (col, header) in headers.iter().enumerate() {
//                 worksheet
//                     .write_string(0, col as u16, header.to_string())
//                     .map_err(|e| {
//                         PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(
//                             format!("Failed to write header: {}", e)
//                         )
//                     })?;
//             }
//             headers
//         }
//         None => {
//             return Ok(());
//         } // Return early if no records
//     };

//     // Calculate optimal chunk size based on CPU cores and record count
//     let num_cpus = num_cpus::get();
//     let chunk_size = (records.len() / num_cpus).max(1);
//     let chunks: Vec<_> = records.chunks(chunk_size).collect();

//     let (tx, rx) = mpsc::channel();
//     let mut handles = vec![];

//     // Process chunks in parallel
//     for (chunk_idx, chunk) in chunks.iter().enumerate() {
//         let tx = tx.clone();
//         let chunk = chunk.to_vec();
//         let headers = headers.clone();

//         let handle = thread::spawn(move || {
//             let mut rows = Vec::with_capacity(chunk.len() * headers.len());
//             for (row_idx, record) in chunk.iter().enumerate() {
//                 for (col, header) in headers.iter().enumerate() {
//                     if let Some(value) = record.get(header) {
//                         rows.push((chunk_idx * chunk_size + row_idx, col, value.to_string()));
//                     }
//                 }
//             }
//             if let Err(e) = tx.send(rows) {
//                 eprintln!("Failed to send rows: {}", e);
//             }
//         });
//         handles.push(handle);
//     }
//     drop(tx);

//     // Write data from all threads
//     let mut error_occurred = false;
//     for rows in rx {
//         for (row, col, value) in rows {
//             if let Err(e) = worksheet.write_string((row + 1) as u32, col as u16, value) {
//                 error_occurred = true;
//                 eprintln!("Failed to write data: {}", e);
//             }
//         }
//     }

//     // Wait for all threads to complete
//     for handle in handles {
//         if let Err(e) = handle.join() {
//             error_occurred = true;
//             eprintln!("Thread panicked: {:?}", e);
//         }
//     }

//     if error_occurred {
//         return Err(
//             PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(
//                 "Failed to write some data to worksheet"
//             )
//         );
//     }

//     worksheet.autofit();
//     if let Some(password) = password {
//         worksheet.protect_with_password(&password);
//     }

//     workbook
//         .save(&file_name)
//         .map_err(|e| {
//             PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
//         })?;

//     Ok(())
// }

/// A Python module implemented in Rust.
#[pymodule]
fn rustpy_xlsxwriter(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(save_records, m)?)?;
    // TODO: Add this back in when we have a better solution
    // m.add_function(wrap_pyfunction!(save_records_multithread, m)?)?;
    m.add_function(wrap_pyfunction!(save_records_multiple_sheets, m)?)?;
    m.add_function(wrap_pyfunction!(get_version, m)?)?;
    m.add_function(wrap_pyfunction!(get_name, m)?)?;
    m.add_function(wrap_pyfunction!(get_authors, m)?)?;
    m.add_function(wrap_pyfunction!(get_description, m)?)?;
    m.add_function(wrap_pyfunction!(get_repository, m)?)?;
    m.add_function(wrap_pyfunction!(get_homepage, m)?)?;
    m.add_function(wrap_pyfunction!(get_license, m)?)?;
    m.add_function(wrap_pyfunction!(validate_sheet_name, m)?)?;
    Ok(())
}
