//! Fast CSV writer — writes Records, Pandas, or Polars data to CSV format.

use std::io::Write;

use pyo3::prelude::*;
use pyo3::types::{PyDateAccess, PyDateTime, PyTimeAccess, PyDate};
use pyo3::Py;

/// Write data to CSV (file path or buffer).
#[pyfunction]
#[pyo3(signature = (records, file_name, delimiter = None))]
pub fn write_csv(
    py: Python,
    records: Py<PyAny>,
    file_name: Py<PyAny>,
    delimiter: Option<String>,
) -> PyResult<()> {
    let delim = delimiter.unwrap_or_else(|| ",".to_string());
    let delim_byte = delim.as_bytes()[0];

    let mut output: Vec<u8> = Vec::new();

    // Check if it's a DataFrame (Pandas/Polars) with __arrow_c_stream__
    // or has "columns" attribute
    let bound = records.bind(py);

    if bound.getattr("columns").is_ok() {
        // DataFrame path — get columns and iterate
        let columns: Vec<String> = bound.getattr("columns")?.extract()?;

        // Write header
        write_csv_row_strings(&mut output, &columns, delim_byte);

        // Get data via tolist per column or iterrows
        if bound.getattr("get_column").is_ok() {
            // Polars
            let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(columns.len());
            for header in &columns {
                let col_series = records.call_method1(py, "get_column", (header.as_str(),))?;
                let col_list = col_series.call_method0(py, "to_list")?;
                col_lists.push(col_list);
            }
            let nrows: usize = records.call_method0(py, "__len__")?.extract(py)?;
            for row in 0..nrows {
                let mut first = true;
                for col_list in &col_lists {
                    if !first {
                        output.push(delim_byte);
                    }
                    first = false;
                    let item = col_list.bind(py).get_item(row)?;
                    write_csv_value(&mut output, &item)?;
                }
                output.push(b'\n');
            }
        } else {
            // Pandas — use itertuples
            let values = records.getattr(py, "values")?;
            if let Ok(rows) = values.bind(py).try_iter() {
                for row_res in rows {
                    let row = row_res?;
                    if let Ok(items) = row.try_iter() {
                        let mut first = true;
                        for item_res in items {
                            let item = item_res?;
                            if !first {
                                output.push(delim_byte);
                            }
                            first = false;
                            write_csv_value(&mut output, &item)?;
                        }
                        output.push(b'\n');
                    }
                }
            }
        }
    } else {
        // Records path (list of dicts / generator)
        let mut headers: Vec<String> = Vec::new();
        let mut headers_written = false;

        if let Ok(rows) = bound.try_iter() {
            for row_res in rows {
                let row_obj = row_res?;
                let row_dict = row_obj.cast::<pyo3::types::PyDict>().map_err(|_| {
                    PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                        "Items in records must be dictionaries",
                    )
                })?;

                if !headers_written {
                    let keys = row_dict.keys();
                    for key in keys.iter() {
                        headers.push(key.extract::<String>()?);
                    }
                    write_csv_row_strings(&mut output, &headers, delim_byte);
                    headers_written = true;
                }

                let mut first = true;
                for value in row_dict.values().iter() {
                    if !first {
                        output.push(delim_byte);
                    }
                    first = false;
                    write_csv_value(&mut output, &value)?;
                }
                output.push(b'\n');
            }
        }
    }

    // Write to file or buffer
    if let Ok(file_name_str) = file_name.extract::<String>(py) {
        std::fs::write(&file_name_str, &output).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to write CSV: {}", e))
        })?;
    } else if let Ok(write_method) = file_name.getattr(py, "write") {
        let py_bytes = pyo3::types::PyBytes::new(py, &output);
        write_method.call1(py, (py_bytes,))?;
    } else {
        return Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
            "Argument must be a string path or a file-like object with a 'write' method",
        ));
    }

    Ok(())
}

fn write_csv_row_strings(output: &mut Vec<u8>, values: &[String], delim: u8) {
    for (i, val) in values.iter().enumerate() {
        if i > 0 {
            output.push(delim);
        }
        write_csv_escaped(output, val);
    }
    output.push(b'\n');
}

fn write_csv_value(output: &mut Vec<u8>, value: &Bound<PyAny>) -> PyResult<()> {
    if value.is_none() {
        // Write nothing for None
        return Ok(());
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        let val = s.to_str()?;
        write_csv_escaped(output, val);
        return Ok(());
    }

    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        if b.is_true() {
            output.extend_from_slice(b"true");
        } else {
            output.extend_from_slice(b"false");
        }
        return Ok(());
    }

    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        let val = f.value();
        if val.is_nan() || val.is_infinite() {
            return Ok(());
        }
        let mut buf = ryu::Buffer::new();
        output.extend_from_slice(buf.format(val).as_bytes());
        return Ok(());
    }

    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        let val: i64 = i.extract()?;
        output.extend_from_slice(val.to_string().as_bytes());
        return Ok(());
    }

    if let Ok(dt) = value.cast::<PyDateTime>() {
        write!(
            output,
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            dt.get_year(), dt.get_month(), dt.get_day(),
            dt.get_hour(), dt.get_minute(), dt.get_second()
        ).unwrap();
        return Ok(());
    }

    if let Ok(d) = value.cast::<PyDate>() {
        write!(
            output,
            "{:04}-{:02}-{:02}",
            d.get_year(), d.get_month(), d.get_day()
        ).unwrap();
        return Ok(());
    }

    // Fallback: try numeric extraction (numpy scalars)
    if let Ok(val) = value.extract::<bool>() {
        output.extend_from_slice(if val { b"true" } else { b"false" });
        return Ok(());
    }

    if let Ok(val) = value.extract::<f64>() {
        if !val.is_nan() && !val.is_infinite() {
            let mut buf = ryu::Buffer::new();
            output.extend_from_slice(buf.format(val).as_bytes());
        }
        return Ok(());
    }

    // Final fallback: string representation
    let s = value.to_string();
    write_csv_escaped(output, &s);
    Ok(())
}

/// Write a string value with CSV escaping (quote if contains delimiter, newline, or quote)
fn write_csv_escaped(output: &mut Vec<u8>, val: &str) {
    if val.contains(',') || val.contains('\n') || val.contains('\r') || val.contains('"') {
        output.push(b'"');
        for b in val.bytes() {
            if b == b'"' {
                output.push(b'"'); // escape quote with double quote
            }
            output.push(b);
        }
        output.push(b'"');
    } else {
        output.extend_from_slice(val.as_bytes());
    }
}
