//! Fast CSV writer — writes Records, Pandas, or Polars data to CSV.

use std::io::Write;

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyTimeAccess};
use pyo3::Py;

use crate::helpers::write_bytes_to_target;

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
    let delim_bytes = delim.as_bytes();
    if delim_bytes.len() != 1 {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "CSV delimiter must be a single ASCII byte",
        ));
    }
    let delim_byte = delim_bytes[0];

    let bound = records.bind(py);
    // Heuristic: 16 bytes per cell is a decent starting point.
    let mut output: Vec<u8> = Vec::with_capacity(4096);

    // Fast path: Arrow zero-copy if the object exposes `__arrow_c_stream__`
    // (Pandas ≥2.0, Polars). Falls back to the per-object paths below on
    // failure (e.g. empty Null-typed columns).
    if bound.hasattr("__arrow_c_stream__")? {
        if write_csv_via_arrow(&records, py, &mut output, delim_byte).is_ok() {
            return write_bytes_to_target(py, &output, file_name);
        }
        output.clear();
    }

    if bound.hasattr("columns")? {
        let columns: Vec<String> = bound.getattr("columns")?.extract()?;
        write_csv_row_strings(&mut output, &columns, delim_byte);

        if bound.hasattr("get_column")? {
            // Polars
            let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(columns.len());
            for header in &columns {
                let col_series = records.call_method1(py, "get_column", (header.as_str(),))?;
                col_lists.push(col_series.call_method0(py, "to_list")?);
            }
            let nrows: usize = records.call_method0(py, "__len__")?.extract(py)?;
            let bound_lists: Vec<Bound<pyo3::types::PyList>> = col_lists
                .iter()
                .map(|c| c.bind(py).cast::<pyo3::types::PyList>().cloned())
                .collect::<Result<_, _>>()?;

            for row in 0..nrows {
                for (i, col_list) in bound_lists.iter().enumerate() {
                    if i > 0 {
                        output.push(delim_byte);
                    }
                    let item = col_list.get_item(row)?;
                    write_csv_value(&mut output, &item)?;
                }
                output.push(b'\n');
            }
        } else {
            // Pandas — iterate rows via `.values`
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
                    for key in row_dict.keys().iter() {
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

    write_bytes_to_target(py, &output, file_name)
}

fn write_csv_via_arrow(
    records: &Py<PyAny>,
    py: Python,
    output: &mut Vec<u8>,
    delim: u8,
) -> PyResult<()> {
    let reader = crate::arrow_ffi::stream_to_reader(records, py)?;
    let schema = reader.schema();
    let headers: Vec<String> = schema
        .fields()
        .iter()
        .map(|f| f.name().clone())
        .collect();
    write_csv_row_strings(output, &headers, delim);

    for batch_result in reader {
        let batch = batch_result.map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to read Arrow batch: {}",
                e
            ))
        })?;
        crate::arrow_writer::write_arrow_batch_csv(output, &batch, delim)?;
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
        return Ok(());
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        write_csv_escaped(output, s.to_str()?);
        return Ok(());
    }

    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        output.extend_from_slice(if b.is_true() { b"true" } else { b"false" });
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
        let mut buf = itoa::Buffer::new();
        output.extend_from_slice(buf.format(val).as_bytes());
        return Ok(());
    }

    if let Ok(dt) = value.cast::<PyDateTime>() {
        write!(
            output,
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            dt.get_year(),
            dt.get_month(),
            dt.get_day(),
            dt.get_hour(),
            dt.get_minute(),
            dt.get_second()
        )
        .expect("write to Vec<u8> is infallible");
        return Ok(());
    }

    if let Ok(d) = value.cast::<PyDate>() {
        write!(
            output,
            "{:04}-{:02}-{:02}",
            d.get_year(),
            d.get_month(),
            d.get_day()
        )
        .expect("write to Vec<u8> is infallible");
        return Ok(());
    }

    // numpy scalar fallback: bool before f64
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

    let s = value.to_string();
    write_csv_escaped(output, &s);
    Ok(())
}

/// Escape per RFC 4180: quote if value contains delimiter-relevant chars.
fn write_csv_escaped(output: &mut Vec<u8>, val: &str) {
    if val.contains(',') || val.contains('\n') || val.contains('\r') || val.contains('"') {
        output.push(b'"');
        for b in val.bytes() {
            if b == b'"' {
                output.push(b'"');
            }
            output.push(b);
        }
        output.push(b'"');
    } else {
        output.extend_from_slice(val.as_bytes());
    }
}
