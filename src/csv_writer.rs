//! Fast CSV writer — writes Records, Pandas, or Polars data to CSV.

use std::io::Write;

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyTimeAccess};
use pyo3::Py;

use crate::helpers::{write_bytes_to_target, write_csv_escaped, ColType};

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
        // Records path (list of dicts / generator). First-row type cache
        // mirrors the Excel Records path — skips the full type cascade
        // after the first row when the column's Python type is stable.
        let mut headers: Vec<String> = Vec::new();
        let mut headers_written = false;
        let mut col_types: Vec<ColType> = Vec::new();

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
                    col_types.resize(headers.len(), ColType::Unknown);
                    headers_written = true;
                }

                for (col, value) in row_dict.values().iter().enumerate() {
                    if col > 0 {
                        output.push(delim_byte);
                    }
                    let cached = col_types.get(col).copied().unwrap_or(ColType::Unknown);
                    let hit = try_cached_csv_value(&mut output, &value, cached)?;
                    if !hit {
                        let detected = write_csv_value(&mut output, &value)?;
                        if col < col_types.len() && col_types[col] == ColType::Unknown {
                            col_types[col] = detected;
                        }
                    }
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

fn emit_string(output: &mut Vec<u8>, s: &str) {
    write_csv_escaped(output, s);
}

fn emit_bool(output: &mut Vec<u8>, b: bool) {
    output.extend_from_slice(if b { b"true" } else { b"false" });
}

fn emit_float(output: &mut Vec<u8>, val: f64) {
    if !val.is_nan() && !val.is_infinite() {
        let mut buf = ryu::Buffer::new();
        output.extend_from_slice(buf.format(val).as_bytes());
    }
}

fn emit_int(output: &mut Vec<u8>, val: i64) {
    let mut buf = itoa::Buffer::new();
    output.extend_from_slice(buf.format(val).as_bytes());
}

fn emit_datetime(output: &mut Vec<u8>, dt: &Bound<PyDateTime>) {
    let _ = write!(
        output,
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
        dt.get_year(),
        dt.get_month(),
        dt.get_day(),
        dt.get_hour(),
        dt.get_minute(),
        dt.get_second()
    );
}

fn emit_date(output: &mut Vec<u8>, d: &Bound<PyDate>) {
    let _ = write!(
        output,
        "{:04}-{:02}-{:02}",
        d.get_year(),
        d.get_month(),
        d.get_day()
    );
}

/// Try the cached column type. Returns `true` if the value matched and was
/// written; `false` if the cache was empty or the value's type didn't match
/// (caller then falls back to the full cascade).
fn try_cached_csv_value(
    output: &mut Vec<u8>,
    value: &Bound<PyAny>,
    cached: ColType,
) -> PyResult<bool> {
    if value.is_none() {
        return Ok(true);
    }
    match cached {
        ColType::String => {
            if let Ok(s) = value.cast::<pyo3::types::PyString>() {
                emit_string(output, s.to_str()?);
                return Ok(true);
            }
        }
        ColType::Bool => {
            if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
                emit_bool(output, b.is_true());
                return Ok(true);
            }
        }
        ColType::Float => {
            if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
                emit_float(output, f.value());
                return Ok(true);
            }
        }
        ColType::Int => {
            // Python bool is a subclass of int and casts to PyInt, so a bool
            // landing in an Int-cached column must miss here and fall back to
            // the cascade (which checks Bool first) — else `True` → `1`.
            if value.cast::<pyo3::types::PyBool>().is_err() {
                if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
                    emit_int(output, i.extract()?);
                    return Ok(true);
                }
            }
        }
        ColType::DateTime => {
            if let Ok(dt) = value.cast::<PyDateTime>() {
                emit_datetime(output, &dt);
                return Ok(true);
            }
        }
        ColType::Date => {
            if let Ok(d) = value.cast::<PyDate>() {
                emit_date(output, &d);
                return Ok(true);
            }
        }
        ColType::Unknown => {}
    }
    Ok(false)
}

fn write_csv_value(output: &mut Vec<u8>, value: &Bound<PyAny>) -> PyResult<ColType> {
    if value.is_none() {
        return Ok(ColType::Unknown);
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        emit_string(output, s.to_str()?);
        return Ok(ColType::String);
    }

    // Bool BEFORE Int (Python bool is subclass of int)
    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        emit_bool(output, b.is_true());
        return Ok(ColType::Bool);
    }

    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        emit_float(output, f.value());
        return Ok(ColType::Float);
    }

    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        emit_int(output, i.extract()?);
        return Ok(ColType::Int);
    }

    if let Ok(dt) = value.cast::<PyDateTime>() {
        emit_datetime(output, &dt);
        return Ok(ColType::DateTime);
    }

    // Date AFTER DateTime (datetime is subclass of date)
    if let Ok(d) = value.cast::<PyDate>() {
        emit_date(output, &d);
        return Ok(ColType::Date);
    }

    // numpy scalar fallback: bool before f64
    if let Ok(val) = value.extract::<bool>() {
        emit_bool(output, val);
        return Ok(ColType::Bool);
    }

    if let Ok(val) = value.extract::<f64>() {
        emit_float(output, val);
        return Ok(ColType::Float);
    }

    emit_string(output, &value.to_string());
    Ok(ColType::String)
}

