//! Fast CSV writer — writes Records, Pandas, or Polars data to CSV.

use std::io::Write;

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyInt, PyTimeAccess};
use pyo3::Py;

use crate::cell::{classify_and_write, try_cached, CellWriter};
use crate::helpers::{write_bytes_to_target, write_csv_escaped_guarded, ColType};

/// Write data to CSV (file path or buffer).
///
/// When `sanitize_formulas` is `true`, string fields that begin with
/// `= + - @` are prefixed with a single quote so spreadsheet apps treat them
/// as text rather than executable formulas (CSV-injection mitigation). It is
/// off by default to keep output byte-identical for existing callers.
#[pyfunction]
#[pyo3(signature = (records, file_name, delimiter = None, sanitize_formulas = false))]
pub fn write_csv(
    py: Python,
    records: Py<PyAny>,
    file_name: Py<PyAny>,
    delimiter: Option<String>,
    sanitize_formulas: bool,
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
        if write_csv_via_arrow(&records, py, &mut output, delim_byte, sanitize_formulas).is_ok() {
            return write_bytes_to_target(py, &output, file_name);
        }
        output.clear();
    }

    if bound.hasattr("columns")? {
        let columns: Vec<String> = bound.getattr("columns")?.extract()?;
        write_csv_row_strings(&mut output, &columns, delim_byte, sanitize_formulas);

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
                    let mut sink = CsvCell::new(&mut output, sanitize_formulas);
                    classify_and_write(&item, &mut sink)?;
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
                            let mut sink = CsvCell::new(&mut output, sanitize_formulas);
                            classify_and_write(&item, &mut sink)?;
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
                    write_csv_row_strings(&mut output, &headers, delim_byte, sanitize_formulas);
                    col_types.resize(headers.len(), ColType::Unknown);
                    headers_written = true;
                }

                // Iterate the dict directly (insertion order == header order)
                // to avoid allocating a fresh `values()` list per row.
                for (col, (_key, value)) in row_dict.iter().enumerate() {
                    if col > 0 {
                        output.push(delim_byte);
                    }
                    let cached = col_types.get(col).copied().unwrap_or(ColType::Unknown);
                    let mut sink = CsvCell::new(&mut output, sanitize_formulas);
                    if !try_cached(&value, cached, &mut sink)? {
                        let detected = classify_and_write(&value, &mut sink)?;
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
    sanitize: bool,
) -> PyResult<()> {
    let reader = crate::arrow_ffi::stream_to_reader(records, py)?;
    let schema = reader.schema();
    let headers: Vec<String> = schema
        .fields()
        .iter()
        .map(|f| f.name().clone())
        .collect();
    write_csv_row_strings(output, &headers, delim, sanitize);

    for batch_result in reader {
        let batch = batch_result.map_err(crate::arrow_ffi::batch_read_err)?;
        crate::arrow_writer::write_arrow_batch_csv(output, &batch, delim, sanitize)?;
    }
    Ok(())
}

fn write_csv_row_strings(output: &mut Vec<u8>, values: &[String], delim: u8, sanitize: bool) {
    for (i, val) in values.iter().enumerate() {
        if i > 0 {
            output.push(delim);
        }
        write_csv_escaped_guarded(output, val, sanitize);
    }
    output.push(b'\n');
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

/// [`CellWriter`] sink that appends one Python scalar to a CSV byte buffer.
/// The type-detection order lives in [`crate::cell`]; this only encodes the
/// per-type CSV serialization (and the optional formula-injection guard for
/// strings).
struct CsvCell<'a> {
    output: &'a mut Vec<u8>,
    sanitize: bool,
}

impl<'a> CsvCell<'a> {
    fn new(output: &'a mut Vec<u8>, sanitize: bool) -> Self {
        CsvCell { output, sanitize }
    }
}

impl CellWriter for CsvCell<'_> {
    fn write_none(&mut self) -> PyResult<()> {
        // CSV: a null is an empty field — emit nothing.
        Ok(())
    }

    fn write_str(&mut self, s: &str) -> PyResult<()> {
        write_csv_escaped_guarded(self.output, s, self.sanitize);
        Ok(())
    }

    fn write_bool(&mut self, b: bool) -> PyResult<()> {
        self.output
            .extend_from_slice(if b { b"true" } else { b"false" });
        Ok(())
    }

    fn write_float(&mut self, f: f64) -> PyResult<()> {
        if !f.is_nan() && !f.is_infinite() {
            let mut buf = ryu::Buffer::new();
            self.output.extend_from_slice(buf.format(f).as_bytes());
        }
        Ok(())
    }

    fn write_int(&mut self, i: &Bound<'_, PyInt>) -> PyResult<()> {
        let val: i64 = i.extract()?;
        let mut buf = itoa::Buffer::new();
        self.output.extend_from_slice(buf.format(val).as_bytes());
        Ok(())
    }

    fn write_datetime(&mut self, dt: &Bound<'_, PyDateTime>) -> PyResult<()> {
        emit_datetime(self.output, dt);
        Ok(())
    }

    fn write_date(&mut self, d: &Bound<'_, PyDate>) -> PyResult<()> {
        emit_date(self.output, d);
        Ok(())
    }
}
