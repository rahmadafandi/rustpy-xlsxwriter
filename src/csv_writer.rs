//! Fast CSV writer — writes Records, Pandas, or Polars data to CSV.

use std::io::Write;

use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyInt, PyTimeAccess};
use pyo3::Py;

use crate::cell::{classify_and_write, try_cached, CellWriter};
use crate::helpers::{write_bytes_to_target, write_csv_escaped_guarded, ColType, NumReps};

/// Turn a failed `try_iter` into a clear message. Skipping the loop instead
/// would write an empty file and report success.
fn not_iterable(_: PyErr) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyTypeError, _>(
        "records must be an iterable of dicts, a DataFrame, or an Arrow stream",
    )
}

// Not a `///` doc comment: pyo3 turns those into `__doc__`, which then
// shadows the richer stub entry in `rustpy_xlsxwriter.pyi` — the one
// mypy, IDEs and the docs site all read. Keep the prose in one place.
// Write data to CSV (file path or buffer).
//
// When `sanitize_formulas` is `true`, string fields that begin with
// `= + - @` are prefixed with a single quote so spreadsheet apps treat them
// as text rather than executable formulas (CSV-injection mitigation). It is
// off by default to keep output byte-identical for existing callers.
//
// `bom` prefixes the UTF-8 byte order mark, which is what makes Excel on
// Windows read the file as UTF-8 instead of the system code page.
//
// `columns` selects and orders the output columns; `header` writes the header
// row. Both apply to every input path.
#[pyfunction]
#[pyo3(signature = (
    records,
    file_name,
    delimiter = None,
    sanitize_formulas = false,
    bom = false,
    columns = None,
    header = true,
    na_rep = None,
    inf_value = None,
))]
#[allow(clippy::too_many_arguments)]
pub fn write_csv(
    py: Python,
    records: Py<PyAny>,
    file_name: Py<PyAny>,
    delimiter: Option<String>,
    sanitize_formulas: bool,
    bom: bool,
    columns: Option<Vec<String>>,
    header: bool,
    na_rep: Option<String>,
    inf_value: Option<String>,
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
    if bom {
        output.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
    }
    let selection = columns.as_ref();
    let reps = NumReps {
        na: na_rep.as_deref(),
        inf: inf_value.as_deref(),
    };

    // Fast path: Arrow zero-copy if the object exposes `__arrow_c_stream__`
    // (Pandas ≥2.0, Polars). Falls back to the per-object paths below on
    // failure (e.g. empty Null-typed columns).
    if bound.hasattr("__arrow_c_stream__")? {
        match write_csv_via_arrow(
            &records,
            py,
            &mut output,
            delim_byte,
            sanitize_formulas,
            selection,
            header,
            reps,
        ) {
            Ok(()) => return write_bytes_to_target(py, &output, file_name),
            // A bad `columns` is the caller's mistake, not a quirk of this
            // path, so it must not fall through to a slower one that would
            // raise the same error later.
            Err(e) if e.is_instance_of::<pyo3::exceptions::PyValueError>(py) => return Err(e),
            Err(_) => {}
        }
        output.truncate(if bom { 3 } else { 0 });
    }

    if bound.hasattr("columns")? {
        let all_columns: Vec<String> = bound.getattr("columns")?.extract()?;
        let indices = select_indices(&all_columns, selection)?;
        let columns = apply_selection(all_columns, indices.as_ref());
        if header {
            write_csv_row_strings(&mut output, &columns, delim_byte, sanitize_formulas);
        }

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
                    let mut sink = CsvCell::new(&mut output, sanitize_formulas, reps);
                    classify_and_write(&item, &mut sink)?;
                }
                output.push(b'\n');
            }
        } else {
            // Pandas — iterate rows via `.values`
            let values = records.getattr(py, "values")?;
            // Propagate rather than skip: silently emitting an empty file is
            // worse than saying the input could not be iterated.
            for row_res in values.bind(py).try_iter()? {
                let row = row_res?;
                match &indices {
                    // `.values` rows are positional, so a selection indexes
                    // into them rather than filtering names.
                    Some(ix) => {
                        for (n, &i) in ix.iter().enumerate() {
                            if n > 0 {
                                output.push(delim_byte);
                            }
                            let item = row.get_item(i)?;
                            let mut sink = CsvCell::new(&mut output, sanitize_formulas, reps);
                            classify_and_write(&item, &mut sink)?;
                        }
                    }
                    None => {
                        let mut first = true;
                        for item_res in row.try_iter()? {
                            let item = item_res?;
                            if !first {
                                output.push(delim_byte);
                            }
                            first = false;
                            let mut sink = CsvCell::new(&mut output, sanitize_formulas, reps);
                            classify_and_write(&item, &mut sink)?;
                        }
                    }
                }
                output.push(b'\n');
            }
        }
    } else {
        // Records path (list of dicts / generator). First-row type cache
        // mirrors the Excel Records path — skips the full type cascade
        // after the first row when the column's Python type is stable.
        let mut headers: Vec<String> = Vec::new();
        let mut headers_written = false;
        let mut col_types: Vec<ColType> = Vec::new();

        let rows: pyo3::Bound<'_, pyo3::types::PyIterator> =
            bound.try_iter().map_err(not_iterable)?;
        for row_res in rows {
            let row_obj = row_res?;
            let row_dict = row_obj.cast::<pyo3::types::PyDict>().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "Items in records must be dictionaries",
                )
            })?;

            if !headers_written {
                let mut keys: Vec<String> = Vec::new();
                for key in row_dict.keys().iter() {
                    keys.push(key.extract::<String>()?);
                }
                // Validated against the first row, which is the only row whose
                // keys are known before the stream is consumed.
                headers = apply_selection(keys.clone(), select_indices(&keys, selection)?.as_ref());
                if header {
                    write_csv_row_strings(&mut output, &headers, delim_byte, sanitize_formulas);
                }
                col_types.resize(headers.len(), ColType::Unknown);
                headers_written = true;
            }

            if selection.is_some() {
                // Selected: look each column up by name, since the dict's own
                // order is no longer the output order.
                for (col, name) in headers.iter().enumerate() {
                    if col > 0 {
                        output.push(delim_byte);
                    }
                    let mut sink = CsvCell::new(&mut output, sanitize_formulas, reps);
                    // A later row missing the key writes an empty field rather
                    // than shifting every column after it.
                    if let Some(value) = row_dict.get_item(name)? {
                        let cached = col_types.get(col).copied().unwrap_or(ColType::Unknown);
                        if !try_cached(&value, cached, &mut sink)? {
                            let detected = classify_and_write(&value, &mut sink)?;
                            if col_types[col] == ColType::Unknown {
                                col_types[col] = detected;
                            }
                        }
                    }
                }
            } else {
                // Iterate the dict directly (insertion order == header order)
                // to avoid allocating a fresh `values()` list per row.
                for (col, (_key, value)) in row_dict.iter().enumerate() {
                    if col > 0 {
                        output.push(delim_byte);
                    }
                    let cached = col_types.get(col).copied().unwrap_or(ColType::Unknown);
                    let mut sink = CsvCell::new(&mut output, sanitize_formulas, reps);
                    if !try_cached(&value, cached, &mut sink)? {
                        let detected = classify_and_write(&value, &mut sink)?;
                        if col < col_types.len() && col_types[col] == ColType::Unknown {
                            col_types[col] = detected;
                        }
                    }
                }
            }
            output.push(b'\n');
        }
    }

    write_bytes_to_target(py, &output, file_name)
}

#[allow(clippy::too_many_arguments)]
fn write_csv_via_arrow(
    records: &Py<PyAny>,
    py: Python,
    output: &mut Vec<u8>,
    delim: u8,
    sanitize: bool,
    columns: Option<&Vec<String>>,
    header: bool,
    reps: NumReps<'_>,
) -> PyResult<()> {
    let reader = crate::arrow_ffi::stream_to_reader(records, py)?;
    let schema = reader.schema();
    let headers: Vec<String> = schema.fields().iter().map(|f| f.name().clone()).collect();
    let indices = select_indices(&headers, columns)?;
    if header {
        let out = apply_selection(headers, indices.as_ref());
        write_csv_row_strings(output, &out, delim, sanitize);
    }

    for batch_result in reader {
        let batch = batch_result.map_err(crate::arrow_ffi::batch_read_err)?;
        match &indices {
            // Projection is an Arc clone per column, so the zero-copy path
            // stays zero-copy.
            Some(ix) => {
                let projected = batch
                    .project(ix)
                    .map_err(crate::arrow_ffi::batch_read_err)?;
                crate::arrow_writer::write_arrow_batch_csv(
                    output, &projected, delim, sanitize, reps,
                )?;
            }
            None => {
                crate::arrow_writer::write_arrow_batch_csv(output, &batch, delim, sanitize, reps)?
            }
        }
    }
    Ok(())
}

/// Positions of `columns` within `headers`, in the order asked for.
///
/// `None` means every column, as they come. An unknown name is an error rather
/// than a warning: `columns` decides the shape of the output, so dropping one
/// quietly would hand back a file that looks complete but is not.
fn select_indices(
    headers: &[String],
    columns: Option<&Vec<String>>,
) -> PyResult<Option<Vec<usize>>> {
    let Some(names) = columns else {
        return Ok(None);
    };
    let mut indices = Vec::with_capacity(names.len());
    for name in names {
        match headers.iter().position(|h| h == name) {
            Some(idx) => indices.push(idx),
            None => {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "columns: '{}' is not in the data (available: {})",
                    name,
                    headers.join(", ")
                )))
            }
        }
    }
    Ok(Some(indices))
}

/// Apply a selection to a header list, or hand it back untouched.
fn apply_selection(headers: Vec<String>, indices: Option<&Vec<usize>>) -> Vec<String> {
    match indices {
        Some(ix) => ix.iter().map(|&i| headers[i].clone()).collect(),
        None => headers,
    }
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
    reps: NumReps<'a>,
}

impl<'a> CsvCell<'a> {
    fn new(output: &'a mut Vec<u8>, sanitize: bool, reps: NumReps<'a>) -> Self {
        CsvCell {
            output,
            sanitize,
            reps,
        }
    }
}

impl CellWriter for CsvCell<'_> {
    fn write_none(&mut self) -> PyResult<()> {
        // A null is an empty field unless the caller named a representation.
        if let Some(text) = self.reps.na {
            write_csv_escaped_guarded(self.output, text, self.sanitize);
        }
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
        if f.is_nan() || f.is_infinite() {
            // Without a representation the field stays empty, as it always has.
            if let Some(text) = self.reps.text_for(f) {
                write_csv_escaped_guarded(self.output, &text, self.sanitize);
            }
            return Ok(());
        }
        let mut buf = ryu::Buffer::new();
        self.output.extend_from_slice(buf.format(f).as_bytes());
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
