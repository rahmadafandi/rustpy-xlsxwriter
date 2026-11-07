//! Fast CSV writer — writes Records, Pandas, or Polars data to CSV.

use std::io::{Seek, SeekFrom, Write};

use pyo3::prelude::*;
use pyo3::types::{
    PyDate, PyDateAccess, PyDateTime, PyDelta, PyDeltaAccess, PyInt, PyTime, PyTimeAccess,
    PyTzInfoAccess,
};
use pyo3::Py;

use crate::cell::{classify_and_write, try_cached, CellWriter};
use crate::helpers::{
    apply_selection, located, select_indices, write_csv_escaped_guarded, ColType, Loc, NumReps,
};
use crate::options::value_err;

/// Output grows to about this much before a path target is written to.
const CHUNK: usize = 64 * 1024;

/// Where `write_csv` sends its bytes.
///
/// The rows are encoded into one `Vec<u8>`. For a path, that buffer is
/// emptied into the file whenever it passes [`CHUNK`], so memory stays at
/// about one chunk however large the file gets. The buffer already batches
/// the writes, so a `BufWriter` around the file would only copy them twice.
/// A file-like object still receives the whole file as one `bytes`, which
/// its `.write()` needs.
enum CsvSink {
    Path {
        path: std::path::PathBuf,
        /// Opened at the first chunk, not up front. A write that fails before
        /// then leaves a file already at the path untouched.
        open: Option<OpenFile>,
    },
    Writer(Py<PyAny>),
}

/// The file bytes are going into, and where it has to end up.
struct OpenFile {
    file: std::fs::File,
    /// A sibling temporary file to rename onto the target when the write
    /// finishes, or `None` when writing straight to the target.
    temp: Option<std::path::PathBuf>,
}

fn io_err(e: std::io::Error) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to write file: {e}"))
}

/// What is already at the target, in one `stat`.
enum Existing {
    /// Nothing there, or nothing readable: a temporary file has nothing to
    /// preserve and nothing to follow.
    Absent,
    /// A regular file, whose mode the replacement should inherit.
    RegularFile(std::fs::Permissions),
    /// A FIFO, a device such as `/dev/stdout`, a socket, a directory, or a
    /// symlink — which a rename would replace with a regular file instead of
    /// following. These are written through directly.
    Special,
}

fn look_at(path: &std::path::Path) -> Existing {
    match std::fs::symlink_metadata(path) {
        Ok(meta) if meta.file_type().is_file() => Existing::RegularFile(meta.permissions()),
        Ok(_) => Existing::Special,
        Err(_) => Existing::Absent,
    }
}

/// A sibling path for the temporary file. It must share the directory, since
/// `rename` is only atomic within one filesystem. The pid and a counter keep
/// two concurrent writes to one target apart.
fn temp_path_for(path: &std::path::Path) -> Option<std::path::PathBuf> {
    use std::sync::atomic::{AtomicU64, Ordering};
    static COUNTER: AtomicU64 = AtomicU64::new(0);

    let name = path.file_name()?.to_str()?;
    let dir = match path.parent() {
        Some(p) if !p.as_os_str().is_empty() => p,
        _ => std::path::Path::new("."),
    };
    Some(dir.join(format!(
        ".{name}.{}.{}.tmp",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    )))
}

impl CsvSink {
    fn new(py: Python, target: Py<PyAny>) -> PyResult<Self> {
        if let Ok(path) = target.extract::<std::path::PathBuf>(py) {
            return Ok(CsvSink::Path { path, open: None });
        }
        if let Ok(write) = target.getattr(py, "write") {
            return Ok(CsvSink::Writer(write));
        }
        Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
            "Argument must be a path (str or os.PathLike) or a file-like object with a 'write' method",
        ))
    }

    /// Open the destination at the first chunk.
    ///
    /// The bytes go to a sibling temporary file that is renamed into place at
    /// the end, so a failure part-way leaves whatever was already at the
    /// target alone and no reader ever sees half a file. Falls back to
    /// writing in place when a temporary cannot stand in for the target.
    fn open_target(path: &std::path::Path) -> PyResult<OpenFile> {
        let existing = look_at(path);
        if !matches!(existing, Existing::Special) {
            if let Some(temp) = temp_path_for(path) {
                if let Ok(file) = std::fs::File::create(&temp) {
                    // Keep the target's mode: the rename brings a new inode,
                    // which would otherwise arrive with the default umask.
                    if let Existing::RegularFile(mode) = existing {
                        let _ = std::fs::set_permissions(&temp, mode);
                    }
                    return Ok(OpenFile {
                        file,
                        temp: Some(temp),
                    });
                }
            }
        }
        Ok(OpenFile {
            file: std::fs::File::create(path).map_err(io_err)?,
            temp: None,
        })
    }

    fn write_out(&mut self, output: &[u8]) -> PyResult<()> {
        if let CsvSink::Path { path, open } = self {
            let f = match open {
                Some(f) => f,
                None => open.insert(Self::open_target(path)?),
            };
            f.file.write_all(output).map_err(io_err)?;
        }
        Ok(())
    }

    /// Write `output` out once it passes a chunk. Callers call this only
    /// between rows, so a row is never split and the per-row rewind still has
    /// the whole row to rewind.
    #[inline]
    fn row_done(&mut self, output: &mut Vec<u8>) -> PyResult<()> {
        if output.len() >= CHUNK && matches!(self, CsvSink::Path { .. }) {
            self.write_out(output)?;
            output.clear();
        }
        Ok(())
    }

    /// Discard everything written so far, for the Arrow path's retry.
    fn restart(&mut self, output: &mut Vec<u8>) -> PyResult<()> {
        output.clear();
        if let CsvSink::Path {
            open: Some(open), ..
        } = self
        {
            open.file.set_len(0).map_err(io_err)?;
            open.file.seek(SeekFrom::Start(0)).map_err(io_err)?;
        }
        Ok(())
    }

    /// Write the tail, then put the file in place.
    ///
    /// The rename is the moment the new file becomes visible, so a reader
    /// sees either the old file or the whole new one.
    fn finish(&mut self, py: Python, output: &[u8]) -> PyResult<()> {
        match self {
            CsvSink::Path { .. } => {
                self.write_out(output)?;
                if let CsvSink::Path {
                    path,
                    open: Some(open),
                } = self
                {
                    if let Some(temp) = open.temp.take() {
                        open.file.flush().map_err(io_err)?;
                        std::fs::rename(&temp, &*path).map_err(|e| {
                            let _ = std::fs::remove_file(&temp);
                            io_err(e)
                        })?;
                    }
                }
                Ok(())
            }
            CsvSink::Writer(write) => {
                write.call1(py, (pyo3::types::PyBytes::new(py, output),))?;
                Ok(())
            }
        }
    }

    /// After an error, leave nothing half-written behind.
    ///
    /// With a temporary file the target is never touched, so only the
    /// temporary goes. Writing in place has no such luxury: the partial file
    /// is removed, because a truncated CSV reads as a complete one with fewer
    /// rows.
    fn abandon(self) {
        if let CsvSink::Path {
            path,
            open: Some(open),
        } = self
        {
            let OpenFile { file, temp } = open;
            drop(file);
            let _ = std::fs::remove_file(temp.as_deref().unwrap_or(&path));
        }
    }
}

/// Write one value through the type cache, falling back to the full cascade.
/// Returns the detected type when the cascade ran, for the caller to cache.
#[inline]
fn write_cell(
    value: &Bound<'_, PyAny>,
    cached: ColType,
    sink: &mut CsvCell<'_>,
) -> PyResult<Option<ColType>> {
    if try_cached(value, cached, sink)? {
        return Ok(None);
    }
    classify_and_write(value, sink).map(Some)
}

/// Turn a failed `try_iter` into a clear message. Skipping the loop instead
/// would write an empty file and report success.
fn not_iterable(_: PyErr) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyTypeError, _>(
        "records must be an iterable of dicts, a DataFrame, or an Arrow stream",
    )
}

/// Name the key a record carries that the header has no column for.
///
/// Only called once a record is known to have an unmatched key.
fn unknown_column(row: &Bound<'_, pyo3::types::PyDict>, headers: &[String], index: usize) -> PyErr {
    for key in row.keys().iter() {
        let Ok(name) = key.extract::<String>() else {
            return PyErr::new::<pyo3::exceptions::PyTypeError, _>("Record keys must be strings");
        };
        if !headers.contains(&name) {
            return value_err(format!(
                "record {index}: no column '{name}'. The header comes from the first record \
                 ({}), and writing this one would put more fields on the line than the header \
                 has — give every record the same keys, or name the columns you want.",
                headers.join(", ")
            ));
        }
    }
    value_err(format!("record {index}: keys do not match the header"))
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

    // Refuse a repeated column name with the same sentence the Excel writer
    // uses. Checked here because the Arrow attempt below reaches pyarrow
    // first, which raises its own differently worded error for the same
    // mistake. Non-string column labels (pandas allows them) are left to the
    // paths that know what to do with them.
    let bound = records.bind(py);
    if let Ok(labels) = bound.getattr("columns") {
        if let Ok(names) = labels.extract::<Vec<String>>() {
            crate::helpers::reject_duplicate_headers(&names)?;
        }
    }

    let mut target = CsvSink::new(py, file_name)?;
    let mut output: Vec<u8> = Vec::with_capacity(CHUNK + 4096);
    let reps = NumReps {
        na: na_rep.as_deref(),
        inf: inf_value.as_deref(),
    };
    let written = write_csv_rows(
        py,
        &records,
        &mut target,
        &mut output,
        delim_byte,
        sanitize_formulas,
        bom,
        columns.as_ref(),
        header,
        reps,
    )
    .and_then(|()| target.finish(py, &output));
    if written.is_err() {
        target.abandon();
    }
    written
}

/// The body of [`write_csv`]: everything between choosing the target and
/// finishing it. `output` holds the rows not yet passed to `target`.
#[allow(clippy::too_many_arguments)]
fn write_csv_rows(
    py: Python,
    records: &Py<PyAny>,
    target: &mut CsvSink,
    output: &mut Vec<u8>,
    delim_byte: u8,
    sanitize_formulas: bool,
    bom: bool,
    selection: Option<&Vec<String>>,
    header: bool,
    reps: NumReps<'_>,
) -> PyResult<()> {
    const BOM: &[u8] = &[0xEF, 0xBB, 0xBF];
    let bound = records.bind(py);
    if bom {
        output.extend_from_slice(BOM);
    }

    // Fast path: Arrow zero-copy if the object exposes `__arrow_c_stream__`
    // (Pandas ≥2.0, Polars). Falls back to the per-object paths below on
    // failure (e.g. empty Null-typed columns).
    if bound.hasattr("__arrow_c_stream__")? {
        match write_csv_via_arrow(
            records,
            py,
            target,
            output,
            delim_byte,
            sanitize_formulas,
            selection,
            header,
            reps,
        ) {
            Ok(()) => return Ok(()),
            // A bad `columns` is the caller's mistake, not a quirk of this
            // path, so it must not fall through to a slower one that would
            // raise the same error later.
            Err(e) if e.is_instance_of::<pyo3::exceptions::PyValueError>(py) => return Err(e),
            Err(_) => {}
        }
        // Part of the file may already be on disk, so the restart has to
        // reach the file as well as the buffer.
        target.restart(output)?;
        if bom {
            output.extend_from_slice(BOM);
        }
    }

    if bound.hasattr("columns")? {
        let all_columns: Vec<String> = bound.getattr("columns")?.extract()?;
        let indices = select_indices(&all_columns, selection)?;
        let columns = apply_selection(all_columns, indices.as_ref());
        if header {
            write_csv_row_strings(output, &columns, delim_byte, sanitize_formulas);
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
                    let mut sink = CsvCell::new(output, sanitize_formulas, reps);
                    classify_and_write(&item, &mut sink)
                        .map_err(|e| located(e, &item, || Loc::Column(&columns[i])))?;
                }
                output.push(b'\n');
                target.row_done(output)?;
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
                            let mut sink = CsvCell::new(output, sanitize_formulas, reps);
                            classify_and_write(&item, &mut sink)
                                .map_err(|e| located(e, &item, || Loc::Column(&columns[n])))?;
                        }
                    }
                    None => {
                        for (n, item_res) in row.try_iter()?.enumerate() {
                            let item = item_res?;
                            if n > 0 {
                                output.push(delim_byte);
                            }
                            let mut sink = CsvCell::new(output, sanitize_formulas, reps);
                            classify_and_write(&item, &mut sink).map_err(|e| {
                                located(e, &item, || {
                                    Loc::Column(columns.get(n).map_or("?", String::as_str))
                                })
                            })?;
                        }
                    }
                }
                output.push(b'\n');
                target.row_done(output)?;
            }
        }
    } else {
        // Records path (list of dicts / generator). First-row type cache
        // mirrors the Excel Records path — skips the full type cascade
        // after the first row when the column's Python type is stable.
        let mut headers: Vec<String> = Vec::new();
        let mut headers_written = false;
        let mut col_types: Vec<ColType> = Vec::new();
        // The first record's key objects, for the pointer check below.
        let mut header_keys: Vec<Py<pyo3::types::PyString>> = Vec::new();

        let rows: pyo3::Bound<'_, pyo3::types::PyIterator> =
            bound.try_iter().map_err(not_iterable)?;
        for (index, row_res) in rows.enumerate() {
            let row_obj = row_res?;
            let row_dict = row_obj.cast::<pyo3::types::PyDict>().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "Items in records must be dictionaries",
                )
            })?;

            if !headers_written {
                let mut keys: Vec<String> = Vec::new();
                for key in row_dict.keys().iter() {
                    let key = key.cast_into::<pyo3::types::PyString>().map_err(|_| {
                        PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                            "Record keys must be strings",
                        )
                    })?;
                    keys.push(key.to_str()?.to_owned());
                    header_keys.push(key.unbind());
                }
                // Validated against the first row, which is the only row whose
                // keys are known before the stream is consumed.
                headers = apply_selection(keys.clone(), select_indices(&keys, selection)?.as_ref());
                if header {
                    write_csv_row_strings(output, &headers, delim_byte, sanitize_formulas);
                }
                col_types.resize(headers.len(), ColType::Unknown);
                headers_written = true;
            }

            // Fast path: a record reusing the first one's key objects in the
            // same order is written from its own iteration order. One pointer
            // comparison per key, and the row is rewound if any of them fails.
            let row_start = output.len();
            let fast = selection.is_none() && row_dict.len() == headers.len() && {
                let mut ok = true;
                for (col, (key, value)) in row_dict.iter().enumerate() {
                    if !header_keys.get(col).is_some_and(|k| k.is(&key)) {
                        ok = false;
                        break;
                    }
                    if col > 0 {
                        output.push(delim_byte);
                    }
                    let cached = col_types[col];
                    let mut sink = CsvCell::new(output, sanitize_formulas, reps);
                    let detected = write_cell(&value, cached, &mut sink).map_err(|e| {
                        located(e, &value, || Loc::Record {
                            index,
                            key: &headers[col],
                        })
                    })?;
                    if let Some(detected) = detected {
                        if col_types[col] == ColType::Unknown {
                            col_types[col] = detected;
                        }
                    }
                }
                ok
            };
            if fast {
                output.push(b'\n');
                target.row_done(output)?;
                continue;
            }
            output.truncate(row_start);

            // Otherwise walk the header and look each column up by name, so a
            // reordered record still lands in the right fields and an extra key
            // cannot lengthen the line.
            let mut matched = 0usize;
            for (col, name) in headers.iter().enumerate() {
                if col > 0 {
                    output.push(delim_byte);
                }
                let mut sink = CsvCell::new(output, sanitize_formulas, reps);
                // A missing key writes an empty field, keeping the field count.
                if let Some(value) = row_dict.get_item(name)? {
                    matched += 1;
                    let cached = col_types.get(col).copied().unwrap_or(ColType::Unknown);
                    let detected = write_cell(&value, cached, &mut sink)
                        .map_err(|e| located(e, &value, || Loc::Record { index, key: name }))?;
                    if let Some(detected) = detected {
                        if col_types[col] == ColType::Unknown {
                            col_types[col] = detected;
                        }
                    }
                }
            }
            // Header names are unique, so more keys than matches means one has
            // no column. With `columns` the header is a chosen subset and the
            // rest are dropped on purpose.
            if selection.is_none() && matched != row_dict.len() {
                return Err(unknown_column(row_dict, &headers, index));
            }
            output.push(b'\n');
            target.row_done(output)?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_csv_via_arrow(
    records: &Py<PyAny>,
    py: Python,
    target: &mut CsvSink,
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
    // Raises ValueError, which the caller does not retry on the slower paths.
    crate::arrow_writer::check_schema(&schema, indices.as_deref())?;
    if header {
        let out = apply_selection(headers, indices.as_ref());
        write_csv_row_strings(output, &out, delim, sanitize);
    }

    // A frame built in one go is often a single batch, so the chunking has
    // to happen per row inside it, not per batch.
    let mut row_done = |out: &mut Vec<u8>| target.row_done(out);
    for batch_result in reader {
        let batch = batch_result.map_err(crate::arrow_ffi::batch_read_err)?;
        // Projection is an Arc clone per column, so the zero-copy path stays
        // zero-copy.
        let batch = match &indices {
            Some(ix) => batch
                .project(ix)
                .map_err(crate::arrow_ffi::batch_read_err)?,
            None => batch,
        };
        crate::arrow_writer::write_arrow_batch_csv(
            output,
            &batch,
            delim,
            sanitize,
            reps,
            &mut row_done,
        )?;
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

/// ISO-8601, as `datetime.isoformat()` writes it: microseconds only when
/// there are some, and the UTC offset when the value is aware, so that the
/// text names one instant rather than a wall-clock time in an unknown zone.
///
/// A year the fields cannot hold — a pandas `Timestamp` past 9999 — raises;
/// see [`crate::helpers::true_year`].
fn emit_datetime(output: &mut Vec<u8>, dt: &Bound<PyDateTime>) -> PyResult<()> {
    let year = dt.get_year();
    if !crate::helpers::is_builtin_date(dt.as_any()) {
        check_subclass_year(dt, year)?;
    }
    crate::helpers::write_iso_date(
        output,
        year as u32,
        u32::from(dt.get_month()),
        u32::from(dt.get_day()),
    );
    output.push(b'T');
    crate::helpers::write_iso_time(
        output,
        u32::from(dt.get_hour()),
        u32::from(dt.get_minute()),
        u32::from(dt.get_second()),
        dt.get_microsecond() * 1_000,
    );
    // Naive values, the common case, stop at a pointer check.
    if dt.get_tzinfo().is_some() {
        emit_utc_offset(output, &dt.call_method0("utcoffset")?)?;
    }
    Ok(())
}

/// The slow half of the year check in [`emit_datetime`], for a `datetime`
/// subclass. Kept out of line so that the common case, a built-in datetime,
/// pays one pointer comparison and nothing else.
#[cold]
#[inline(never)]
fn check_subclass_year(dt: &Bound<PyDateTime>, year: i32) -> PyResult<()> {
    if crate::helpers::true_year(dt.as_any(), year)? != year {
        // The caller knows where the value came from and says so through
        // `located`, which rebuilds this message.
        return Err(crate::helpers::iso_range_err(
            Loc::Unknown,
            &crate::helpers::iso_text(dt.as_any()),
            crate::helpers::WRONG_YEAR_NOTE,
        ));
    }
    Ok(())
}

/// The same for a `datetime.time`.
fn emit_time(output: &mut Vec<u8>, t: &Bound<PyTime>) -> PyResult<()> {
    crate::helpers::write_iso_time(
        output,
        u32::from(t.get_hour()),
        u32::from(t.get_minute()),
        u32::from(t.get_second()),
        t.get_microsecond() * 1_000,
    );
    if t.get_tzinfo().is_some() {
        emit_utc_offset(output, &t.call_method0("utcoffset")?)?;
    }
    Ok(())
}

/// Append `+HH:MM` (with `:SS` when the offset has seconds) for the result of
/// `utcoffset()`. `None` — a tzinfo that declines to give an offset — adds
/// nothing, which is also what `isoformat()` does.
fn emit_utc_offset(output: &mut Vec<u8>, offset: &Bound<PyAny>) -> PyResult<()> {
    let Ok(delta) = offset.cast::<PyDelta>() else {
        return Ok(());
    };
    let total = i64::from(delta.get_days()) * 86_400 + i64::from(delta.get_seconds());
    crate::helpers::write_iso_offset(output, total);
    Ok(())
}

fn emit_date(output: &mut Vec<u8>, d: &Bound<PyDate>) {
    crate::helpers::write_iso_date(
        output,
        d.get_year() as u32,
        u32::from(d.get_month()),
        u32::from(d.get_day()),
    );
}

/// [`CellWriter`] sink that appends one Python scalar to a CSV byte buffer.
/// The type-detection order lives in [`crate::cell`]; this only encodes the
/// per-type CSV serialization (and the optional formula-injection guard for
/// strings).
///
/// A value it cannot write raises without saying where it came from; the
/// caller adds that with [`crate::helpers::located`], on the error path only.
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
        emit_datetime(self.output, dt)
    }

    fn write_date(&mut self, d: &Bound<'_, PyDate>) -> PyResult<()> {
        emit_date(self.output, d);
        Ok(())
    }

    fn write_time(&mut self, t: &Bound<'_, PyTime>) -> PyResult<()> {
        emit_time(self.output, t)
    }

    /// `str()` of a Decimal is its exact text, so `Decimal(text)` reads back
    /// the same value, trailing zeros and exponent included: `1.10` stays
    /// `1.10`. NaN and the infinities take `na_rep` and `inf_value`, as a
    /// float's do.
    fn write_decimal(&mut self, d: &Bound<'_, PyAny>) -> PyResult<()> {
        let text = d.str()?;
        let text = text.to_str()?;
        // Covers `NaN`, `-NaN`, `sNaN` and `-sNaN`.
        if text.ends_with("NaN") {
            return self.write_float(f64::NAN);
        }
        if text.ends_with("Infinity") {
            let inf = if text.starts_with('-') {
                f64::NEG_INFINITY
            } else {
                f64::INFINITY
            };
            return self.write_float(inf);
        }
        self.output.extend_from_slice(text.as_bytes());
        Ok(())
    }
}
