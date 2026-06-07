use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateTime, PyInt};
use pyo3::{Py, PyAny, Python};
use rust_xlsxwriter::{Format, Workbook};
use std::collections::HashSet;

use crate::cell::{classify_and_write, try_cached, CellWriter};
use crate::data_types::{FreezePanesConfig, WorksheetData};
use crate::helpers::{
    py_date_to_excel, py_datetime_to_excel, save_workbook, write_all_headers, write_num,
    write_number_opt, ColType,
};
use crate::utils::ensure_valid_sheet_name;

pub fn xlsx_err(e: impl std::fmt::Display) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("Excel write error: {}", e))
}

/// [`CellWriter`] sink that writes one Python scalar to an Excel cell.
///
/// `col_override` is an explicit per-column `Format` (from `column_formats`).
/// When `Some`, it wins over `float_fmt` for numeric cells and over
/// `datetime_fmt` for datetime cells. The detection order lives in
/// [`crate::cell`]; this only encodes the per-type Excel action.
struct ExcelCell<'a> {
    worksheet: &'a mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    float_fmt: Option<&'a Format>,
    datetime_fmt: &'a Format,
    datetime_cols_set: &'a mut HashSet<u16>,
    col_override: Option<&'a Format>,
}

impl ExcelCell<'_> {
    /// Set the datetime column format once, on first datetime cell in the column.
    fn ensure_datetime_format(&mut self) -> PyResult<()> {
        let fmt = self.col_override.unwrap_or(self.datetime_fmt);
        if self.datetime_cols_set.insert(self.col) {
            self.worksheet
                .set_column_format(self.col, fmt)
                .map_err(xlsx_err)?;
        }
        Ok(())
    }
}

impl CellWriter for ExcelCell<'_> {
    fn write_none(&mut self) -> PyResult<()> {
        self.worksheet
            .write_string(self.row, self.col, "")
            .map_err(xlsx_err)?;
        Ok(())
    }

    fn write_str(&mut self, s: &str) -> PyResult<()> {
        self.worksheet
            .write_string(self.row, self.col, s)
            .map_err(xlsx_err)?;
        Ok(())
    }

    fn write_bool(&mut self, b: bool) -> PyResult<()> {
        self.worksheet
            .write_boolean(self.row, self.col, b)
            .map_err(xlsx_err)?;
        Ok(())
    }

    fn write_float(&mut self, f: f64) -> PyResult<()> {
        write_num(
            self.worksheet,
            self.row,
            self.col,
            f,
            self.col_override.or(self.float_fmt),
        )
    }

    fn write_int(&mut self, i: &Bound<'_, PyInt>) -> PyResult<()> {
        let val: f64 = i.extract()?;
        write_number_opt(self.worksheet, self.row, self.col, val, self.col_override)
    }

    fn write_datetime(&mut self, dt: &Bound<'_, PyDateTime>) -> PyResult<()> {
        self.ensure_datetime_format()?;
        let excel_dt = py_datetime_to_excel(dt)?;
        self.worksheet
            .write_datetime(self.row, self.col, &excel_dt)
            .map_err(xlsx_err)?;
        Ok(())
    }

    fn write_date(&mut self, d: &Bound<'_, PyDate>) -> PyResult<()> {
        self.ensure_datetime_format()?;
        let excel_dt = py_date_to_excel(d)?;
        self.worksheet
            .write_datetime(self.row, self.col, &excel_dt)
            .map_err(xlsx_err)?;
        Ok(())
    }
}

/// Per-column scalar classification shared by the Pandas and Polars writers.
/// Resolved once per column (outside the row loop) to skip per-cell dtype
/// probing.
#[derive(Copy, Clone, PartialEq, Eq)]
enum ScalarKind {
    Int,
    Float,
    Bool,
    Temporal,
    Other,
}

/// Classify a Polars dtype from its stringified form. Cheaper than three
/// `call_method0("is_*")` Python round-trips per column.
fn polars_kind(dtype_str: &str) -> ScalarKind {
    if dtype_str.starts_with("Int") || dtype_str.starts_with("UInt") {
        ScalarKind::Int
    } else if dtype_str.starts_with("Float") || dtype_str == "Decimal" {
        ScalarKind::Float
    } else if dtype_str == "Boolean" {
        ScalarKind::Bool
    } else if dtype_str.starts_with("Date")
        || dtype_str.starts_with("Datetime")
        || dtype_str.starts_with("Time")
        || dtype_str.starts_with("Duration")
    {
        ScalarKind::Temporal
    } else {
        ScalarKind::Other
    }
}

#[allow(clippy::too_many_arguments)]
fn write_worksheet_content(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    records: &WorksheetData,
    password: Option<&String>,
    freeze_row: Option<u32>,
    freeze_col: Option<u16>,
    float_format: Option<&String>,
    datetime_format: Option<&String>,
    index_columns: Option<&Vec<String>>,
    autofit: bool,
    bold_headers: bool,
    column_width: Option<f64>,
    column_widths: Option<&Bound<'_, PyAny>>,
    column_formats: Option<&Bound<'_, PyAny>>,
    header_format: Option<&crate::format::Format>,
    py: Python,
) -> PyResult<()> {
    let float_fmt = float_format.map(|s| Format::new().set_num_format(s));
    let dt_fmt_str = datetime_format
        .map(|s| s.as_str())
        .unwrap_or("yyyy-mm-ddThh:mm:ss");
    let datetime_fmt = Format::new().set_num_format(dt_fmt_str);
    let bold_fmt = Format::new().set_bold();
    let mut datetime_cols_set: HashSet<u16> = HashSet::new();
    let mut final_headers: Vec<String> = Vec::new();

    match records {
        WorksheetData::ArrowDataFrame(stream_obj) => {
            let arrow_ok = (|| -> PyResult<()> {
                let reader = crate::arrow_ffi::stream_to_reader(stream_obj, py)?;

                let schema = reader.schema();
                final_headers = schema
                    .fields()
                    .iter()
                    .map(|f| f.name().to_string())
                    .collect();
                write_all_headers(
                    worksheet,
                    &final_headers,
                    bold_headers,
                    &bold_fmt,
                    index_columns,
                    header_format.map(|h| &h.inner),
                )?;

                let mut current_row: u32 = 1;
                let mut formats_set = false;

                // Resolve per-column formats ONCE (after headers are known).
                // Apply set_column_format BEFORE the first batch — constant memory mode requires
                // column formats to be set before their data rows.
                let col_formats: Vec<Option<crate::format::Format>> =
                    crate::format::resolve_column_formats(column_formats, &final_headers, py)?;

                for batch_result in reader {
                    let batch = batch_result.map_err(crate::arrow_ffi::batch_read_err)?;

                    if !formats_set {
                        // Auto datetime column formats first…
                        crate::arrow_writer::set_datetime_column_formats(
                            worksheet,
                            &batch,
                            &datetime_fmt,
                        )?;
                        // …then explicit column_formats override them (and any other cols).
                        crate::format::apply_column_formats(worksheet, &col_formats)?;
                        formats_set = true;
                    }

                    crate::arrow_writer::write_arrow_batch(
                        worksheet,
                        &batch,
                        current_row,
                        float_fmt.as_ref(),
                        &col_formats,
                    )?;

                    current_row += batch.num_rows() as u32;
                }
                Ok(())
            })();

            // If Arrow FFI failed (e.g. Null-typed empty DataFrame), at
            // least write the header row from `.columns`.
            if arrow_ok.is_err() {
                if let Ok(cols) = stream_obj.getattr(py, "columns") {
                    final_headers = cols.extract(py).unwrap_or_default();
                    write_all_headers(
                        worksheet,
                        &final_headers,
                        bold_headers,
                        &bold_fmt,
                        index_columns,
                        header_format.map(|h| &h.inner),
                    )?;
                }
            }
        }

        WorksheetData::Records(records_list) => {
            if let Ok(rows) = records_list.bind(py).try_iter() {
                let mut headers: Vec<String> = Vec::new();
                let mut headers_written = false;
                let mut col_types: Vec<ColType> = Vec::new();
                // Resolved once when headers are first seen; kept alive for the
                // entire row loop so we can hand out &Format borrows per cell.
                let mut col_formats: Vec<Option<crate::format::Format>> = Vec::new();

                for (row_idx, row_res) in rows.enumerate() {
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
                        write_all_headers(
                            worksheet,
                            &headers,
                            bold_headers,
                            &bold_fmt,
                            index_columns,
                            header_format.map(|h| &h.inner),
                        )?;
                        col_types.resize(headers.len(), ColType::Unknown);
                        final_headers = headers.clone();
                        // Resolve per-column formats ONCE and keep them for the whole loop.
                        // Apply set_column_format BEFORE writing data rows (constant memory mode).
                        col_formats =
                            crate::format::resolve_column_formats(column_formats, &final_headers, py)?;
                        crate::format::apply_column_formats(worksheet, &col_formats)?;
                        headers_written = true;
                    }

                    let row_u32 = (row_idx + 1) as u32;
                    // Iterate the dict directly (insertion order == header order)
                    // to avoid allocating a fresh `values()` list per row.
                    for (col, (_key, value)) in row_dict.iter().enumerate() {
                        let col_u16 = col as u16;
                        let cached = col_types
                            .get(col)
                            .copied()
                            .unwrap_or(ColType::Unknown);

                        // Column format override: wins over float_fmt / datetime_fmt.
                        let col_override = crate::format::col_override(&col_formats, col);

                        let mut sink = ExcelCell {
                            worksheet: &mut *worksheet,
                            row: row_u32,
                            col: col_u16,
                            float_fmt: float_fmt.as_ref(),
                            datetime_fmt: &datetime_fmt,
                            datetime_cols_set: &mut datetime_cols_set,
                            col_override,
                        };

                        if !try_cached(&value, cached, &mut sink)? {
                            let detected = classify_and_write(&value, &mut sink)?;
                            if col < col_types.len() && col_types[col] == ColType::Unknown {
                                col_types[col] = detected;
                            }
                        }
                    }
                }
            }
        }

        WorksheetData::PandasDataFrame(df) => {
            write_dataframe(
                worksheet,
                py,
                df,
                &mut final_headers,
                column_formats,
                float_fmt.as_ref(),
                &datetime_fmt,
                &mut datetime_cols_set,
                bold_headers,
                &bold_fmt,
                index_columns,
                header_format,
                "__getitem__",
                "tolist",
                |dtype| {
                    let kind: String = dtype.getattr("kind")?.extract()?;
                    Ok(map_pandas_kind(kind.chars().next().unwrap_or('O')))
                },
            )?;
        }

        WorksheetData::PolarsDataFrame(df) => {
            write_dataframe(
                worksheet,
                py,
                df,
                &mut final_headers,
                column_formats,
                float_fmt.as_ref(),
                &datetime_fmt,
                &mut datetime_cols_set,
                bold_headers,
                &bold_fmt,
                index_columns,
                header_format,
                "get_column",
                "to_list",
                |dtype| Ok(polars_kind(&dtype.to_string())),
            )?;
        }
    }

    if freeze_row.is_some() || freeze_col.is_some() {
        worksheet
            .set_freeze_panes(freeze_row.unwrap_or(0), freeze_col.unwrap_or(0))
            .map_err(xlsx_err)?;
    }

    if autofit {
        worksheet.autofit();
    }
    crate::helpers::apply_column_widths(
        worksheet,
        &final_headers,
        column_width,
        column_widths,
        py,
    )?;

    if let Some(password) = password {
        worksheet.protect_with_password(password);
    }

    Ok(())
}

fn map_pandas_kind(kind: char) -> ScalarKind {
    match kind {
        'i' | 'u' => ScalarKind::Int,
        'f' => ScalarKind::Float,
        'b' => ScalarKind::Bool,
        'M' => ScalarKind::Temporal,
        _ => ScalarKind::Other,
    }
}

/// Per-column accessor. `.tolist()`/`.to_list()` return Python lists, so
/// the fast path goes through `PyList` directly (`PyList_GET_ITEM`); the
/// generic branch is kept only for non-list iterables.
enum ColAccess<'py> {
    List(Bound<'py, pyo3::types::PyList>),
    Generic(Bound<'py, PyAny>),
}

impl<'py> ColAccess<'py> {
    fn get(&self, row: usize) -> PyResult<Bound<'py, PyAny>> {
        match self {
            ColAccess::List(list) => list.get_item(row),
            ColAccess::Generic(obj) => obj.get_item(row),
        }
    }
}

/// Shared row writer for Pandas and Polars paths. Both frameworks provide
/// per-column Python lists plus a per-column kind — the row-level dispatch
/// is identical, so a single helper handles both.
///
/// `col_formats` is a per-column slice of optional `Format`s resolved from
/// `column_formats`. When `Some`, the column format wins over `float_fmt` for
/// numeric/float columns, and is applied to string columns as well. Pass an
/// empty slice when no overrides are needed (byte-identical behaviour).
#[allow(clippy::too_many_arguments)]
fn write_df_rows<F>(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    py: Python,
    nrows: usize,
    col_lists: &[Py<PyAny>],
    kind_at: F,
    float_fmt: Option<&Format>,
    datetime_fmt: &Format,
    datetime_cols_set: &mut HashSet<u16>,
    col_formats: &[Option<crate::format::Format>],
) -> PyResult<()>
where
    F: Fn(usize) -> ScalarKind,
{
    let bound_cols: Vec<ColAccess> = col_lists
        .iter()
        .map(|c| {
            let b = c.bind(py).clone();
            match b.cast_into::<pyo3::types::PyList>() {
                Ok(list) => ColAccess::List(list),
                Err(e) => ColAccess::Generic(e.into_inner()),
            }
        })
        .collect();

    // Per-column format override is fixed for the whole column — resolve once
    // instead of per cell.
    let col_overrides: Vec<Option<&Format>> = (0..bound_cols.len())
        .map(|c| crate::format::col_override(col_formats, c))
        .collect();

    for row in 0..nrows {
        let row_u32 = (row + 1) as u32;
        for (col_idx, col_list) in bound_cols.iter().enumerate() {
            let col_u16 = col_idx as u16;
            let item = col_list.get(row)?;
            let col_override = col_overrides[col_idx];

            if item.is_none() {
                worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                continue;
            }

            match kind_at(col_idx) {
                ScalarKind::Int => {
                    let val: f64 = item.extract()?;
                    write_number_opt(worksheet, row_u32, col_u16, val, col_override)?;
                }
                ScalarKind::Float => {
                    let val: f64 = item.extract()?;
                    write_num(worksheet, row_u32, col_u16, val, col_override.or(float_fmt))?;
                }
                ScalarKind::Bool => {
                    let val: bool = item.extract()?;
                    worksheet
                        .write_boolean(row_u32, col_u16, val)
                        .map_err(xlsx_err)?;
                }
                ScalarKind::Temporal => {
                    if let Ok(dt) = item.cast::<PyDateTime>() {
                        let excel_dt = py_datetime_to_excel(&dt)?;
                        worksheet
                            .write_datetime(row_u32, col_u16, &excel_dt)
                            .map_err(xlsx_err)?;
                    } else if let Ok(d) = item.cast::<PyDate>() {
                        let excel_dt = py_date_to_excel(&d)?;
                        worksheet
                            .write_datetime(row_u32, col_u16, &excel_dt)
                            .map_err(xlsx_err)?;
                    } else {
                        worksheet
                            .write_string(row_u32, col_u16, item.to_string())
                            .map_err(xlsx_err)?;
                    }
                }
                ScalarKind::Other => {
                    let mut sink = ExcelCell {
                        worksheet: &mut *worksheet,
                        row: row_u32,
                        col: col_u16,
                        float_fmt,
                        datetime_fmt,
                        datetime_cols_set: &mut *datetime_cols_set,
                        col_override,
                    };
                    classify_and_write(&item, &mut sink)?;
                }
            }
        }
    }
    Ok(())
}

/// Shared Pandas/Polars writer. Both expose `columns`, `dtypes`, `__len__`
/// and per-column series access; they differ only in the series accessor
/// method names and how a dtype maps to a [`ScalarKind`]. Those three points
/// are passed in so the surrounding orchestration lives in one place.
///
/// - `get_column_method`: `"__getitem__"` (Pandas) or `"get_column"` (Polars)
/// - `to_list_method`: `"tolist"` (Pandas) or `"to_list"` (Polars)
/// - `classify_dtype`: maps one dtype object to a `ScalarKind`
#[allow(clippy::too_many_arguments)]
fn write_dataframe<C>(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    py: Python,
    df: &Py<PyAny>,
    final_headers: &mut Vec<String>,
    column_formats: Option<&Bound<'_, PyAny>>,
    float_fmt: Option<&Format>,
    datetime_fmt: &Format,
    datetime_cols_set: &mut HashSet<u16>,
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
    header_format: Option<&crate::format::Format>,
    get_column_method: &str,
    to_list_method: &str,
    classify_dtype: C,
) -> PyResult<()>
where
    C: Fn(&Bound<'_, PyAny>) -> PyResult<ScalarKind>,
{
    let headers: Vec<String> = df.getattr(py, "columns")?.extract(py)?;
    *final_headers = headers.clone();
    let dtypes = df.getattr(py, "dtypes")?;
    let dtypes_list: Vec<Bound<'_, PyAny>> =
        dtypes.bind(py).try_iter()?.collect::<Result<Vec<_>, _>>()?;

    write_all_headers(
        worksheet,
        &headers,
        bold_headers,
        bold_fmt,
        index_columns,
        header_format.map(|h| &h.inner),
    )?;

    let mut col_kinds: Vec<ScalarKind> = Vec::with_capacity(headers.len());
    let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(headers.len());

    // NOTE: `to_list`/`tolist` materializes each column as a full Python list,
    // so peak memory here is O(rows) per column — this path is NOT constant
    // memory. It is only reached when the Arrow zero-copy path is unavailable
    // (old pandas without `__arrow_c_stream__`, or exotic dtypes). Modern
    // pandas ≥2 and Polars hit the Arrow path in `data_types.rs` instead.
    for (col_idx, header) in headers.iter().enumerate() {
        col_kinds.push(classify_dtype(&dtypes_list[col_idx])?);
        let col_series = df.call_method1(py, get_column_method, (header.as_str(),))?;
        col_lists.push(col_series.call_method0(py, to_list_method)?);
    }

    let nrows: usize = df.call_method0(py, "__len__")?.extract(py)?;

    // Auto datetime column formats first, then explicit column_formats
    // override (constant memory: BEFORE writing data rows).
    for (col_idx, kind) in col_kinds.iter().enumerate() {
        if *kind == ScalarKind::Temporal && datetime_cols_set.insert(col_idx as u16) {
            worksheet
                .set_column_format(col_idx as u16, datetime_fmt)
                .map_err(xlsx_err)?;
        }
    }
    let col_formats: Vec<Option<crate::format::Format>> =
        crate::format::resolve_column_formats(column_formats, final_headers, py)?;
    crate::format::apply_column_formats(worksheet, &col_formats)?;

    write_df_rows(
        worksheet,
        py,
        nrows,
        &col_lists,
        |col_idx| col_kinds[col_idx],
        float_fmt,
        datetime_fmt,
        datetime_cols_set,
        &col_formats,
    )
}

/// Resolve a per-sheet value from a dict keyed by sheet name, falling back
/// to the `"general"` key. Returns the matching Python value, if any.
fn keyed_get<'py>(
    dict: Option<&Bound<'py, pyo3::types::PyDict>>,
    sheet: &str,
) -> PyResult<Option<Bound<'py, PyAny>>> {
    let Some(dict) = dict else { return Ok(None) };
    if let Some(v) = dict.get_item(sheet)? {
        return Ok(Some(v));
    }
    dict.get_item("general")
}

#[allow(clippy::too_many_arguments)]
#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false, column_width = None, column_widths = None, column_formats = None, header_format = None))]
pub fn write_worksheets(
    py: Python,
    records_with_sheet_name: Vec<(String, WorksheetData)>,
    file_name: Py<PyAny>,
    password: Option<String>,
    freeze_panes: Option<FreezePanesConfig>,
    float_format: Option<String>,
    datetime_format: Option<String>,
    index_columns: Option<Vec<String>>,
    autofit: bool,
    bold_headers: bool,
    column_width: Option<Bound<'_, pyo3::types::PyDict>>,
    column_widths: Option<Bound<'_, pyo3::types::PyDict>>,
    column_formats: Option<Bound<'_, pyo3::types::PyDict>>,
    header_format: Option<Bound<'_, pyo3::types::PyDict>>,
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    for (sheet_name, records) in records_with_sheet_name {
        ensure_valid_sheet_name(&sheet_name)?;

        let mut worksheet = workbook.add_worksheet_with_constant_memory();
        worksheet.set_name(&sheet_name).map_err(xlsx_err)?;

        let pane = freeze_panes
            .as_ref()
            .map(|c| c.resolve(&sheet_name))
            .unwrap_or_default();

        let sheet_uniform: Option<f64> = keyed_get(column_width.as_ref(), &sheet_name)?
            .map(|v| v.extract())
            .transpose()?;
        let sheet_spec: Option<Bound<'_, PyAny>> =
            keyed_get(column_widths.as_ref(), &sheet_name)?;

        let sheet_col_fmts: Option<Bound<'_, PyAny>> =
            keyed_get(column_formats.as_ref(), &sheet_name)?;
        let sheet_hdr_fmt: Option<crate::format::Format> =
            match keyed_get(header_format.as_ref(), &sheet_name)? {
                Some(v) => Some(v.extract()?),
                None => None,
            };

        write_worksheet_content(
            &mut worksheet,
            &records,
            password.as_ref(),
            pane.row,
            pane.col,
            float_format.as_ref(),
            datetime_format.as_ref(),
            index_columns.as_ref(),
            autofit,
            bold_headers,
            sheet_uniform,
            sheet_spec.as_ref(),
            sheet_col_fmts.as_ref(),
            sheet_hdr_fmt.as_ref(),
            py,
        )?;
    }

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None, freeze_row = None, freeze_col = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false, column_width = None, column_widths = None, column_formats = None, header_format = None))]
pub fn write_worksheet(
    py: Python,
    records: WorksheetData,
    file_name: Py<PyAny>,
    sheet_name: Option<String>,
    password: Option<String>,
    freeze_row: Option<u32>,
    freeze_col: Option<u16>,
    float_format: Option<String>,
    datetime_format: Option<String>,
    index_columns: Option<Vec<String>>,
    autofit: bool,
    bold_headers: bool,
    column_width: Option<f64>,
    column_widths: Option<Bound<'_, PyAny>>,
    column_formats: Option<Bound<'_, PyAny>>,
    header_format: Option<Bound<'_, crate::format::Format>>,
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    let mut worksheet = workbook.add_worksheet_with_constant_memory();

    if let Some(sheet_name) = sheet_name {
        ensure_valid_sheet_name(&sheet_name)?;
        worksheet.set_name(sheet_name).map_err(xlsx_err)?;
    }

    let hdr_borrow = header_format.as_ref().map(|h| h.borrow());
    let hdr_ref = hdr_borrow.as_deref();

    write_worksheet_content(
        &mut worksheet,
        &records,
        password.as_ref(),
        freeze_row,
        freeze_col,
        float_format.as_ref(),
        datetime_format.as_ref(),
        index_columns.as_ref(),
        autofit,
        bold_headers,
        column_width,
        column_widths.as_ref(),
        column_formats.as_ref(),
        hdr_ref,
        py,
    )?;

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}
