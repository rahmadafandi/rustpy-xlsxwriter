use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateTime, PyInt};
use pyo3::{Py, PyAny, Python};
use rust_xlsxwriter::{Format, Workbook};
use std::collections::HashSet;

use crate::cell::{classify_and_write, try_cached, CellWriter};
use crate::data_types::{FreezePanesConfig, WorksheetData};
use crate::helpers::{
    py_date_to_excel, py_datetime_to_excel, save_workbook, write_all_headers, write_bool_opt,
    write_datetime_opt, write_num, write_number_opt, write_string_opt, write_url_or_text,
    ColType,
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
    /// Format for cells that otherwise carry none (strings, bools, blanks).
    /// Only set on banded rows.
    text_fmt: Option<&'a Format>,
    float_fmt: Option<&'a Format>,
    datetime_fmt: &'a Format,
    datetime_cols_set: &'a mut HashSet<u16>,
    col_override: Option<&'a Format>,
    /// When banding is on, datetimes are formatted per cell instead of via a
    /// one-shot column format — a column format is fixed for every row and so
    /// cannot alternate.
    per_cell_datetime: bool,
    /// This column was listed in `url_columns`, so strings become links.
    is_url: bool,
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

    fn put_datetime(&mut self, excel_dt: &rust_xlsxwriter::ExcelDateTime) -> PyResult<()> {
        if self.per_cell_datetime {
            let fmt = self.col_override.unwrap_or(self.datetime_fmt);
            self.worksheet
                .write_datetime_with_format(self.row, self.col, excel_dt, fmt)
                .map_err(xlsx_err)?;
        } else {
            self.ensure_datetime_format()?;
            self.worksheet
                .write_datetime(self.row, self.col, excel_dt)
                .map_err(xlsx_err)?;
        }
        Ok(())
    }

    fn put_string(&mut self, s: &str) -> PyResult<()> {
        if self.is_url && !s.is_empty() {
            return write_url_or_text(self.worksheet, self.row, self.col, s, self.text_fmt);
        }
        match self.text_fmt {
            Some(fmt) => self
                .worksheet
                .write_string_with_format(self.row, self.col, s, fmt)
                .map_err(xlsx_err)?,
            None => self
                .worksheet
                .write_string(self.row, self.col, s)
                .map_err(xlsx_err)?,
        };
        Ok(())
    }
}

impl CellWriter for ExcelCell<'_> {
    fn write_none(&mut self) -> PyResult<()> {
        self.put_string("")
    }

    fn write_str(&mut self, s: &str) -> PyResult<()> {
        self.put_string(s)
    }

    fn write_bool(&mut self, b: bool) -> PyResult<()> {
        match self.text_fmt {
            Some(fmt) => self
                .worksheet
                .write_boolean_with_format(self.row, self.col, b, fmt)
                .map_err(xlsx_err)?,
            None => self
                .worksheet
                .write_boolean(self.row, self.col, b)
                .map_err(xlsx_err)?,
        };
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
        // On a band row `text_fmt` carries the fill, so an unformatted integer
        // column still gets shaded.
        write_number_opt(
            self.worksheet,
            self.row,
            self.col,
            val,
            self.col_override.or(self.text_fmt),
        )
    }

    fn write_datetime(&mut self, dt: &Bound<'_, PyDateTime>) -> PyResult<()> {
        let excel_dt = py_datetime_to_excel(dt)?;
        self.put_datetime(&excel_dt)
    }

    fn write_date(&mut self, d: &Bound<'_, PyDate>) -> PyResult<()> {
        let excel_dt = py_date_to_excel(d)?;
        self.put_datetime(&excel_dt)
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
    layout: &crate::helpers::SheetLayout,
    url_columns: Option<&Vec<String>>,
    formula_columns: Option<&Bound<'_, PyAny>>,
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
    // Number of data rows written, needed for the autofilter range.
    let mut data_rows: u32 = 0;

    // Merges, row heights and row formats must all land before the first data
    // cell — constant-memory mode cannot revisit a flushed row.
    layout.apply(worksheet)?;
    let formula_cols = crate::helpers::resolve_formula_columns(formula_columns)?;
    // Applies to every formula this sheet writes, so it must be set before any
    // of them. The crate's default cached result is 0, which readers that trust
    // the cache take at face value — `pandas.read_excel` reports 0, and
    // LibreOffice displays 0 instead of recalculating. An empty result reads
    // back as "not computed" and forces a recalculation on open.
    worksheet.set_formula_result_default("");
    let first_data_row = layout.first_data_row();
    let header_row = layout.header_row;
    let banding = layout.band_color.is_some();

    match records {
        WorksheetData::ArrowDataFrame(stream_obj) => {
            // Producing the stream can fail even though the object advertises
            // `__arrow_c_stream__` — pandas raises ImportError when pyarrow is
            // not installed, and an all-Null empty frame has no Arrow type. The
            // reader is therefore built up front: while nothing has been
            // written yet, falling back to the column-by-column writer is safe.
            // Once batches start landing a failure has to propagate, because
            // re-writing from the top would duplicate rows.
            let reader = crate::arrow_ffi::stream_to_reader(stream_obj, py);
            let arrow_ok = match reader {
                Err(stream_err) => Err(stream_err),
                Ok(reader) => (|| -> PyResult<()> {
                let schema = reader.schema();
                final_headers = schema
                    .fields()
                    .iter()
                    .map(|f| f.name().to_string())
                    .collect();
                for fc in &formula_cols {
                    final_headers.push(fc.header.clone());
                }
                let n_data_cols = final_headers.len() - formula_cols.len();
                write_all_headers(
                    worksheet,
                    header_row,
                    &final_headers,
                    bold_headers,
                    &bold_fmt,
                    index_columns,
                    header_format.map(|h| &h.inner),
                )?;

                let mut current_row: u32 = layout.first_data_row();
                let mut formats_set = false;

                // Resolve per-column formats ONCE (after headers are known).
                // Apply set_column_format BEFORE the first batch — constant memory mode requires
                // column formats to be set before their data rows.
                let col_formats: Vec<Option<crate::format::Format>> =
                    crate::format::resolve_column_formats(column_formats, &final_headers, py)?;
                let (plain, banded) = crate::format::build_palettes(
                    &col_formats,
                    float_fmt.as_ref(),
                    &datetime_fmt,
                    layout.band_color.as_deref(),
                )?;
                let url_cols =
                    crate::helpers::resolve_url_columns(url_columns, &final_headers, py)?;

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
                        &plain,
                        banded.as_ref(),
                        layout,
                        &url_cols,
                        &formula_cols,
                        n_data_cols,
                    )?;

                    current_row += batch.num_rows() as u32;
                }
                data_rows = current_row - layout.first_data_row();
                Ok(())
                })(),
            };

            if let Err(stream_err) = arrow_ok {
                let df = stream_obj.bind(py);
                // Polars exposes `get_column`; pandas indexes with `[]`.
                let polars = df.hasattr("get_column")?;
                if !df.hasattr("columns")? || !df.hasattr("dtypes")? {
                    return Err(stream_err);
                }
                if polars {
                    write_dataframe(
                        worksheet, py, stream_obj, &mut final_headers, &mut data_rows,
                        column_formats, float_fmt.as_ref(), &datetime_fmt,
                        &mut datetime_cols_set, bold_headers, &bold_fmt, index_columns,
                        header_format, layout, url_columns, &formula_cols,
                        "get_column", "to_list",
                        |dtype| Ok(polars_kind(&dtype.to_string())),
                    )?;
                } else {
                    write_dataframe(
                        worksheet, py, stream_obj, &mut final_headers, &mut data_rows,
                        column_formats, float_fmt.as_ref(), &datetime_fmt,
                        &mut datetime_cols_set, bold_headers, &bold_fmt, index_columns,
                        header_format, layout, url_columns, &formula_cols,
                        "__getitem__", "tolist",
                        |dtype| {
                            let kind: String = dtype.getattr("kind")?.extract()?;
                            Ok(map_pandas_kind(kind.chars().next().unwrap_or('O')))
                        },
                    )?;
                }
            }
        }

        WorksheetData::Records(records_list) => {
            // Propagate rather than skip: a non-iterable input used to
            // produce an empty sheet and report success.
            let rows = records_list.bind(py).try_iter().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "records must be an iterable of dicts, a DataFrame, or an Arrow stream",
                )
            })?;
            let mut headers: Vec<String> = Vec::new();
            let mut headers_written = false;
            let mut col_types: Vec<ColType> = Vec::new();
            // Resolved once when headers are first seen; kept alive for the
            // entire row loop so we can hand out &Format borrows per cell.
            // `banded` is None unless the sheet asked for alternating rows.
            let mut palettes: Option<(
                crate::format::RowPalette,
                Option<crate::format::RowPalette>,
            )> = None;
            let mut url_cols: Vec<bool> = Vec::new();
            let mut n_data_cols: usize = 0;

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
                    for fc in &formula_cols {
                        headers.push(fc.header.clone());
                    }
                    write_all_headers(
                        worksheet,
                        header_row,
                        &headers,
                        bold_headers,
                        &bold_fmt,
                        index_columns,
                        header_format.map(|h| &h.inner),
                    )?;
                    col_types.resize(headers.len(), ColType::Unknown);
                    final_headers = headers.clone();
                    n_data_cols = headers.len() - formula_cols.len();
                    // Resolve per-column formats ONCE and keep them for the whole loop.
                    // Apply set_column_format BEFORE writing data rows (constant memory mode).
                    let col_formats =
                        crate::format::resolve_column_formats(column_formats, &final_headers, py)?;
                    crate::format::apply_column_formats(worksheet, &col_formats)?;
                    palettes = Some(crate::format::build_palettes(
                        &col_formats,
                        float_fmt.as_ref(),
                        &datetime_fmt,
                        layout.band_color.as_deref(),
                    )?);
                    url_cols =
                        crate::helpers::resolve_url_columns(url_columns, &final_headers, py)?;
                    headers_written = true;
                }

                let row_u32 = layout.first_data_row() + row_idx as u32;
                let (plain, banded) = palettes
                    .as_ref()
                    .expect("palettes are built with the header row");
                let pal = if layout.is_banded(row_u32) {
                    banded.as_ref().unwrap_or(plain)
                } else {
                    plain
                };
                // Everything but `col`/`col_override` is fixed for the row,
                // so build the sink once and step it across the columns
                // rather than reassembling all nine fields per cell.
                let mut sink = ExcelCell {
                    worksheet: &mut *worksheet,
                    row: row_u32,
                    col: 0,
                    text_fmt: pal.text.as_ref(),
                    float_fmt: pal.float.as_ref(),
                    datetime_fmt: &pal.datetime,
                    datetime_cols_set: &mut datetime_cols_set,
                    col_override: None,
                    per_cell_datetime: banding,
                    is_url: false,
                };

                // Iterate the dict directly (insertion order == header order)
                // to avoid allocating a fresh `values()` list per row.
                for (col, (_key, value)) in row_dict.iter().enumerate() {
                    let cached = col_types
                        .get(col)
                        .copied()
                        .unwrap_or(ColType::Unknown);

                    sink.col = col as u16;
                    // Column format override: wins over float_fmt / datetime_fmt.
                    sink.col_override = pal.col(col);
                    sink.is_url = url_cols.get(col).copied().unwrap_or(false);

                    if !try_cached(&value, cached, &mut sink)? {
                        let detected = classify_and_write(&value, &mut sink)?;
                        if col < col_types.len() && col_types[col] == ColType::Unknown {
                            col_types[col] = detected;
                        }
                    }
                }
                if !formula_cols.is_empty() {
                    crate::helpers::write_formula_row(
                        worksheet,
                        &formula_cols,
                        n_data_cols as u16,
                        row_u32,
                        first_data_row,
                        None,
                    )?;
                }
                data_rows = row_idx as u32 + 1;
            }

        }

        WorksheetData::PandasDataFrame(df) => {
            write_dataframe(
                worksheet,
                py,
                df,
                &mut final_headers,
                &mut data_rows,
                column_formats,
                float_fmt.as_ref(),
                &datetime_fmt,
                &mut datetime_cols_set,
                bold_headers,
                &bold_fmt,
                index_columns,
                header_format,
                layout,
                url_columns,
                &formula_cols,
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
                &mut data_rows,
                column_formats,
                float_fmt.as_ref(),
                &datetime_fmt,
                &mut datetime_cols_set,
                bold_headers,
                &bold_fmt,
                index_columns,
                header_format,
                layout,
                url_columns,
                &formula_cols,
                "get_column",
                "to_list",
                |dtype| Ok(polars_kind(&dtype.to_string())),
            )?;
        }
    }

    layout.apply_autofilter(worksheet, data_rows, final_headers.len())?;
    layout.apply_totals(worksheet, &final_headers, data_rows, py)?;

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
    datetime_cols_set: &mut HashSet<u16>,
    plain: &crate::format::RowPalette,
    banded: Option<&crate::format::RowPalette>,
    layout: &crate::helpers::SheetLayout,
    url_cols: &[bool],
    formula_cols: &[crate::helpers::FormulaColumn],
    n_data_cols: usize,
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

    // Per-column format override is fixed for the whole column *within a
    // palette* — resolve both variants once instead of per cell.
    let plain_cols: Vec<Option<&Format>> =
        (0..bound_cols.len()).map(|c| plain.col(c)).collect();
    let banded_cols: Vec<Option<&Format>> = match banded {
        Some(b) => (0..bound_cols.len()).map(|c| b.col(c)).collect(),
        None => Vec::new(),
    };
    let banding = layout.band_color.is_some();

    for row in 0..nrows {
        let row_u32 = layout.first_data_row() + row as u32;
        let use_band = layout.is_banded(row_u32);
        let pal = if use_band { banded.unwrap_or(plain) } else { plain };
        let overrides = if use_band && banded.is_some() {
            &banded_cols
        } else {
            &plain_cols
        };
        let text_fmt = pal.text.as_ref();

        for (col_idx, col_list) in bound_cols.iter().enumerate() {
            let col_u16 = col_idx as u16;
            let item = col_list.get(row)?;
            let col_override = overrides[col_idx];

            if item.is_none() {
                write_string_opt(worksheet, row_u32, col_u16, "", text_fmt)?;
                continue;
            }

            match kind_at(col_idx) {
                ScalarKind::Int => {
                    let val: f64 = item.extract()?;
                    write_number_opt(
                        worksheet,
                        row_u32,
                        col_u16,
                        val,
                        col_override.or(text_fmt),
                    )?;
                }
                ScalarKind::Float => {
                    let val: f64 = item.extract()?;
                    write_num(
                        worksheet,
                        row_u32,
                        col_u16,
                        val,
                        col_override.or(pal.float.as_ref()),
                    )?;
                }
                ScalarKind::Bool => {
                    let val: bool = item.extract()?;
                    write_bool_opt(worksheet, row_u32, col_u16, val, text_fmt)?;
                }
                ScalarKind::Temporal => {
                    // With banding the shade has to ride on the cell, since a
                    // column format cannot alternate between rows.
                    let dt_fmt = banding.then(|| col_override.unwrap_or(&pal.datetime));
                    if let Ok(dt) = item.cast::<PyDateTime>() {
                        let excel_dt = py_datetime_to_excel(&dt)?;
                        write_datetime_opt(worksheet, row_u32, col_u16, &excel_dt, dt_fmt)?;
                    } else if let Ok(d) = item.cast::<PyDate>() {
                        let excel_dt = py_date_to_excel(&d)?;
                        write_datetime_opt(worksheet, row_u32, col_u16, &excel_dt, dt_fmt)?;
                    } else {
                        write_string_opt(
                            worksheet,
                            row_u32,
                            col_u16,
                            &item.to_string(),
                            text_fmt,
                        )?;
                    }
                }
                ScalarKind::Other => {
                    let mut sink = ExcelCell {
                        worksheet: &mut *worksheet,
                        row: row_u32,
                        col: col_u16,
                        text_fmt,
                        float_fmt: pal.float.as_ref(),
                        datetime_fmt: &pal.datetime,
                        datetime_cols_set: &mut *datetime_cols_set,
                        col_override,
                        per_cell_datetime: banding,
                        is_url: url_cols.get(col_idx).copied().unwrap_or(false),
                    };
                    classify_and_write(&item, &mut sink)?;
                }
            }
        }
        if !formula_cols.is_empty() {
            crate::helpers::write_formula_row(
                worksheet,
                formula_cols,
                n_data_cols as u16,
                row_u32,
                layout.first_data_row(),
                None,
            )?;
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
    data_rows: &mut u32,
    column_formats: Option<&Bound<'_, PyAny>>,
    float_fmt: Option<&Format>,
    datetime_fmt: &Format,
    datetime_cols_set: &mut HashSet<u16>,
    bold_headers: bool,
    bold_fmt: &Format,
    index_columns: Option<&Vec<String>>,
    header_format: Option<&crate::format::Format>,
    layout: &crate::helpers::SheetLayout,
    url_columns: Option<&Vec<String>>,
    formula_cols: &[crate::helpers::FormulaColumn],
    get_column_method: &str,
    to_list_method: &str,
    classify_dtype: C,
) -> PyResult<()>
where
    C: Fn(&Bound<'_, PyAny>) -> PyResult<ScalarKind>,
{
    let headers: Vec<String> = df.getattr(py, "columns")?.extract(py)?;
    let n_data_cols = headers.len();
    *final_headers = headers.clone();
    for fc in formula_cols {
        final_headers.push(fc.header.clone());
    }
    let dtypes = df.getattr(py, "dtypes")?;
    let dtypes_list: Vec<Bound<'_, PyAny>> =
        dtypes.bind(py).try_iter()?.collect::<Result<Vec<_>, _>>()?;

    write_all_headers(
        worksheet,
        layout.header_row,
        final_headers,
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
    *data_rows = nrows as u32;

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
    let (plain, banded) = crate::format::build_palettes(
        &col_formats,
        float_fmt,
        datetime_fmt,
        layout.band_color.as_deref(),
    )?;

    let url_cols = crate::helpers::resolve_url_columns(url_columns, final_headers, py)?;

    write_df_rows(
        worksheet,
        py,
        nrows,
        &col_lists,
        |col_idx| col_kinds[col_idx],
        datetime_cols_set,
        &plain,
        banded.as_ref(),
        layout,
        &url_cols,
        formula_cols,
        n_data_cols,
    )
}

/// Resolve a per-sheet value from a dict keyed by sheet name, falling back
/// to the `"general"` key. Returns the matching Python value, if any.
/// [`keyed_get`] plus extraction, treating an explicit `None` as absent.
/// Callers that want a Rust value rather than the raw object use this so the
/// `filter`/`map`/`transpose` dance is written once.
fn keyed_extract<'py, T>(
    dict: Option<&Bound<'py, pyo3::types::PyDict>>,
    sheet: &str,
) -> PyResult<Option<T>>
where
    T: for<'a> FromPyObject<'a, 'py, Error = PyErr>,
{
    keyed_get(dict, sheet)?
        .filter(|v| !v.is_none())
        .map(|v| v.extract())
        .transpose()
}

/// [`keyed_extract`] for `Format`, which has its own `FromPyObject` error type
/// and so does not fit that generic bound.
fn keyed_format(
    dict: Option<&Bound<'_, pyo3::types::PyDict>>,
    sheet: &str,
) -> PyResult<Option<crate::format::Format>> {
    match keyed_get(dict, sheet)? {
        Some(v) if !v.is_none() => Ok(Some(v.extract()?)),
        _ => Ok(None),
    }
}

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
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false, column_width = None, column_widths = None, column_formats = None, header_format = None, dedupe_strings = None, header_row = None, merge_ranges = None, row_heights = None, row_formats = None, banded_rows = None, autofilter = None, url_columns = None, totals_row = None, totals_label = None, totals_format = None, formula_columns = None))]
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
    dedupe_strings: Option<Bound<'_, pyo3::types::PyDict>>,
    header_row: Option<Bound<'_, pyo3::types::PyDict>>,
    merge_ranges: Option<Bound<'_, pyo3::types::PyDict>>,
    row_heights: Option<Bound<'_, pyo3::types::PyDict>>,
    row_formats: Option<Bound<'_, pyo3::types::PyDict>>,
    banded_rows: Option<Bound<'_, pyo3::types::PyDict>>,
    autofilter: Option<Bound<'_, pyo3::types::PyDict>>,
    url_columns: Option<Bound<'_, pyo3::types::PyDict>>,
    totals_row: Option<Bound<'_, pyo3::types::PyDict>>,
    totals_label: Option<Bound<'_, pyo3::types::PyDict>>,
    totals_format: Option<Bound<'_, pyo3::types::PyDict>>,
    formula_columns: Option<Bound<'_, pyo3::types::PyDict>>,
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    for (sheet_name, records) in records_with_sheet_name {
        ensure_valid_sheet_name(&sheet_name)?;

        let dedupe = keyed_extract::<bool>(dedupe_strings.as_ref(), &sheet_name)?
            .unwrap_or(false);

        let mut worksheet = if dedupe {
            workbook.add_worksheet()
        } else {
            workbook.add_worksheet_with_constant_memory()
        };
        worksheet.set_name(&sheet_name).map_err(xlsx_err)?;

        let pane = freeze_panes
            .as_ref()
            .map(|c| c.resolve(&sheet_name))
            .unwrap_or_default();

        let sheet_uniform = keyed_extract::<f64>(column_width.as_ref(), &sheet_name)?;
        let sheet_spec: Option<Bound<'_, PyAny>> =
            keyed_get(column_widths.as_ref(), &sheet_name)?;

        let sheet_col_fmts: Option<Bound<'_, PyAny>> =
            keyed_get(column_formats.as_ref(), &sheet_name)?;
        let sheet_hdr_fmt = keyed_format(header_format.as_ref(), &sheet_name)?;

        let sheet_header_row =
            keyed_extract::<u32>(header_row.as_ref(), &sheet_name)?.unwrap_or(0);
        let sheet_band = keyed_extract::<String>(banded_rows.as_ref(), &sheet_name)?;
        let layout = crate::helpers::resolve_layout(
            sheet_header_row,
            keyed_get(merge_ranges.as_ref(), &sheet_name)?.as_ref(),
            keyed_get(row_heights.as_ref(), &sheet_name)?.as_ref(),
            keyed_get(row_formats.as_ref(), &sheet_name)?.as_ref(),
            sheet_band,
            keyed_extract::<bool>(autofilter.as_ref(), &sheet_name)?.unwrap_or(false),
            keyed_get(totals_row.as_ref(), &sheet_name)?.as_ref(),
            keyed_extract::<String>(totals_label.as_ref(), &sheet_name)?,
            keyed_format(totals_format.as_ref(), &sheet_name)?.map(|f| f.inner),
        )?;

        let sheet_urls = keyed_extract::<Vec<String>>(url_columns.as_ref(), &sheet_name)?;

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
            &layout,
            sheet_urls.as_ref(),
            keyed_get(formula_columns.as_ref(), &sheet_name)?.as_ref(),
            py,
        )?;
    }

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None, freeze_row = None, freeze_col = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false, column_width = None, column_widths = None, column_formats = None, header_format = None, dedupe_strings = false, header_row = 0, merge_ranges = None, row_heights = None, row_formats = None, banded_rows = None, autofilter = false, url_columns = None, totals_row = None, totals_label = None, totals_format = None, formula_columns = None))]
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
    dedupe_strings: bool,
    header_row: u32,
    merge_ranges: Option<Bound<'_, PyAny>>,
    row_heights: Option<Bound<'_, PyAny>>,
    row_formats: Option<Bound<'_, PyAny>>,
    banded_rows: Option<String>,
    autofilter: bool,
    url_columns: Option<Vec<String>>,
    totals_row: Option<Bound<'_, PyAny>>,
    totals_label: Option<String>,
    totals_format: Option<Bound<'_, crate::format::Format>>,
    formula_columns: Option<Bound<'_, PyAny>>,
) -> PyResult<()> {
    let layout = crate::helpers::resolve_layout(
        header_row,
        merge_ranges.as_ref(),
        row_heights.as_ref(),
        row_formats.as_ref(),
        banded_rows,
        autofilter,
        totals_row.as_ref(),
        totals_label,
        totals_format.map(|f| f.borrow().inner.clone()),
    )?;
    let mut workbook = Workbook::new();
    let mut worksheet = if dedupe_strings {
        workbook.add_worksheet()
    } else {
        workbook.add_worksheet_with_constant_memory()
    };

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
        &layout,
        url_columns.as_ref(),
        formula_columns.as_ref(),
        py,
    )?;

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}
