use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateTime};
use pyo3::{Py, PyAny, Python};
use rust_xlsxwriter::{Format, Workbook};
use std::collections::HashSet;

use crate::data_types::{FreezePanesConfig, WorksheetData};
use crate::helpers::{
    py_date_to_excel, py_datetime_to_excel, save_workbook, write_header, write_num, ColType,
};
use crate::utils::ensure_valid_sheet_name;

pub fn xlsx_err(e: impl std::fmt::Display) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("Excel write error: {}", e))
}

/// Generic cell writer with full Python type cascade. Returns the
/// detected column type so callers may cache and fast-path subsequent rows.
fn write_py_any(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &Bound<PyAny>,
    float_fmt: Option<&Format>,
    datetime_fmt: Option<&Format>,
    datetime_cols_set: &mut HashSet<u16>,
) -> PyResult<ColType> {
    if value.is_none() {
        worksheet.write_string(row, col, "").map_err(xlsx_err)?;
        return Ok(ColType::Unknown);
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        worksheet
            .write_string(row, col, s.to_str()?)
            .map_err(xlsx_err)?;
        return Ok(ColType::String);
    }

    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        write_num(worksheet, row, col, f.value(), float_fmt)?;
        return Ok(ColType::Float);
    }

    // Bool BEFORE Int (Python bool is a subclass of int)
    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        worksheet
            .write_boolean(row, col, b.is_true())
            .map_err(xlsx_err)?;
        return Ok(ColType::Bool);
    }

    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        let val: f64 = i.extract()?;
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        return Ok(ColType::Int);
    }

    if let Ok(dt) = value.cast::<PyDateTime>() {
        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }
        let excel_dt = py_datetime_to_excel(&dt)?;
        worksheet
            .write_datetime(row, col, &excel_dt)
            .map_err(xlsx_err)?;
        return Ok(ColType::DateTime);
    }

    // Date AFTER DateTime (datetime is a subclass of date)
    if let Ok(d) = value.cast::<PyDate>() {
        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }
        let excel_dt = py_date_to_excel(&d)?;
        worksheet
            .write_datetime(row, col, &excel_dt)
            .map_err(xlsx_err)?;
        return Ok(ColType::Date);
    }

    // numpy scalar fallback: bool before f64 (numpy.bool_ extracts as f64 too)
    if let Ok(val) = value.extract::<bool>() {
        worksheet.write_boolean(row, col, val).map_err(xlsx_err)?;
        return Ok(ColType::Bool);
    }

    if let Ok(val) = value.extract::<f64>() {
        write_num(worksheet, row, col, val, float_fmt)?;
        return Ok(ColType::Float);
    }

    worksheet
        .write_string(row, col, value.to_string())
        .map_err(xlsx_err)?;
    Ok(ColType::String)
}

/// Fast-path dispatch using a cached `ColType`. Returns `true` if the
/// value matched the cached type and was written; `false` to fall back
/// to [`write_py_any`].
fn try_cached_write(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &Bound<PyAny>,
    cached: ColType,
    float_fmt: Option<&Format>,
    datetime_fmt: &Format,
    datetime_cols_set: &mut HashSet<u16>,
) -> PyResult<bool> {
    match cached {
        ColType::String => {
            if let Ok(s) = value.cast::<pyo3::types::PyString>() {
                worksheet
                    .write_string(row, col, s.to_str()?)
                    .map_err(xlsx_err)?;
                return Ok(true);
            }
        }
        ColType::Float => {
            if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
                write_num(worksheet, row, col, f.value(), float_fmt)?;
                return Ok(true);
            }
        }
        ColType::Bool => {
            if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
                worksheet
                    .write_boolean(row, col, b.is_true())
                    .map_err(xlsx_err)?;
                return Ok(true);
            }
        }
        ColType::Int => {
            if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
                let val: f64 = i.extract()?;
                worksheet.write_number(row, col, val).map_err(xlsx_err)?;
                return Ok(true);
            }
        }
        ColType::DateTime => {
            if let Ok(dt) = value.cast::<PyDateTime>() {
                if datetime_cols_set.insert(col) {
                    worksheet
                        .set_column_format(col, datetime_fmt)
                        .map_err(xlsx_err)?;
                }
                let excel_dt = py_datetime_to_excel(&dt)?;
                worksheet
                    .write_datetime(row, col, &excel_dt)
                    .map_err(xlsx_err)?;
                return Ok(true);
            }
        }
        ColType::Date => {
            if let Ok(d) = value.cast::<PyDate>() {
                if datetime_cols_set.insert(col) {
                    worksheet
                        .set_column_format(col, datetime_fmt)
                        .map_err(xlsx_err)?;
                }
                let excel_dt = py_date_to_excel(&d)?;
                worksheet
                    .write_datetime(row, col, &excel_dt)
                    .map_err(xlsx_err)?;
                return Ok(true);
            }
        }
        ColType::Unknown => {}
    }
    Ok(false)
}

/// Polars dtype classification by stringified dtype. Cheaper than
/// three `call_method0("is_*")` Python round-trips per column.
#[derive(Copy, Clone, PartialEq, Eq)]
enum PolarsKind {
    Int,
    Float,
    Bool,
    Temporal,
    Other,
}

fn polars_kind(dtype_str: &str) -> PolarsKind {
    if dtype_str.starts_with("Int") || dtype_str.starts_with("UInt") {
        PolarsKind::Int
    } else if dtype_str.starts_with("Float") || dtype_str == "Decimal" {
        PolarsKind::Float
    } else if dtype_str == "Boolean" {
        PolarsKind::Bool
    } else if dtype_str.starts_with("Date")
        || dtype_str.starts_with("Datetime")
        || dtype_str.starts_with("Time")
        || dtype_str.starts_with("Duration")
    {
        PolarsKind::Temporal
    } else {
        PolarsKind::Other
    }
}

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
    py: Python,
) -> PyResult<()> {
    let float_fmt = float_format.map(|s| Format::new().set_num_format(s));
    let dt_fmt_str = datetime_format
        .map(|s| s.as_str())
        .unwrap_or("yyyy-mm-ddThh:mm:ss");
    let datetime_fmt = Format::new().set_num_format(dt_fmt_str);
    let bold_fmt = Format::new().set_bold();
    let mut datetime_cols_set: HashSet<u16> = HashSet::new();

    match records {
        WorksheetData::ArrowDataFrame(stream_obj) => {
            let arrow_ok = (|| -> PyResult<()> {
                let reader = crate::arrow_ffi::stream_to_reader(stream_obj, py)?;

                let schema = reader.schema();
                for (col, field) in schema.fields().iter().enumerate() {
                    write_header(
                        worksheet,
                        col as u16,
                        field.name().as_str(),
                        bold_headers,
                        &bold_fmt,
                        index_columns,
                    )?;
                }

                let mut current_row: u32 = 1;
                let mut formats_set = false;

                for batch_result in reader {
                    let batch = batch_result.map_err(|e| {
                        PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                            "Failed to read Arrow batch: {}",
                            e
                        ))
                    })?;

                    if !formats_set {
                        crate::arrow_writer::set_datetime_column_formats(
                            worksheet,
                            &batch,
                            &datetime_fmt,
                        )?;
                        formats_set = true;
                    }

                    crate::arrow_writer::write_arrow_batch(
                        worksheet,
                        &batch,
                        current_row,
                        float_fmt.as_ref(),
                    )?;

                    current_row += batch.num_rows() as u32;
                }
                Ok(())
            })();

            // If Arrow FFI failed (e.g. Null-typed empty DataFrame), at
            // least write the header row from `.columns`.
            if arrow_ok.is_err() {
                if let Ok(cols) = stream_obj.getattr(py, "columns") {
                    let headers: Vec<String> = cols.extract(py).unwrap_or_default();
                    for (col, header) in headers.iter().enumerate() {
                        write_header(
                            worksheet,
                            col as u16,
                            header,
                            bold_headers,
                            &bold_fmt,
                            index_columns,
                        )?;
                    }
                }
            }
        }

        WorksheetData::Records(records_list) => {
            if let Ok(rows) = records_list.bind(py).try_iter() {
                let mut headers: Vec<String> = Vec::new();
                let mut headers_written = false;
                let mut col_types: Vec<ColType> = Vec::new();

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
                        for (col, header) in headers.iter().enumerate() {
                            write_header(
                                worksheet,
                                col as u16,
                                header,
                                bold_headers,
                                &bold_fmt,
                                index_columns,
                            )?;
                        }
                        col_types.resize(headers.len(), ColType::Unknown);
                        headers_written = true;
                    }

                    let row_u32 = (row_idx + 1) as u32;
                    for (col, value) in row_dict.values().iter().enumerate() {
                        let col_u16 = col as u16;
                        let cached = col_types
                            .get(col)
                            .copied()
                            .unwrap_or(ColType::Unknown);

                        let written = try_cached_write(
                            worksheet,
                            row_u32,
                            col_u16,
                            &value,
                            cached,
                            float_fmt.as_ref(),
                            &datetime_fmt,
                            &mut datetime_cols_set,
                        )?;

                        if !written {
                            let detected = write_py_any(
                                worksheet,
                                row_u32,
                                col_u16,
                                &value,
                                float_fmt.as_ref(),
                                Some(&datetime_fmt),
                                &mut datetime_cols_set,
                            )?;
                            if col < col_types.len() && col_types[col] == ColType::Unknown {
                                col_types[col] = detected;
                            }
                        }
                    }
                }
            }
        }

        WorksheetData::PandasDataFrame(df) => {
            let columns = df.getattr(py, "columns")?;
            let headers: Vec<String> = columns.extract(py)?;
            let dtypes = df.getattr(py, "dtypes")?;
            let dtypes_list: Vec<Bound<'_, PyAny>> =
                dtypes.bind(py).try_iter()?.collect::<Result<Vec<_>, _>>()?;

            for (col, header) in headers.iter().enumerate() {
                write_header(
                    worksheet,
                    col as u16,
                    header,
                    bold_headers,
                    &bold_fmt,
                    index_columns,
                )?;
            }

            let mut col_kinds: Vec<char> = Vec::with_capacity(headers.len());
            let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(headers.len());

            for (col_idx, header) in headers.iter().enumerate() {
                let kind: String = dtypes_list[col_idx].getattr("kind")?.extract()?;
                col_kinds.push(kind.chars().next().unwrap_or('O'));

                let col_series = df.call_method1(py, "__getitem__", (header.as_str(),))?;
                col_lists.push(col_series.call_method0(py, "tolist")?);
            }

            let nrows: usize = df.call_method0(py, "__len__")?.extract(py)?;

            for (col_idx, kind) in col_kinds.iter().enumerate() {
                if *kind == 'M' && datetime_cols_set.insert(col_idx as u16) {
                    worksheet
                        .set_column_format(col_idx as u16, &datetime_fmt)
                        .map_err(xlsx_err)?;
                }
            }

            write_df_rows(
                worksheet,
                py,
                nrows,
                &col_lists,
                |col_idx| map_pandas_kind(col_kinds[col_idx]),
                float_fmt.as_ref(),
                &datetime_fmt,
                &mut datetime_cols_set,
            )?;
        }

        WorksheetData::PolarsDataFrame(df) => {
            let columns: Vec<String> = df.getattr(py, "columns")?.extract(py)?;
            let dtypes = df.getattr(py, "dtypes")?;
            let dtypes_list: Vec<Bound<'_, PyAny>> =
                dtypes.bind(py).try_iter()?.collect::<Result<Vec<_>, _>>()?;

            for (col, header) in columns.iter().enumerate() {
                write_header(
                    worksheet,
                    col as u16,
                    header,
                    bold_headers,
                    &bold_fmt,
                    index_columns,
                )?;
            }

            let mut col_kinds: Vec<PolarsKind> = Vec::with_capacity(columns.len());
            let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(columns.len());

            for (col_idx, header) in columns.iter().enumerate() {
                col_kinds.push(polars_kind(&dtypes_list[col_idx].to_string()));
                let col_series = df.call_method1(py, "get_column", (header.as_str(),))?;
                col_lists.push(col_series.call_method0(py, "to_list")?);
            }

            let nrows: usize = df.call_method0(py, "__len__")?.extract(py)?;

            for (col_idx, kind) in col_kinds.iter().enumerate() {
                if *kind == PolarsKind::Temporal
                    && datetime_cols_set.insert(col_idx as u16)
                {
                    worksheet
                        .set_column_format(col_idx as u16, &datetime_fmt)
                        .map_err(xlsx_err)?;
                }
            }

            write_df_rows(
                worksheet,
                py,
                nrows,
                &col_lists,
                |col_idx| col_kinds[col_idx],
                float_fmt.as_ref(),
                &datetime_fmt,
                &mut datetime_cols_set,
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
    if let Some(password) = password {
        worksheet.protect_with_password(password);
    }

    Ok(())
}

fn map_pandas_kind(kind: char) -> PolarsKind {
    match kind {
        'i' | 'u' => PolarsKind::Int,
        'f' => PolarsKind::Float,
        'b' => PolarsKind::Bool,
        'M' => PolarsKind::Temporal,
        _ => PolarsKind::Other,
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
) -> PyResult<()>
where
    F: Fn(usize) -> PolarsKind,
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

    for row in 0..nrows {
        let row_u32 = (row + 1) as u32;
        for (col_idx, col_list) in bound_cols.iter().enumerate() {
            let col_u16 = col_idx as u16;
            let item = col_list.get(row)?;

            if item.is_none() {
                worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                continue;
            }

            match kind_at(col_idx) {
                PolarsKind::Int => {
                    let val: f64 = item.extract()?;
                    worksheet
                        .write_number(row_u32, col_u16, val)
                        .map_err(xlsx_err)?;
                }
                PolarsKind::Float => {
                    let val: f64 = item.extract()?;
                    write_num(worksheet, row_u32, col_u16, val, float_fmt)?;
                }
                PolarsKind::Bool => {
                    let val: bool = item.extract()?;
                    worksheet
                        .write_boolean(row_u32, col_u16, val)
                        .map_err(xlsx_err)?;
                }
                PolarsKind::Temporal => {
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
                PolarsKind::Other => {
                    write_py_any(
                        worksheet,
                        row_u32,
                        col_u16,
                        &item,
                        float_fmt,
                        Some(datetime_fmt),
                        datetime_cols_set,
                    )?;
                }
            }
        }
    }
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false))]
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
            py,
        )?;
    }

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}

#[pyfunction]
#[pyo3(signature = (records, file_name, sheet_name = None, password = None, freeze_row = None, freeze_col = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false))]
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
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    let mut worksheet = workbook.add_worksheet_with_constant_memory();

    if let Some(sheet_name) = sheet_name {
        ensure_valid_sheet_name(&sheet_name)?;
        worksheet.set_name(sheet_name).map_err(xlsx_err)?;
    }

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
        py,
    )?;

    save_workbook(py, &mut workbook, file_name)?;
    Ok(())
}
