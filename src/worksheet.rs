use pyo3::{Py, PyAny, Python};
use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateAccess, PyDateTime, PyTimeAccess};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};
use std::collections::HashSet;

use crate::data_types::{FreezePaneConfig, WorksheetData};
use crate::utils::validate_sheet_name;

/// Helper to convert rust_xlsxwriter errors to PyErr
pub fn xlsx_err(e: impl std::fmt::Display) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("Excel write error: {}", e))
}

fn save_workbook(py: Python, workbook: &mut Workbook, file_or_buffer: Py<PyAny>) -> PyResult<()> {
    if let Ok(file_name) = file_or_buffer.extract::<String>(py) {
        workbook.save(&file_name).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e))
        })?;
    } else {
        let buffer = workbook.save_to_buffer().map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to save workbook to buffer: {}",
                e
            ))
        })?;

        if let Ok(write_method) = file_or_buffer.getattr(py, "write") {
            let py_bytes = pyo3::types::PyBytes::new(py, &buffer);
            write_method.call1(py, (py_bytes,))?;
        } else {
            return Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                "Argument 'file_name' must be a string path or a file-like object with a 'write' method",
            ));
        }
    }
    Ok(())
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
    // Pre-create Format objects once instead of per-cell
    let float_fmt = float_format.map(|s| Format::new().set_num_format(s));
    let dt_fmt_str = datetime_format.map(|s| s.as_str()).unwrap_or("yyyy-mm-ddThh:mm:ss");
    let datetime_fmt = Format::new().set_num_format(dt_fmt_str);
    let bold_fmt = Format::new().set_bold();
    let mut datetime_cols_set: HashSet<u16> = HashSet::new();

    match records {
        WorksheetData::ArrowStream(stream_obj) => {
            // Zero-copy Arrow path via Arrow C Stream Interface (no pyo3-arrow dependency)
            let arrow_ok = (|| -> PyResult<()> {
                let reader = crate::arrow_ffi::stream_to_reader(stream_obj, py)?;

                // Write headers from schema (handles empty DataFrames with 0 batches)
                let schema = reader.schema();
                for (col, field) in schema.fields().iter().enumerate() {
                    let name = field.name().as_str();
                    if bold_headers {
                        worksheet.write_string_with_format(0, col as u16, name, &bold_fmt).map_err(xlsx_err)?;
                    } else {
                        worksheet.write_string(0, col as u16, name).map_err(xlsx_err)?;
                    }
                    if let Some(idx_cols) = index_columns {
                        if idx_cols.iter().any(|c| c == name) {
                            worksheet.set_column_format(col as u16, &bold_fmt).map_err(xlsx_err)?;
                        }
                    }
                }

                let mut current_row: u32 = 1;
                let mut formats_set = false;

                for batch_result in reader {
                    let batch = batch_result.map_err(|e| {
                        PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                            "Failed to read Arrow batch: {}", e
                        ))
                    })?;

                    if !formats_set {
                        crate::arrow_writer::set_datetime_column_formats(
                            worksheet, &batch, &datetime_fmt,
                        )?;
                        formats_set = true;
                    }

                    crate::arrow_writer::write_arrow_batch(
                        worksheet, &batch, current_row,
                        float_fmt.as_ref(),
                    )?;

                    current_row += batch.num_rows() as u32;
                }
                Ok(())
            })();

            // If Arrow path fails (e.g. empty Null-typed DataFrames), write headers only
            if arrow_ok.is_err() {
                if let Ok(cols) = stream_obj.getattr(py, "columns") {
                    let headers: Vec<String> = cols.extract(py).unwrap_or_default();
                    for (col, header) in headers.iter().enumerate() {
                        if bold_headers {
                            worksheet.write_string_with_format(0, col as u16, header, &bold_fmt).map_err(xlsx_err)?;
                        } else {
                            worksheet.write_string(0, col as u16, header).map_err(xlsx_err)?;
                        }
                    }
                }
            }
        }
        WorksheetData::Records(records_list) => {
            if let Ok(rows) = records_list.bind(py).try_iter() {
                let mut headers: Vec<String> = Vec::new();
                let mut headers_written = false;
                // Column type cache: detected from first row, used for fast dispatch
                // 0=unknown, 1=string, 2=float, 3=bool, 4=int, 5=datetime, 6=date
                let mut col_types: Vec<u8> = Vec::new();

                for (row_idx, row_res) in rows.enumerate() {
                    let row_obj = row_res?;
                    let row_dict = row_obj.cast::<pyo3::types::PyDict>()
                        .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>("Items in records must be dictionaries"))?;

                    if !headers_written {
                        let keys = row_dict.keys();
                        for key in keys.iter() {
                            let key_str = key.extract::<String>()?;
                            headers.push(key_str);
                        }

                        for (col, header) in headers.iter().enumerate() {
                            if bold_headers {
                                worksheet
                                    .write_string_with_format(0, col as u16, header, &bold_fmt)
                                    .map_err(xlsx_err)?;
                            } else {
                                worksheet
                                    .write_string(0, col as u16, header)
                                    .map_err(xlsx_err)?;
                            }

                            if let Some(index_cols) = index_columns {
                                if index_cols.contains(header) {
                                    worksheet
                                        .set_column_format(col as u16, &bold_fmt)
                                        .map_err(xlsx_err)?;
                                }
                            }
                        }
                        col_types.resize(headers.len(), 0);
                        headers_written = true;
                    }

                    // Iterate values() directly — avoids per-cell hash lookup
                    let row_u32 = (row_idx + 1) as u32;
                    for (col, value) in row_dict.values().iter().enumerate() {
                        let col_u16 = col as u16;

                        if value.is_none() {
                            worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                            continue;
                        }

                        // Try cached type first (fast path)
                        let cached = if col < col_types.len() { col_types[col] } else { 0 };
                        let written = match cached {
                            1 => {
                                if let Ok(s) = value.cast::<pyo3::types::PyString>() {
                                    worksheet.write_string(row_u32, col_u16, s.to_str()?).map_err(xlsx_err)?;
                                    true
                                } else { false }
                            }
                            2 => {
                                if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
                                    let val = f.value();
                                    if let Some(fmt) = float_fmt.as_ref() {
                                        worksheet.write_number_with_format(row_u32, col_u16, val, fmt).map_err(xlsx_err)?;
                                    } else {
                                        worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                                    }
                                    true
                                } else { false }
                            }
                            3 => {
                                if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
                                    worksheet.write_boolean(row_u32, col_u16, b.is_true()).map_err(xlsx_err)?;
                                    true
                                } else { false }
                            }
                            4 => {
                                if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
                                    let val: f64 = i.extract()?;
                                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                                    true
                                } else { false }
                            }
                            5 => {
                                if let Ok(dt) = value.cast::<PyDateTime>() {
                                    if datetime_cols_set.insert(col_u16) {
                                        worksheet.set_column_format(col_u16, &datetime_fmt).map_err(xlsx_err)?;
                                    }
                                    let excel_dt = ExcelDateTime::from_ymd(dt.get_year() as u16, dt.get_month() as u8, dt.get_day() as u8)
                                        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create datetime: {}", e)))?
                                        .and_hms(dt.get_hour() as u16, dt.get_minute() as u8, dt.get_second() as u8)
                                        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create timestamp: {}", e)))?;
                                    worksheet.write_datetime(row_u32, col_u16, &excel_dt).map_err(xlsx_err)?;
                                    true
                                } else { false }
                            }
                            6 => {
                                if let Ok(d) = value.cast::<PyDate>() {
                                    if datetime_cols_set.insert(col_u16) {
                                        worksheet.set_column_format(col_u16, &datetime_fmt).map_err(xlsx_err)?;
                                    }
                                    let excel_dt = ExcelDateTime::from_ymd(d.get_year() as u16, d.get_month() as u8, d.get_day() as u8)
                                        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create date: {}", e)))?;
                                    worksheet.write_datetime(row_u32, col_u16, &excel_dt).map_err(xlsx_err)?;
                                    true
                                } else { false }
                            }
                            _ => false,
                        };

                        if !written {
                            // Full type dispatch (first row or cache miss)
                            let detected = write_py_any_bound_detect(
                                worksheet, row_u32, col_u16, &value,
                                float_fmt.as_ref(), Some(&datetime_fmt), &mut datetime_cols_set,
                            )?;
                            // Cache detected type for this column
                            if col < col_types.len() && col_types[col] == 0 {
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

            // Write headers
            for (col, header) in headers.iter().enumerate() {
                if bold_headers {
                    worksheet
                        .write_string_with_format(0, col as u16, header, &bold_fmt)
                        .map_err(xlsx_err)?;
                } else {
                    worksheet
                        .write_string(0, col as u16, header)
                        .map_err(xlsx_err)?;
                }

                if let Some(index_cols) = index_columns {
                    if index_cols.contains(header) {
                        worksheet
                            .set_column_format(col as u16, &bold_fmt)
                            .map_err(xlsx_err)?;
                    }
                }
            }

            // Pre-compute dtype kinds and bulk-convert columns via tolist()
            let mut col_kinds: Vec<String> = Vec::with_capacity(headers.len());
            let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(headers.len());

            for header in &headers {
                let dtype = dtypes.call_method1(py, "__getitem__", (header.as_str(),))?;
                let kind: String = dtype.getattr(py, "kind")?.extract(py)?;
                col_kinds.push(kind);

                let col_series = df.call_method1(py, "__getitem__", (header.as_str(),))?;
                let col_list = col_series.call_method0(py, "tolist")?;
                col_lists.push(col_list);
            }

            let nrows: usize = df.call_method0(py, "__len__")?.extract(py)?;

            // Set datetime column formats upfront
            for (col_idx, kind) in col_kinds.iter().enumerate() {
                if kind == "M" && datetime_cols_set.insert(col_idx as u16) {
                    worksheet.set_column_format(col_idx as u16, &datetime_fmt).map_err(xlsx_err)?;
                }
            }

            // Write data row-by-row (constant_memory compatible) with dtype-aware dispatch
            for row in 0..nrows {
                let row_u32 = (row + 1) as u32;
                for (col_idx, (col_list, kind)) in col_lists.iter().zip(col_kinds.iter()).enumerate() {
                    let col_u16 = col_idx as u16;
                    let item = col_list.bind(py).get_item(row)?;

                    if item.is_none() {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                        continue;
                    }

                    match kind.as_str() {
                        "i" | "u" => {
                            let val: f64 = item.extract()?;
                            worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                        }
                        "f" => {
                            let val: f64 = item.extract()?;
                            if val.is_nan() || val.is_infinite() {
                                worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                            } else if let Some(fmt) = float_fmt.as_ref() {
                                worksheet.write_number_with_format(row_u32, col_u16, val, fmt).map_err(xlsx_err)?;
                            } else {
                                worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                            }
                        }
                        "b" => {
                            let val: bool = item.extract()?;
                            worksheet.write_boolean(row_u32, col_u16, val).map_err(xlsx_err)?;
                        }
                        "M" => {
                            if let Ok(dt) = item.cast::<PyDateTime>() {
                                let excel_dt = ExcelDateTime::from_ymd(
                                    dt.get_year() as u16,
                                    dt.get_month() as u8,
                                    dt.get_day() as u8,
                                )
                                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                                    format!("Failed to create datetime: {}", e),
                                ))?
                                .and_hms(
                                    dt.get_hour() as u16,
                                    dt.get_minute() as u8,
                                    dt.get_second() as u8,
                                )
                                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                                    format!("Failed to create timestamp: {}", e),
                                ))?;
                                worksheet.write_datetime(row_u32, col_u16, &excel_dt).map_err(xlsx_err)?;
                            } else {
                                worksheet.write_string(row_u32, col_u16, item.to_string()).map_err(xlsx_err)?;
                            }
                        }
                        "U" | "S" => {
                            worksheet.write_string(row_u32, col_u16, item.to_string()).map_err(xlsx_err)?;
                        }
                        _ => {
                            write_py_any_bound(
                                worksheet,
                                row_u32,
                                col_u16,
                                &item,
                                float_fmt.as_ref(),
                                Some(&datetime_fmt),
                                &mut datetime_cols_set,
                            )?;
                        }
                    }
                }
            }
        }
        WorksheetData::PolarsDataFrame(df) => {
            let columns: Vec<String> = df.getattr(py, "columns")?.extract(py)?;
            let dtypes = df.getattr(py, "dtypes")?;

            // Write headers
            for (col, header) in columns.iter().enumerate() {
                if bold_headers {
                    worksheet
                        .write_string_with_format(0, col as u16, header, &bold_fmt)
                        .map_err(xlsx_err)?;
                } else {
                    worksheet
                        .write_string(0, col as u16, header)
                        .map_err(xlsx_err)?;
                }

                if let Some(index_cols) = index_columns {
                    if index_cols.contains(header) {
                        worksheet
                            .set_column_format(col as u16, &bold_fmt)
                            .map_err(xlsx_err)?;
                    }
                }
            }

            // Pre-compute dtype flags and column data via get_column().to_list()
            let dtypes_bound = dtypes.bind(py);
            let dtypes_list: Vec<Bound<'_, PyAny>> = dtypes_bound.try_iter()?
                .collect::<Result<Vec<_>, _>>()?;

            let mut col_lists: Vec<Py<PyAny>> = Vec::with_capacity(columns.len());
            let mut col_flags: Vec<&str> = Vec::with_capacity(columns.len());

            for (col_idx, header) in columns.iter().enumerate() {
                let dtype = &dtypes_list[col_idx];
                let dtype_str = dtype.to_string();

                // Detect type via is_*() methods and string matching
                let flag = if dtype.call_method0("is_integer")?.extract::<bool>()? {
                    "int"
                } else if dtype.call_method0("is_float")?.extract::<bool>()? {
                    "float"
                } else if dtype.call_method0("is_temporal")?.extract::<bool>()? {
                    "temporal"
                } else if dtype_str == "Boolean" {
                    "bool"
                } else {
                    "other"
                };
                col_flags.push(flag);

                let col_series = df.call_method1(py, "get_column", (header.as_str(),))?;
                let col_list = col_series.call_method0(py, "to_list")?;
                col_lists.push(col_list);
            }

            let nrows: usize = df.call_method0(py, "__len__")?.extract(py)?;

            // Set datetime column formats upfront
            for (col_idx, flag) in col_flags.iter().enumerate() {
                if *flag == "temporal" && datetime_cols_set.insert(col_idx as u16) {
                    worksheet.set_column_format(col_idx as u16, &datetime_fmt).map_err(xlsx_err)?;
                }
            }

            // Write data row-by-row (constant_memory compatible)
            for row in 0..nrows {
                let row_u32 = (row + 1) as u32;
                for (col_idx, (col_list, flag)) in col_lists.iter().zip(col_flags.iter()).enumerate() {
                    let col_u16 = col_idx as u16;
                    let item = col_list.bind(py).get_item(row)?;

                    if item.is_none() {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                        continue;
                    }

                    match *flag {
                        "int" => {
                            let val: f64 = item.extract()?;
                            worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                        }
                        "float" => {
                            let val: f64 = item.extract()?;
                            if val.is_nan() || val.is_infinite() {
                                worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                            } else if let Some(fmt) = float_fmt.as_ref() {
                                worksheet.write_number_with_format(row_u32, col_u16, val, fmt).map_err(xlsx_err)?;
                            } else {
                                worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                            }
                        }
                        "bool" => {
                            let val: bool = item.extract()?;
                            worksheet.write_boolean(row_u32, col_u16, val).map_err(xlsx_err)?;
                        }
                        "temporal" => {
                            if let Ok(dt) = item.cast::<PyDateTime>() {
                                let excel_dt = ExcelDateTime::from_ymd(
                                    dt.get_year() as u16,
                                    dt.get_month() as u8,
                                    dt.get_day() as u8,
                                )
                                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                                    format!("Failed to create datetime: {}", e),
                                ))?
                                .and_hms(
                                    dt.get_hour() as u16,
                                    dt.get_minute() as u8,
                                    dt.get_second() as u8,
                                )
                                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                                    format!("Failed to create timestamp: {}", e),
                                ))?;
                                worksheet.write_datetime(row_u32, col_u16, &excel_dt).map_err(xlsx_err)?;
                            } else if let Ok(d) = item.cast::<PyDate>() {
                                let excel_dt = ExcelDateTime::from_ymd(
                                    d.get_year() as u16,
                                    d.get_month() as u8,
                                    d.get_day() as u8,
                                )
                                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                                    format!("Failed to create date: {}", e),
                                ))?;
                                worksheet.write_datetime(row_u32, col_u16, &excel_dt).map_err(xlsx_err)?;
                            } else {
                                worksheet.write_string(row_u32, col_u16, item.to_string()).map_err(xlsx_err)?;
                            }
                        }
                        _ => {
                            // String or other — per-cell type dispatch
                            write_py_any_bound(
                                worksheet, row_u32, col_u16, &item,
                                float_fmt.as_ref(), Some(&datetime_fmt), &mut datetime_cols_set,
                            )?;
                        }
                    }
                }
            }
        }
    }

    // Set freeze panes if specified
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

fn write_py_any_bound(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &Bound<PyAny>,
    float_fmt: Option<&Format>,
    datetime_fmt: Option<&Format>,
    datetime_cols_set: &mut HashSet<u16>,
) -> PyResult<()> {
    if value.is_none() {
        worksheet.write_string(row, col, "").map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        worksheet.write_string(row, col, s.to_str()?).map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        let val = f.value();
        if let Some(fmt) = float_fmt {
            worksheet.write_number_with_format(row, col, val, fmt).map_err(xlsx_err)?;
        } else {
            worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        }
        return Ok(());
    }

    // Check Bool BEFORE Int (Python bool is subclass of int)
    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        worksheet.write_boolean(row, col, b.is_true()).map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        let val: f64 = i.extract()?;
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        return Ok(());
    }

    if let Ok(datetime) = value.cast::<PyDateTime>() {
        let year = datetime.get_year() as u16;
        let month = datetime.get_month() as u8;
        let day = datetime.get_day() as u8;
        let hour = datetime.get_hour() as u16;
        let minute = datetime.get_minute() as u8;
        let second = datetime.get_second() as u8;

        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }

        let excel_datetime = ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Failed to create datetime: {}",
                    e
                ))
            })?
            .and_hms(hour, minute, second)
            .map_err(|e| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Failed to create timestamp: {}",
                    e
                ))
            })?;

        worksheet.write_datetime(row, col, &excel_datetime).map_err(xlsx_err)?;
        return Ok(());
    }

    // Check Date (must be after DateTime since datetime is subclass of date)
    if let Ok(d) = value.cast::<PyDate>() {
        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }

        let excel_date = ExcelDateTime::from_ymd(
            d.get_year() as u16,
            d.get_month() as u8,
            d.get_day() as u8,
        )
        .map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "Failed to create date: {}",
                e
            ))
        })?;

        worksheet.write_datetime(row, col, &excel_date).map_err(xlsx_err)?;
        return Ok(());
    }

    // Fallback: try bool extraction for numpy.bool_ (must be before f64, since bool extracts as f64 too)
    if let Ok(val) = value.extract::<bool>() {
        worksheet.write_boolean(row, col, val).map_err(xlsx_err)?;
        return Ok(());
    }

    // Fallback: try numeric extraction for numpy scalar types (numpy.int64, numpy.float64, etc.)
    if let Ok(val) = value.extract::<f64>() {
        if let Some(fmt) = float_fmt {
            worksheet.write_number_with_format(row, col, val, fmt).map_err(xlsx_err)?;
        } else {
            worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        }
        return Ok(());
    }

    // Final fallback to string representation
    worksheet.write_string(row, col, value.to_string()).map_err(xlsx_err)?;
    Ok(())
}

/// Same as write_py_any_bound but returns detected type ID for caching.
/// 0=unknown, 1=string, 2=float, 3=bool, 4=int, 5=datetime, 6=date
fn write_py_any_bound_detect(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &Bound<PyAny>,
    float_fmt: Option<&Format>,
    datetime_fmt: Option<&Format>,
    datetime_cols_set: &mut HashSet<u16>,
) -> PyResult<u8> {
    if value.is_none() {
        worksheet.write_string(row, col, "").map_err(xlsx_err)?;
        return Ok(0);
    }

    if let Ok(s) = value.cast::<pyo3::types::PyString>() {
        worksheet.write_string(row, col, s.to_str()?).map_err(xlsx_err)?;
        return Ok(1);
    }

    if let Ok(f) = value.cast::<pyo3::types::PyFloat>() {
        let val = f.value();
        if let Some(fmt) = float_fmt {
            worksheet.write_number_with_format(row, col, val, fmt).map_err(xlsx_err)?;
        } else {
            worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        }
        return Ok(2);
    }

    if let Ok(b) = value.cast::<pyo3::types::PyBool>() {
        worksheet.write_boolean(row, col, b.is_true()).map_err(xlsx_err)?;
        return Ok(3);
    }

    if let Ok(i) = value.cast::<pyo3::types::PyInt>() {
        let val: f64 = i.extract()?;
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        return Ok(4);
    }

    if let Ok(dt) = value.cast::<PyDateTime>() {
        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }
        let excel_dt = ExcelDateTime::from_ymd(dt.get_year() as u16, dt.get_month() as u8, dt.get_day() as u8)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create datetime: {}", e)))?
            .and_hms(dt.get_hour() as u16, dt.get_minute() as u8, dt.get_second() as u8)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create timestamp: {}", e)))?;
        worksheet.write_datetime(row, col, &excel_dt).map_err(xlsx_err)?;
        return Ok(5);
    }

    if let Ok(d) = value.cast::<PyDate>() {
        if let Some(fmt) = datetime_fmt {
            if datetime_cols_set.insert(col) {
                worksheet.set_column_format(col, fmt).map_err(xlsx_err)?;
            }
        }
        let excel_dt = ExcelDateTime::from_ymd(d.get_year() as u16, d.get_month() as u8, d.get_day() as u8)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to create date: {}", e)))?;
        worksheet.write_datetime(row, col, &excel_dt).map_err(xlsx_err)?;
        return Ok(6);
    }

    if let Ok(val) = value.extract::<bool>() {
        worksheet.write_boolean(row, col, val).map_err(xlsx_err)?;
        return Ok(3);
    }

    if let Ok(val) = value.extract::<f64>() {
        if let Some(fmt) = float_fmt {
            worksheet.write_number_with_format(row, col, val, fmt).map_err(xlsx_err)?;
        } else {
            worksheet.write_number(row, col, val).map_err(xlsx_err)?;
        }
        return Ok(2);
    }

    worksheet.write_string(row, col, value.to_string()).map_err(xlsx_err)?;
    Ok(1)
}


#[pyfunction]
#[pyo3(signature = (records_with_sheet_name, file_name, password = None, freeze_panes = None, float_format = None, datetime_format = None, index_columns = None, autofit = true, bold_headers = false))]
pub fn write_worksheets(
    py: Python,
    records_with_sheet_name: Vec<indexmap::IndexMap<String, WorksheetData>>,
    file_name: Py<PyAny>,
    password: Option<String>,
    freeze_panes: Option<FreezePaneConfig>,
    float_format: Option<String>,
    datetime_format: Option<String>,
    index_columns: Option<Vec<String>>,
    autofit: bool,
    bold_headers: bool,
) -> PyResult<()> {
    let mut workbook = Workbook::new();
    for record_map in records_with_sheet_name {
        for (sheet_name, records) in record_map {
            if !validate_sheet_name(&sheet_name) {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\",
                    sheet_name
                )));
            }

            let mut worksheet = workbook.add_worksheet_with_constant_memory();
            worksheet.set_name(&sheet_name).map_err(xlsx_err)?;

            let (mut freeze_row, mut freeze_col) = (None, None);
            if let Some(ref freeze_panes_config) = freeze_panes {
                if let Some(general_config) = freeze_panes_config.config.get("general") {
                    if let Some(&row) = general_config.get("row") {
                        freeze_row = Some(row as u32);
                    }
                    if let Some(&col) = general_config.get("col") {
                        freeze_col = Some(col as u16);
                    }
                }

                if let Some(sheet_config) = freeze_panes_config.config.get(&sheet_name) {
                    if let Some(&row) = sheet_config.get("row") {
                        freeze_row = Some(row as u32);
                    }
                    if let Some(&col) = sheet_config.get("col") {
                        freeze_col = Some(col as u16);
                    }
                }
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
        }
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
        if !validate_sheet_name(&sheet_name) {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "Invalid sheet name '{}'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\",
                sheet_name
            )));
        }
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
