use arrow_array::cast::AsArray;
use arrow_array::types::*;
use arrow_array::{Array, ArrayRef, BooleanArray, RecordBatch};
use arrow_schema::{DataType, TimeUnit};
use pyo3::prelude::*;
use rust_xlsxwriter::{ExcelDateTime, Format, Worksheet};

use crate::worksheet::xlsx_err;

/// Column type classification done once (outside the row loop) to avoid
/// paying `match data_type()` on every cell.
#[derive(Copy, Clone)]
enum ColKind {
    Int8,
    Int16,
    Int32,
    Int64,
    UInt8,
    UInt16,
    UInt32,
    UInt64,
    Float16,
    Float32,
    Float64,
    Bool,
    Utf8,
    LargeUtf8,
    Utf8View,
    Date32,
    Date64,
    Timestamp(TimeUnit),
    Unsupported,
}

fn classify(dt: &DataType) -> ColKind {
    match dt {
        DataType::Int8 => ColKind::Int8,
        DataType::Int16 => ColKind::Int16,
        DataType::Int32 => ColKind::Int32,
        DataType::Int64 => ColKind::Int64,
        DataType::UInt8 => ColKind::UInt8,
        DataType::UInt16 => ColKind::UInt16,
        DataType::UInt32 => ColKind::UInt32,
        DataType::UInt64 => ColKind::UInt64,
        DataType::Float16 => ColKind::Float16,
        DataType::Float32 => ColKind::Float32,
        DataType::Float64 => ColKind::Float64,
        DataType::Boolean => ColKind::Bool,
        DataType::Utf8 => ColKind::Utf8,
        DataType::LargeUtf8 => ColKind::LargeUtf8,
        DataType::Utf8View => ColKind::Utf8View,
        DataType::Date32 => ColKind::Date32,
        DataType::Date64 => ColKind::Date64,
        DataType::Timestamp(unit, _) => ColKind::Timestamp(*unit),
        _ => ColKind::Unsupported,
    }
}

fn write_float(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: f64,
    float_fmt: Option<&Format>,
) -> PyResult<()> {
    if val.is_nan() || val.is_infinite() {
        worksheet.write_string(row, col, "").map_err(xlsx_err)?;
    } else if let Some(fmt) = float_fmt {
        worksheet
            .write_number_with_format(row, col, val, fmt)
            .map_err(xlsx_err)?;
    } else {
        worksheet.write_number(row, col, val).map_err(xlsx_err)?;
    }
    Ok(())
}

/// Write an Arrow `RecordBatch` starting at `start_row`.
/// Column types are classified once; each cell dispatches on the cached kind.
pub fn write_arrow_batch(
    worksheet: &mut Worksheet,
    batch: &RecordBatch,
    start_row: u32,
    float_fmt: Option<&Format>,
) -> PyResult<()> {
    let num_cols = batch.num_columns();
    let num_rows = batch.num_rows();

    let columns: Vec<ArrayRef> = (0..num_cols).map(|c| batch.column(c).clone()).collect();
    let kinds: Vec<ColKind> = columns.iter().map(|c| classify(c.data_type())).collect();

    for row in 0..num_rows {
        let row_u32 = start_row + row as u32;
        for col_idx in 0..num_cols {
            let col_u16 = col_idx as u16;
            let column = &columns[col_idx];

            if column.is_null(row) {
                worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                continue;
            }

            match kinds[col_idx] {
                ColKind::Int8 => {
                    let val = column.as_primitive::<Int8Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Int16 => {
                    let val = column.as_primitive::<Int16Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Int32 => {
                    let val = column.as_primitive::<Int32Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Int64 => {
                    let val = column.as_primitive::<Int64Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::UInt8 => {
                    let val = column.as_primitive::<UInt8Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::UInt16 => {
                    let val = column.as_primitive::<UInt16Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::UInt32 => {
                    let val = column.as_primitive::<UInt32Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::UInt64 => {
                    let val = column.as_primitive::<UInt64Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Float16 => {
                    let val = column.as_primitive::<Float16Type>().value(row).to_f64();
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Float32 => {
                    let val = column.as_primitive::<Float32Type>().value(row) as f64;
                    write_float(worksheet, row_u32, col_u16, val, float_fmt)?;
                }
                ColKind::Float64 => {
                    let val = column.as_primitive::<Float64Type>().value(row);
                    write_float(worksheet, row_u32, col_u16, val, float_fmt)?;
                }
                ColKind::Bool => {
                    let arr = column
                        .as_any()
                        .downcast_ref::<BooleanArray>()
                        .expect("Boolean dtype guarantees downcast");
                    worksheet
                        .write_boolean(row_u32, col_u16, arr.value(row))
                        .map_err(xlsx_err)?;
                }
                ColKind::Utf8 => {
                    let val = column.as_string::<i32>().value(row);
                    worksheet.write_string(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::LargeUtf8 => {
                    let val = column.as_string::<i64>().value(row);
                    worksheet.write_string(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Utf8View => {
                    let val = column.as_string_view().value(row);
                    worksheet.write_string(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                ColKind::Date32 => {
                    let days = column.as_primitive::<Date32Type>().value(row);
                    match days_to_excel_date(days as i64) {
                        Some(dt) => worksheet
                            .write_datetime(row_u32, col_u16, &dt)
                            .map_err(xlsx_err)?,
                        None => worksheet
                            .write_string(row_u32, col_u16, "")
                            .map_err(xlsx_err)?,
                    };
                }
                ColKind::Date64 => {
                    let ms = column.as_primitive::<Date64Type>().value(row);
                    match millis_to_excel_datetime(ms) {
                        Some(dt) => worksheet
                            .write_datetime(row_u32, col_u16, &dt)
                            .map_err(xlsx_err)?,
                        None => worksheet
                            .write_string(row_u32, col_u16, "")
                            .map_err(xlsx_err)?,
                    };
                }
                ColKind::Timestamp(unit) => {
                    let micros = match unit {
                        TimeUnit::Second => {
                            column.as_primitive::<TimestampSecondType>().value(row) * 1_000_000
                        }
                        TimeUnit::Millisecond => {
                            column.as_primitive::<TimestampMillisecondType>().value(row) * 1_000
                        }
                        TimeUnit::Microsecond => {
                            column.as_primitive::<TimestampMicrosecondType>().value(row)
                        }
                        TimeUnit::Nanosecond => {
                            column.as_primitive::<TimestampNanosecondType>().value(row) / 1_000
                        }
                    };
                    match micros_to_excel_datetime(micros) {
                        Some(dt) => worksheet
                            .write_datetime(row_u32, col_u16, &dt)
                            .map_err(xlsx_err)?,
                        None => worksheet
                            .write_string(row_u32, col_u16, "")
                            .map_err(xlsx_err)?,
                    };
                }
                ColKind::Unsupported => {
                    worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                }
            }
        }
    }

    Ok(())
}

pub fn set_datetime_column_formats(
    worksheet: &mut Worksheet,
    batch: &RecordBatch,
    datetime_fmt: &Format,
) -> PyResult<()> {
    for (col_idx, field) in batch.schema().fields().iter().enumerate() {
        if matches!(
            field.data_type(),
            DataType::Date32 | DataType::Date64 | DataType::Timestamp(_, _)
        ) {
            worksheet
                .set_column_format(col_idx as u16, datetime_fmt)
                .map_err(xlsx_err)?;
        }
    }
    Ok(())
}

fn days_to_excel_date(days_since_epoch: i64) -> Option<ExcelDateTime> {
    let (y, m, d) = chrono_from_days(days_since_epoch)?;
    ExcelDateTime::from_ymd(y, m, d).ok()
}

fn millis_to_excel_datetime(ms: i64) -> Option<ExcelDateTime> {
    micros_to_excel_datetime(ms * 1000)
}

fn micros_to_excel_datetime(micros: i64) -> Option<ExcelDateTime> {
    let total_secs = micros / 1_000_000;
    let days = total_secs / 86400;
    let day_secs = (total_secs % 86400).unsigned_abs();

    let (year, month, day) = chrono_from_days(days)?;
    let hour = (day_secs / 3600) as u16;
    let minute = ((day_secs % 3600) / 60) as u8;
    let second = (day_secs % 60) as u8;

    ExcelDateTime::from_ymd(year, month, day)
        .ok()?
        .and_hms(hour, minute, second)
        .ok()
}

/// Howard Hinnant's civil_from_days — days since Unix epoch → (y, m, d).
fn chrono_from_days(days: i64) -> Option<(u16, u8, u8)> {
    let z = days + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = (z - era * 146097) as u32;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };

    if y < 0 || y > 9999 {
        return None;
    }
    Some((y as u16, m as u8, d as u8))
}
