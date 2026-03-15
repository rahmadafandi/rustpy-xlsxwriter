use arrow_array::cast::AsArray;
use arrow_array::types::*;
use arrow_array::{Array, BooleanArray, RecordBatch};
use arrow_schema::DataType;
use pyo3::prelude::*;
use rust_xlsxwriter::{ExcelDateTime, Format, Worksheet};

use crate::worksheet::xlsx_err;

/// Write an Arrow RecordBatch data rows to a worksheet (zero-copy, no Python object conversion).
/// Headers and formatting are handled by the caller.
pub fn write_arrow_batch(
    worksheet: &mut Worksheet,
    batch: &RecordBatch,
    start_row: u32,
    float_fmt: Option<&Format>,
) -> PyResult<()> {
    let num_cols = batch.num_columns();
    let num_rows = batch.num_rows();

    // Write data row-by-row for constant_memory compatibility
    // (headers are written by the caller before calling this function)
    for row in 0..num_rows {
        let row_u32 = start_row + row as u32;

        for col_idx in 0..num_cols {
            let col_u16 = col_idx as u16;
            let column = batch.column(col_idx);

            if column.is_null(row) {
                worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                continue;
            }

            match column.data_type() {
                DataType::Int8 => {
                    let val = column.as_primitive::<Int8Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Int16 => {
                    let val = column.as_primitive::<Int16Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Int32 => {
                    let val = column.as_primitive::<Int32Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Int64 => {
                    let val = column.as_primitive::<Int64Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::UInt8 => {
                    let val = column.as_primitive::<UInt8Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::UInt16 => {
                    let val = column.as_primitive::<UInt16Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::UInt32 => {
                    let val = column.as_primitive::<UInt32Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::UInt64 => {
                    let val = column.as_primitive::<UInt64Type>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Float16 => {
                    let val = column.as_primitive::<Float16Type>().value(row).to_f64();
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Float32 => {
                    let val = column.as_primitive::<Float32Type>().value(row) as f64;
                    if val.is_nan() || val.is_infinite() {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                    } else if let Some(fmt) = float_fmt {
                        worksheet
                            .write_number_with_format(row_u32, col_u16, val, fmt)
                            .map_err(xlsx_err)?;
                    } else {
                        worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                    }
                }
                DataType::Float64 => {
                    let val = column.as_primitive::<Float64Type>().value(row);
                    if val.is_nan() || val.is_infinite() {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                    } else if let Some(fmt) = float_fmt {
                        worksheet
                            .write_number_with_format(row_u32, col_u16, val, fmt)
                            .map_err(xlsx_err)?;
                    } else {
                        worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                    }
                }
                DataType::Boolean => {
                    let arr = column.as_any().downcast_ref::<BooleanArray>().unwrap();
                    worksheet
                        .write_boolean(row_u32, col_u16, arr.value(row))
                        .map_err(xlsx_err)?;
                }
                DataType::Utf8 => {
                    let val = column.as_string::<i32>().value(row);
                    worksheet.write_string(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::LargeUtf8 => {
                    let val = column.as_string::<i64>().value(row);
                    worksheet.write_string(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Utf8View => {
                    let val = column.as_string_view().value(row);
                    worksheet.write_string(row_u32, col_u16, val).map_err(xlsx_err)?;
                }
                DataType::Date32 => {
                    let days = column.as_primitive::<Date32Type>().value(row);
                    if let Some(dt) = days_to_excel_date(days as i64) {
                        worksheet.write_datetime(row_u32, col_u16, &dt).map_err(xlsx_err)?;
                    } else {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                    }
                }
                DataType::Date64 => {
                    let ms = column.as_primitive::<Date64Type>().value(row);
                    if let Some(dt) = millis_to_excel_datetime(ms) {
                        worksheet.write_datetime(row_u32, col_u16, &dt).map_err(xlsx_err)?;
                    } else {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                    }
                }
                DataType::Timestamp(unit, _) => {
                    let micros = match unit {
                        arrow_schema::TimeUnit::Second => {
                            column.as_primitive::<TimestampSecondType>().value(row) * 1_000_000
                        }
                        arrow_schema::TimeUnit::Millisecond => {
                            column.as_primitive::<TimestampMillisecondType>().value(row) * 1_000
                        }
                        arrow_schema::TimeUnit::Microsecond => {
                            column.as_primitive::<TimestampMicrosecondType>().value(row)
                        }
                        arrow_schema::TimeUnit::Nanosecond => {
                            column.as_primitive::<TimestampNanosecondType>().value(row) / 1_000
                        }
                    };
                    if let Some(dt) = micros_to_excel_datetime(micros) {
                        worksheet.write_datetime(row_u32, col_u16, &dt).map_err(xlsx_err)?;
                    } else {
                        worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                    }
                }
                _ => {
                    // Fallback: write empty for unsupported types
                    worksheet.write_string(row_u32, col_u16, "").map_err(xlsx_err)?;
                }
            }
        }
    }

    Ok(())
}

/// Set datetime column formats for an Arrow schema
pub fn set_datetime_column_formats(
    worksheet: &mut Worksheet,
    batch: &RecordBatch,
    datetime_fmt: &Format,
) -> PyResult<()> {
    for (col_idx, field) in batch.schema().fields().iter().enumerate() {
        match field.data_type() {
            DataType::Date32 | DataType::Date64 | DataType::Timestamp(_, _) => {
                worksheet
                    .set_column_format(col_idx as u16, datetime_fmt)
                    .map_err(xlsx_err)?;
            }
            _ => {}
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Date/time conversion helpers
// ---------------------------------------------------------------------------

fn days_to_excel_date(days_since_epoch: i64) -> Option<ExcelDateTime> {
    // Unix epoch = 1970-01-01, convert days to y/m/d
    let dt = chrono_from_days(days_since_epoch)?;
    ExcelDateTime::from_ymd(dt.0, dt.1, dt.2).ok()
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

/// Convert days since Unix epoch to (year, month, day)
fn chrono_from_days(days: i64) -> Option<(u16, u8, u8)> {
    // Algorithm from Howard Hinnant's civil_from_days
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
