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

            macro_rules! write_int {
                ($ty:ty) => {{
                    let val = column.as_primitive::<$ty>().value(row) as f64;
                    worksheet.write_number(row_u32, col_u16, val).map_err(xlsx_err)?;
                }};
            }

            match kinds[col_idx] {
                ColKind::Int8 => write_int!(Int8Type),
                ColKind::Int16 => write_int!(Int16Type),
                ColKind::Int32 => write_int!(Int32Type),
                ColKind::Int64 => write_int!(Int64Type),
                ColKind::UInt8 => write_int!(UInt8Type),
                ColKind::UInt16 => write_int!(UInt16Type),
                ColKind::UInt32 => write_int!(UInt32Type),
                ColKind::UInt64 => write_int!(UInt64Type),
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

/// Emit an Arrow `RecordBatch` as CSV rows (no header — caller writes it).
/// Zero-copy over the Arrow buffers; only the output bytes are newly
/// allocated.
pub fn write_arrow_batch_csv(
    output: &mut Vec<u8>,
    batch: &RecordBatch,
    delim: u8,
) -> PyResult<()> {
    let num_cols = batch.num_columns();
    let num_rows = batch.num_rows();

    let columns: Vec<ArrayRef> = (0..num_cols).map(|c| batch.column(c).clone()).collect();
    let kinds: Vec<ColKind> = columns.iter().map(|c| classify(c.data_type())).collect();

    for row in 0..num_rows {
        for col_idx in 0..num_cols {
            if col_idx > 0 {
                output.push(delim);
            }
            let column = &columns[col_idx];
            if column.is_null(row) {
                continue;
            }
            emit_arrow_cell_csv(output, column, kinds[col_idx], row);
        }
        output.push(b'\n');
    }
    Ok(())
}

fn emit_arrow_cell_csv(output: &mut Vec<u8>, column: &ArrayRef, kind: ColKind, row: usize) {
    use std::io::Write;

    macro_rules! emit_int {
        ($ty:ty) => {{
            let val = column.as_primitive::<$ty>().value(row) as i64;
            let mut buf = itoa::Buffer::new();
            output.extend_from_slice(buf.format(val).as_bytes());
        }};
    }

    macro_rules! emit_float {
        ($val:expr) => {{
            let v = $val;
            if !v.is_nan() && !v.is_infinite() {
                let mut buf = ryu::Buffer::new();
                output.extend_from_slice(buf.format(v).as_bytes());
            }
        }};
    }

    match kind {
        ColKind::Int8 => emit_int!(Int8Type),
        ColKind::Int16 => emit_int!(Int16Type),
        ColKind::Int32 => emit_int!(Int32Type),
        ColKind::Int64 => emit_int!(Int64Type),
        ColKind::UInt8 => emit_int!(UInt8Type),
        ColKind::UInt16 => emit_int!(UInt16Type),
        ColKind::UInt32 => emit_int!(UInt32Type),
        ColKind::UInt64 => emit_int!(UInt64Type),
        ColKind::Float16 => {
            let val = column.as_primitive::<Float16Type>().value(row).to_f64();
            emit_float!(val);
        }
        ColKind::Float32 => emit_float!(column.as_primitive::<Float32Type>().value(row) as f64),
        ColKind::Float64 => emit_float!(column.as_primitive::<Float64Type>().value(row)),
        ColKind::Bool => {
            let arr = column
                .as_any()
                .downcast_ref::<BooleanArray>()
                .expect("Boolean dtype guarantees downcast");
            output.extend_from_slice(if arr.value(row) { b"true" } else { b"false" });
        }
        ColKind::Utf8 => write_csv_escaped(output, column.as_string::<i32>().value(row)),
        ColKind::LargeUtf8 => write_csv_escaped(output, column.as_string::<i64>().value(row)),
        ColKind::Utf8View => write_csv_escaped(output, column.as_string_view().value(row)),
        ColKind::Date32 => {
            let days = column.as_primitive::<Date32Type>().value(row);
            if let Some((y, m, d)) = chrono_from_days(days as i64) {
                let _ = write!(output, "{:04}-{:02}-{:02}", y, m, d);
            }
        }
        ColKind::Date64 => {
            let ms = column.as_primitive::<Date64Type>().value(row);
            emit_timestamp_csv(output, ms * 1000);
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
            emit_timestamp_csv(output, micros);
        }
        ColKind::Unsupported => {}
    }
}

fn emit_timestamp_csv(output: &mut Vec<u8>, micros: i64) {
    use std::io::Write;
    let total_secs = micros / 1_000_000;
    let days = total_secs / 86400;
    let day_secs = (total_secs % 86400).unsigned_abs();
    let Some((y, m, d)) = chrono_from_days(days) else {
        return;
    };
    let hour = day_secs / 3600;
    let minute = (day_secs % 3600) / 60;
    let second = day_secs % 60;
    let _ = write!(
        output,
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
        y, m, d, hour, minute, second
    );
}

fn write_csv_escaped(output: &mut Vec<u8>, val: &str) {
    if val.contains(',')
        || val.contains('\n')
        || val.contains('\r')
        || val.contains('"')
    {
        output.push(b'"');
        for b in val.bytes() {
            if b == b'"' {
                output.push(b'"');
            }
            output.push(b);
        }
        output.push(b'"');
    } else {
        output.extend_from_slice(val.as_bytes());
    }
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
