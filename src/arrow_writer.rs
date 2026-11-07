use arrow_array::cast::AsArray;
use arrow_array::types::*;
use arrow_array::{Array, ArrayRef, BooleanArray, RecordBatch};
use arrow_schema::{DataType, Schema, TimeUnit};
use pyo3::prelude::*;
use rust_xlsxwriter::{ExcelDateTime, Format, Worksheet};

use crate::helpers::{
    write_bool_opt, write_csv_escaped_guarded, write_num, write_number_opt, write_string_opt,
};
use crate::options::value_err;
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
    /// The flag is set when the column has a timezone. Arrow stores those
    /// values as UTC instants.
    Timestamp(TimeUnit, bool),
    /// Time of day. The unit also picks the array type: seconds and
    /// milliseconds are `Time32`, microseconds and nanoseconds `Time64`.
    Time(TimeUnit),
    /// Decimals carry their scale; the value is written as the nearest `f64`.
    Decimal32(i8),
    Decimal64(i8),
    Decimal128(i8),
    Decimal256(i8),
    /// The all-null type. Every cell is blank.
    Null,
}

/// The error for a column this writer cannot represent.
///
/// These columns used to be written as empty cells with no error, so the
/// message says so: a caller whose export "worked" before needs to know the
/// earlier files were missing this column.
fn unsupported(name: &str, dt: &DataType, advice: &str) -> PyErr {
    value_err(format!(
        "column '{name}' has Arrow type {dt}, which cannot be written. {advice} \
         (Earlier versions wrote this column as empty cells without an error.)"
    ))
}

/// Map a column's Arrow type onto the accessor that reads it, or refuse it.
///
/// A dictionary column (pandas `category`, polars `Categorical` and `Enum`) is
/// classified by its value type; [`Col::new`] unwraps the dictionary.
fn classify(name: &str, dt: &DataType) -> PyResult<ColKind> {
    Ok(match dt {
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
        DataType::Timestamp(unit, tz) => ColKind::Timestamp(*unit, tz.is_some()),
        // Only these four pairings are valid Arrow. Matching them exactly
        // matters: `as_primitive` panics when the array is not the type asked.
        DataType::Time32(unit @ (TimeUnit::Second | TimeUnit::Millisecond))
        | DataType::Time64(unit @ (TimeUnit::Microsecond | TimeUnit::Nanosecond)) => {
            ColKind::Time(*unit)
        }
        DataType::Decimal32(_, scale) => ColKind::Decimal32(*scale),
        DataType::Decimal64(_, scale) => ColKind::Decimal64(*scale),
        DataType::Decimal128(_, scale) => ColKind::Decimal128(*scale),
        DataType::Decimal256(_, scale) => ColKind::Decimal256(*scale),
        DataType::Null => ColKind::Null,
        DataType::Dictionary(_, value) if !matches!(**value, DataType::Dictionary(..)) => {
            classify(name, value)?
        }
        DataType::Duration(_) => {
            return Err(unsupported(
                name,
                dt,
                "Excel and CSV have no duration type, and a guess at the unit would be \
                 silently wrong. Cast the column to a string, or to a number of seconds, first.",
            ))
        }
        other => {
            return Err(unsupported(
                name,
                other,
                "Convert the column to a string or a number first.",
            ))
        }
    })
}

/// Refuse a schema with a column the writer cannot represent, before anything
/// is written.
///
/// Checking up front keeps a bad column from leaving half a file behind, and
/// covers a frame with no rows, which never reaches the per-batch code.
/// `indices` limits the check to the selected columns.
pub fn check_schema(schema: &Schema, indices: Option<&[usize]>) -> PyResult<()> {
    let fields = schema.fields();
    let check = |i: usize| classify(fields[i].name(), fields[i].data_type()).map(|_| ());
    match indices {
        Some(ix) => ix.iter().try_for_each(|&i| check(i)),
        None => (0..fields.len()).try_for_each(check),
    }
}

/// One column of a batch, ready to read cell by cell.
///
/// For a dictionary column, `values` is the dictionary, `keys` maps each row to
/// its entry in it, and `dict` is the column itself, whose nulls (a null key)
/// `values` cannot show. Otherwise `values` is the column and a row is its own
/// position.
struct Col {
    values: ArrayRef,
    keys: Option<Vec<usize>>,
    dict: Option<ArrayRef>,
    kind: ColKind,
}

impl Col {
    fn new(name: &str, array: &ArrayRef) -> PyResult<Col> {
        let kind = classify(name, array.data_type())?;
        Ok(match array.as_any_dictionary_opt() {
            Some(dict) => Col {
                values: dict.values().clone(),
                // `normalized_keys` panics on an empty dictionary. Only an
                // all-null column has one, and with no keys every row reads
                // as null in `at`.
                keys: Some(if dict.values().is_empty() {
                    Vec::new()
                } else {
                    dict.normalized_keys()
                }),
                dict: Some(array.clone()),
                kind,
            },
            None => Col {
                values: array.clone(),
                keys: None,
                dict: None,
                kind,
            },
        })
    }

    /// Where row `row`'s value sits in `values`, or `None` for a null.
    #[inline]
    fn at(&self, row: usize) -> Option<usize> {
        let idx = match &self.keys {
            None => row,
            Some(keys) => {
                if self.dict.as_ref().is_some_and(|d| d.is_null(row)) {
                    return None;
                }
                *keys.get(row)?
            }
        };
        (!self.values.is_null(idx)).then_some(idx)
    }

    /// The text at `idx`, if this is a string column.
    fn str_at(&self, idx: usize) -> Option<&str> {
        match self.kind {
            ColKind::Utf8 => Some(self.values.as_string::<i32>().value(idx)),
            ColKind::LargeUtf8 => Some(self.values.as_string::<i64>().value(idx)),
            ColKind::Utf8View => Some(self.values.as_string_view().value(idx)),
            _ => None,
        }
    }
}

fn batch_cols(batch: &RecordBatch) -> PyResult<Vec<Col>> {
    let schema = batch.schema();
    (0..batch.num_columns())
        .map(|c| Col::new(schema.field(c).name(), batch.column(c)))
        .collect()
}

/// A decimal value's exact text, with as many fractional digits as its scale.
fn decimal_text<T: DecimalType>(values: &ArrayRef, idx: usize, scale: i8) -> String {
    T::format_decimal(values.as_primitive::<T>().value(idx), 0, scale)
}

/// The nearest `f64` to a decimal value.
///
/// Goes through the decimal's text so the result is correctly rounded: dividing
/// the unscaled integer by a power of ten rounds twice once that integer is
/// past 2^53.
fn decimal_f64<T: DecimalType>(values: &ArrayRef, idx: usize, scale: i8) -> PyResult<f64> {
    let text = decimal_text::<T>(values, idx, scale);
    text.parse()
        .map_err(|_| value_err(format!("cannot read decimal value '{text}' as a number")))
}

fn decimal_at(col: &Col, idx: usize) -> PyResult<Option<f64>> {
    Ok(Some(match col.kind {
        ColKind::Decimal32(s) => decimal_f64::<Decimal32Type>(&col.values, idx, s)?,
        ColKind::Decimal64(s) => decimal_f64::<Decimal64Type>(&col.values, idx, s)?,
        ColKind::Decimal128(s) => decimal_f64::<Decimal128Type>(&col.values, idx, s)?,
        ColKind::Decimal256(s) => decimal_f64::<Decimal256Type>(&col.values, idx, s)?,
        _ => return Ok(None),
    }))
}

/// A time-of-day value as `(hour, minute, second, nanosecond)`.
fn time_parts(values: &ArrayRef, unit: TimeUnit, idx: usize) -> (u32, u32, u32, u32) {
    let nanos: i64 = match unit {
        TimeUnit::Second => {
            values.as_primitive::<Time32SecondType>().value(idx) as i64 * 1_000_000_000
        }
        TimeUnit::Millisecond => {
            values.as_primitive::<Time32MillisecondType>().value(idx) as i64 * 1_000_000
        }
        TimeUnit::Microsecond => values.as_primitive::<Time64MicrosecondType>().value(idx) * 1_000,
        TimeUnit::Nanosecond => values.as_primitive::<Time64NanosecondType>().value(idx),
    };
    // A valid time is within one day; clamp rather than wrap anything else.
    let nanos = nanos.clamp(0, 86_400 * 1_000_000_000 - 1);
    let secs = (nanos / 1_000_000_000) as u32;
    (
        secs / 3600,
        secs % 3600 / 60,
        secs % 60,
        (nanos % 1_000_000_000) as u32,
    )
}

/// Write one temporal cell. `dt_fmt` is only attached when banding is on —
/// otherwise the column format set before the first row already carries the
/// number format.
///
/// `value` is `None` when Excel's calendar cannot hold the value, and that
/// raises `out_of_range()`. It used to write an empty cell, while the same
/// value from a list of dicts raised.
fn write_temporal(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: Option<rust_xlsxwriter::ExcelDateTime>,
    banding: bool,
    dt_fmt: &Format,
    out_of_range: impl FnOnce() -> PyErr,
) -> PyResult<()> {
    match value {
        Some(dt) => {
            crate::helpers::write_datetime_opt(worksheet, row, col, &dt, banding.then_some(dt_fmt))
        }
        None => Err(out_of_range()),
    }
}

/// The error for a date Excel cannot store, naming the column and the value.
fn out_of_range(name: &str, col: &Col, idx: usize) -> PyErr {
    let mut text = Vec::new();
    let reps = crate::helpers::NumReps {
        na: None,
        inf: None,
    };
    let _ = emit_arrow_cell_csv(&mut text, col, idx, false, reps);
    let what = match String::from_utf8(text) {
        Ok(t) if !t.is_empty() => format!("{t} is"),
        // ISO text cannot show a year outside 0..=9999 either.
        _ => "a date before the year 0 or after 9999 is".to_string(),
    };
    value_err(format!(
        "column '{name}': {what} outside Excel's calendar (1900-01-01 to 9999-12-31), \
         so there is no cell value for it. Convert the column to text to keep it. \
         (Earlier versions wrote it as an empty cell without an error.)"
    ))
}

/// Write an Arrow `RecordBatch` starting at `start_row`.
/// Column types are classified once; each cell dispatches on the cached kind.
///
/// `plain` and `banded` carry the per-column `Format`s resolved from
/// `column_formats`. When a column has an explicit format, it wins over the
/// sheet-wide float, datetime and time formats.
#[allow(clippy::too_many_arguments)]
pub fn write_arrow_batch(
    worksheet: &mut Worksheet,
    batch: &RecordBatch,
    start_row: u32,
    plain: &crate::format::RowPalette,
    banded: Option<&crate::format::RowPalette>,
    layout: &crate::layout::SheetLayout,
    url_cols: &[crate::helpers::UrlCol],
    formula_cols: &[crate::formula::FormulaColumn],
    n_data_cols: usize,
) -> PyResult<()> {
    let num_cols = batch.num_columns();
    let num_rows = batch.num_rows();

    let schema = batch.schema();
    let cols = batch_cols(batch)?;
    // Per-column format override is fixed for the whole column *within a
    // palette* — resolve both variants once instead of per cell in the row×col
    // loop below.
    let plain_cols: Vec<Option<&Format>> = (0..num_cols).map(|c| plain.col(c)).collect();
    let banded_cols: Vec<Option<&Format>> = match banded {
        Some(b) => (0..num_cols).map(|c| b.col(c)).collect(),
        None => Vec::new(),
    };
    let banding = layout.band_color.is_some();

    for row in 0..num_rows {
        let row_u32 = start_row + row as u32;
        let use_band = layout.is_banded(row_u32);
        let pal = if use_band {
            banded.unwrap_or(plain)
        } else {
            plain
        };
        let overrides = if use_band && banded.is_some() {
            &banded_cols
        } else {
            &plain_cols
        };
        // On a band row this carries the fill for cells that would otherwise
        // be written unformatted.
        let text_fmt = pal.text.as_ref();

        for col_idx in 0..num_cols {
            let col_u16 = col_idx as u16;
            let col = &cols[col_idx];
            let column = &col.values;
            let col_override = overrides[col_idx].or(text_fmt);

            let Some(i) = col.at(row) else {
                write_string_opt(
                    worksheet,
                    row_u32,
                    col_u16,
                    layout.reps().na_text(),
                    text_fmt,
                )?;
                continue;
            };

            macro_rules! write_int {
                ($ty:ty) => {{
                    let val = column.as_primitive::<$ty>().value(i) as f64;
                    write_number_opt(worksheet, row_u32, col_u16, val, col_override)?;
                }};
            }
            // A url_columns entry turns text cells into links; anything that
            // is not a valid URL falls back to plain text.
            let url_col = url_cols.get(col_idx).copied().unwrap_or_default();
            let as_url = url_col.link;
            // The display text is another column of this same row; only a
            // string column can supply one, anything else shows the URL.
            let url_text = url_col.text_col.and_then(|ti| {
                let c = cols.get(ti)?;
                c.str_at(c.at(row)?)
            });
            macro_rules! write_str {
                ($val:expr) => {{
                    let val: &str = $val;
                    if as_url && !val.is_empty() {
                        crate::helpers::write_url_or_text(
                            worksheet,
                            row_u32,
                            col_u16,
                            val,
                            col_override,
                            url_text,
                        )?
                    } else {
                        write_string_opt(worksheet, row_u32, col_u16, val, col_override)?
                    }
                }};
            }

            match col.kind {
                ColKind::Int8 => write_int!(Int8Type),
                ColKind::Int16 => write_int!(Int16Type),
                ColKind::Int32 => write_int!(Int32Type),
                ColKind::Int64 => write_int!(Int64Type),
                ColKind::UInt8 => write_int!(UInt8Type),
                ColKind::UInt16 => write_int!(UInt16Type),
                ColKind::UInt32 => write_int!(UInt32Type),
                ColKind::UInt64 => write_int!(UInt64Type),
                ColKind::Float16 => {
                    let val = column.as_primitive::<Float16Type>().value(i).to_f64();
                    write_number_opt(worksheet, row_u32, col_u16, val, col_override)?;
                }
                ColKind::Float32 => {
                    let val = column.as_primitive::<Float32Type>().value(i) as f64;
                    write_num(
                        worksheet,
                        row_u32,
                        col_u16,
                        val,
                        overrides[col_idx].or(pal.float.as_ref()),
                        layout.reps(),
                    )?;
                }
                ColKind::Float64 => {
                    let val = column.as_primitive::<Float64Type>().value(i);
                    write_num(
                        worksheet,
                        row_u32,
                        col_u16,
                        val,
                        overrides[col_idx].or(pal.float.as_ref()),
                        layout.reps(),
                    )?;
                }
                ColKind::Decimal32(_)
                | ColKind::Decimal64(_)
                | ColKind::Decimal128(_)
                | ColKind::Decimal256(_) => {
                    let val = decimal_at(col, i)?.expect("decimal kind");
                    write_num(
                        worksheet,
                        row_u32,
                        col_u16,
                        val,
                        overrides[col_idx].or(pal.float.as_ref()),
                        layout.reps(),
                    )?;
                }
                ColKind::Bool => {
                    let arr = column
                        .as_any()
                        .downcast_ref::<BooleanArray>()
                        .expect("Boolean dtype guarantees downcast");
                    write_bool_opt(worksheet, row_u32, col_u16, arr.value(i), text_fmt)?;
                }
                ColKind::Utf8 | ColKind::LargeUtf8 | ColKind::Utf8View => {
                    write_str!(col.str_at(i).expect("string kind"))
                }
                // With banding the shade has to ride on the cell: the datetime
                // column format set up front is fixed for every row.
                ColKind::Date32 => {
                    let days = column.as_primitive::<Date32Type>().value(i);
                    write_temporal(
                        worksheet,
                        row_u32,
                        col_u16,
                        days_to_excel_date(days as i64),
                        banding,
                        overrides[col_idx].unwrap_or(&pal.datetime),
                        || out_of_range(schema.field(col_idx).name(), col, i),
                    )?;
                }
                ColKind::Date64 => {
                    let ms = column.as_primitive::<Date64Type>().value(i);
                    let (secs, nanos) = split_epoch(ms, 1_000);
                    write_temporal(
                        worksheet,
                        row_u32,
                        col_u16,
                        epoch_to_excel_datetime(secs, nanos),
                        banding,
                        overrides[col_idx].unwrap_or(&pal.datetime),
                        || out_of_range(schema.field(col_idx).name(), col, i),
                    )?;
                }
                // An aware column is written as its UTC wall-clock time: Excel
                // has no timezone, and turning the instant into the column's
                // own local time would take a timezone database.
                ColKind::Timestamp(unit, _) => {
                    let (secs, nanos) = timestamp_parts(column, unit, i);
                    write_temporal(
                        worksheet,
                        row_u32,
                        col_u16,
                        epoch_to_excel_datetime(secs, nanos),
                        banding,
                        overrides[col_idx].unwrap_or(&pal.datetime),
                        || out_of_range(schema.field(col_idx).name(), col, i),
                    )?;
                }
                // No column format is set for times, so the format always
                // rides on the cell.
                ColKind::Time(unit) => {
                    let (h, m, s, ns) = time_parts(column, unit, i);
                    write_temporal(
                        worksheet,
                        row_u32,
                        col_u16,
                        crate::helpers::excel_time(h, m, s, ns),
                        true,
                        overrides[col_idx].unwrap_or(&pal.time),
                        || out_of_range(schema.field(col_idx).name(), col, i),
                    )?;
                }
                ColKind::Null => {
                    write_string_opt(worksheet, row_u32, col_u16, "", text_fmt)?;
                }
            }
        }
        // Written inside the row loop: constant-memory mode flushes each row as
        // the next begins, so a pass afterwards would be ignored.
        if !formula_cols.is_empty() {
            crate::formula::write_formula_row(
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

/// Emit an Arrow `RecordBatch` as CSV rows (no header — caller writes it).
/// Zero-copy over the Arrow buffers; only the output bytes are newly
/// allocated. `row_done` runs after each row, which is where the caller may
/// pass `output` on to the file.
pub fn write_arrow_batch_csv<F>(
    output: &mut Vec<u8>,
    batch: &RecordBatch,
    delim: u8,
    sanitize: bool,
    reps: crate::helpers::NumReps<'_>,
    row_done: &mut F,
) -> PyResult<()>
where
    F: FnMut(&mut Vec<u8>) -> PyResult<()>,
{
    let num_rows = batch.num_rows();
    let cols = batch_cols(batch)?;

    for row in 0..num_rows {
        for (col_idx, col) in cols.iter().enumerate() {
            if col_idx > 0 {
                output.push(delim);
            }
            let Some(i) = col.at(row) else {
                if let Some(text) = reps.na {
                    crate::helpers::write_csv_escaped_guarded(output, text, sanitize);
                }
                continue;
            };
            emit_arrow_cell_csv(output, col, i, sanitize, reps)?;
        }
        output.push(b'\n');
        row_done(output)?;
    }
    Ok(())
}

fn emit_arrow_cell_csv(
    output: &mut Vec<u8>,
    col: &Col,
    i: usize,
    sanitize: bool,
    reps: crate::helpers::NumReps<'_>,
) -> PyResult<()> {
    let column = &col.values;

    macro_rules! emit_int {
        ($ty:ty) => {{
            let val = column.as_primitive::<$ty>().value(i) as i64;
            let mut buf = itoa::Buffer::new();
            output.extend_from_slice(buf.format(val).as_bytes());
        }};
    }

    macro_rules! emit_float {
        ($val:expr) => {{
            let v = $val;
            if v.is_nan() || v.is_infinite() {
                // Without a representation the field stays empty, as it always
                // has — matching the Records path exactly.
                if let Some(text) = reps.text_for(v) {
                    crate::helpers::write_csv_escaped_guarded(output, &text, sanitize);
                }
            } else {
                let mut buf = ryu::Buffer::new();
                output.extend_from_slice(buf.format(v).as_bytes());
            }
        }};
    }

    match col.kind {
        ColKind::Int8 => emit_int!(Int8Type),
        ColKind::Int16 => emit_int!(Int16Type),
        ColKind::Int32 => emit_int!(Int32Type),
        ColKind::Int64 => emit_int!(Int64Type),
        ColKind::UInt8 => emit_int!(UInt8Type),
        ColKind::UInt16 => emit_int!(UInt16Type),
        ColKind::UInt32 => emit_int!(UInt32Type),
        ColKind::UInt64 => emit_int!(UInt64Type),
        ColKind::Float16 => {
            let val = column.as_primitive::<Float16Type>().value(i).to_f64();
            emit_float!(val);
        }
        ColKind::Float32 => emit_float!(column.as_primitive::<Float32Type>().value(i) as f64),
        ColKind::Float64 => emit_float!(column.as_primitive::<Float64Type>().value(i)),
        // The exact digits: the scale is in the dtype, so no float comes
        // in between, and `1.10` stays `1.10`.
        ColKind::Decimal32(s) => {
            output.extend_from_slice(decimal_text::<Decimal32Type>(column, i, s).as_bytes())
        }
        ColKind::Decimal64(s) => {
            output.extend_from_slice(decimal_text::<Decimal64Type>(column, i, s).as_bytes())
        }
        ColKind::Decimal128(s) => {
            output.extend_from_slice(decimal_text::<Decimal128Type>(column, i, s).as_bytes())
        }
        ColKind::Decimal256(s) => {
            output.extend_from_slice(decimal_text::<Decimal256Type>(column, i, s).as_bytes())
        }
        ColKind::Bool => {
            let arr = column
                .as_any()
                .downcast_ref::<BooleanArray>()
                .expect("Boolean dtype guarantees downcast");
            output.extend_from_slice(if arr.value(i) { b"true" } else { b"false" });
        }
        ColKind::Utf8 | ColKind::LargeUtf8 | ColKind::Utf8View => {
            write_csv_escaped_guarded(output, col.str_at(i).expect("string kind"), sanitize)
        }
        ColKind::Date32 => {
            let days = column.as_primitive::<Date32Type>().value(i);
            if let Some((y, m, d)) = chrono_from_days(days as i64) {
                crate::helpers::write_iso_date(output, u32::from(y), u32::from(m), u32::from(d));
            }
        }
        ColKind::Date64 => {
            let (secs, nanos) = split_epoch(column.as_primitive::<Date64Type>().value(i), 1_000);
            emit_timestamp_csv(output, secs, nanos, false);
        }
        ColKind::Timestamp(unit, aware) => {
            let (secs, nanos) = timestamp_parts(column, unit, i);
            emit_timestamp_csv(output, secs, nanos, aware);
        }
        ColKind::Time(unit) => {
            let (h, m, s, ns) = time_parts(column, unit, i);
            crate::helpers::write_iso_time(output, h, m, s, ns);
        }
        ColKind::Null => {}
    }
    Ok(())
}

/// Split a count of `per_sec` ticks since the epoch into whole seconds and
/// nanoseconds. Floor division, so a moment before 1970 keeps a fraction in
/// `0..1s` instead of a negative one that would wrap to the wrong second.
fn split_epoch(ticks: i64, per_sec: i64) -> (i64, u32) {
    let nanos_per_tick = 1_000_000_000 / per_sec;
    (
        ticks.div_euclid(per_sec),
        (ticks.rem_euclid(per_sec) * nanos_per_tick) as u32,
    )
}

/// An Arrow timestamp as `(seconds since epoch, nanoseconds)`, whatever its
/// stored `TimeUnit`.
fn timestamp_parts(column: &ArrayRef, unit: TimeUnit, row: usize) -> (i64, u32) {
    match unit {
        TimeUnit::Second => (column.as_primitive::<TimestampSecondType>().value(row), 0),
        TimeUnit::Millisecond => split_epoch(
            column.as_primitive::<TimestampMillisecondType>().value(row),
            1_000,
        ),
        TimeUnit::Microsecond => split_epoch(
            column.as_primitive::<TimestampMicrosecondType>().value(row),
            1_000_000,
        ),
        TimeUnit::Nanosecond => split_epoch(
            column.as_primitive::<TimestampNanosecondType>().value(row),
            1_000_000_000,
        ),
    }
}

/// Decompose seconds since the epoch into civil `(year, month, day, hour,
/// minute, second)`. Returns `None` if the year is outside 0..=9999.
fn epoch_to_civil(secs: i64) -> Option<(u16, u8, u8, u32, u32, u32)> {
    let days = secs.div_euclid(86_400);
    let day_secs = secs.rem_euclid(86_400) as u32;
    let (year, month, day) = chrono_from_days(days)?;
    Some((
        year,
        month,
        day,
        day_secs / 3600,
        day_secs % 3600 / 60,
        day_secs % 60,
    ))
}

/// ISO-8601 with the fraction when there is one. An aware column's values are
/// UTC instants, so they are written in UTC and say so with `+00:00`: the text
/// then names the same instant the column holds. Writing the column's own
/// local time instead would take a timezone database.
fn emit_timestamp_csv(output: &mut Vec<u8>, secs: i64, nanos: u32, aware: bool) {
    let Some((y, m, d, hour, minute, second)) = epoch_to_civil(secs) else {
        return;
    };
    crate::helpers::write_iso_date(output, u32::from(y), u32::from(m), u32::from(d));
    output.push(b'T');
    crate::helpers::write_iso_time(output, hour, minute, second, nanos);
    if aware {
        crate::helpers::write_iso_offset(output, 0);
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

fn epoch_to_excel_datetime(secs: i64, nanos: u32) -> Option<ExcelDateTime> {
    let (year, month, day, hour, minute, second) = epoch_to_civil(secs)?;
    let whole = ExcelDateTime::from_ymd(year, month, day)
        .ok()?
        .and_hms(hour as u16, minute as u8, second)
        .ok()?;
    Some(crate::helpers::with_subsecond(whole, nanos))
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

    if !(0..=9999).contains(&y) {
        return None;
    }
    Some((y as u16, m as u8, d as u8))
}
