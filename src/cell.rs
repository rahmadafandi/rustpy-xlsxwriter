//! Shared Python-value type cascade for the Records paths (Excel and CSV).
//!
//! The detection order is defined exactly once here. Each output target
//! (worksheet cell vs CSV byte buffer) implements [`CellWriter`]; the generic
//! [`classify_and_write`] and [`try_cached`] drive both. This replaces the
//! four hand-kept-in-sync copies the cascade used to live in.

use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateTime, PyFloat, PyInt, PyString};

use crate::helpers::ColType;

/// A sink for one Python scalar. Implementors encode the per-target action for
/// each detected type (formatting, escaping, etc.); the cascade itself only
/// decides which method to call.
pub trait CellWriter {
    fn write_none(&mut self) -> PyResult<()>;
    fn write_str(&mut self, s: &str) -> PyResult<()>;
    fn write_bool(&mut self, b: bool) -> PyResult<()>;
    fn write_float(&mut self, f: f64) -> PyResult<()>;
    fn write_int(&mut self, i: &Bound<'_, PyInt>) -> PyResult<()>;
    fn write_datetime(&mut self, dt: &Bound<'_, PyDateTime>) -> PyResult<()>;
    fn write_date(&mut self, d: &Bound<'_, PyDate>) -> PyResult<()>;
}

/// Full type cascade. Returns the detected [`ColType`] so callers can cache it
/// for the first-row fast path. Order is significant:
/// - `bool` before `int` (Python `bool` is a subclass of `int`),
/// - native casts before numpy-scalar `extract` fallbacks.
pub fn classify_and_write<W: CellWriter>(
    value: &Bound<'_, PyAny>,
    w: &mut W,
) -> PyResult<ColType> {
    if value.is_none() {
        w.write_none()?;
        return Ok(ColType::Unknown);
    }
    if let Ok(s) = value.cast::<PyString>() {
        w.write_str(s.to_str()?)?;
        return Ok(ColType::String);
    }
    // Bool BEFORE Int (Python bool is a subclass of int).
    if let Ok(b) = value.cast::<PyBool>() {
        w.write_bool(b.is_true())?;
        return Ok(ColType::Bool);
    }
    if let Ok(f) = value.cast::<PyFloat>() {
        w.write_float(f.value())?;
        return Ok(ColType::Float);
    }
    if let Ok(i) = value.cast::<PyInt>() {
        w.write_int(i)?;
        return Ok(ColType::Int);
    }
    if let Ok(dt) = value.cast::<PyDateTime>() {
        w.write_datetime(dt)?;
        return Ok(ColType::DateTime);
    }
    // Date AFTER DateTime (datetime is a subclass of date).
    if let Ok(d) = value.cast::<PyDate>() {
        w.write_date(d)?;
        return Ok(ColType::Date);
    }
    // numpy scalar fallback: bool before f64 (numpy.bool_ extracts as f64 too).
    if let Ok(val) = value.extract::<bool>() {
        w.write_bool(val)?;
        return Ok(ColType::Bool);
    }
    if let Ok(val) = value.extract::<f64>() {
        w.write_float(val)?;
        return Ok(ColType::Float);
    }
    w.write_str(&value.to_string())?;
    Ok(ColType::String)
}

/// Fast path using a cached [`ColType`]. Returns `true` if the value matched
/// the cached type and was written; `false` to fall back to
/// [`classify_and_write`].
pub fn try_cached<W: CellWriter>(
    value: &Bound<'_, PyAny>,
    cached: ColType,
    w: &mut W,
) -> PyResult<bool> {
    if value.is_none() {
        w.write_none()?;
        return Ok(true);
    }
    match cached {
        ColType::String => {
            if let Ok(s) = value.cast::<PyString>() {
                w.write_str(s.to_str()?)?;
                return Ok(true);
            }
        }
        ColType::Bool => {
            if let Ok(b) = value.cast::<PyBool>() {
                w.write_bool(b.is_true())?;
                return Ok(true);
            }
        }
        ColType::Float => {
            if let Ok(f) = value.cast::<PyFloat>() {
                w.write_float(f.value())?;
                return Ok(true);
            }
        }
        ColType::Int => {
            // Python bool is a subclass of int and casts to PyInt, so a bool
            // landing in an Int-cached column must miss here and fall back to
            // the cascade (which checks Bool first) — else `True` → `1`.
            if value.cast::<PyBool>().is_err() {
                if let Ok(i) = value.cast::<PyInt>() {
                    w.write_int(i)?;
                    return Ok(true);
                }
            }
        }
        ColType::DateTime => {
            if let Ok(dt) = value.cast::<PyDateTime>() {
                w.write_datetime(dt)?;
                return Ok(true);
            }
        }
        ColType::Date => {
            if let Ok(d) = value.cast::<PyDate>() {
                w.write_date(d)?;
                return Ok(true);
            }
        }
        ColType::Unknown => {}
    }
    Ok(false)
}
