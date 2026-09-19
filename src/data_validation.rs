//! Data validation, parsed from a `data_validations` mapping.
//!
//! Rules are given per column and cover that column's data rows, so — like
//! the conditional formats — they are applied after the data, once the row
//! count is known. The common case by a wide margin is a dropdown:
//! `{"status": {"type": "list", "values": ["open", "closed"]}}`.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{
    DataValidation, DataValidationErrorStyle, DataValidationRule, Formula, Worksheet,
};

use crate::worksheet::xlsx_err;

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

const TYPES: [&str; 6] = [
    "list",
    "whole_number",
    "decimal",
    "text_length",
    "custom",
    "any",
];

const CRITERIA: &str = "==, !=, >, >=, <, <=, between, not_between";

#[derive(Default)]
pub struct DataValidations(Vec<(String, DataValidation)>);

fn get<'py>(map: &Bound<'py, PyDict>, key: &str) -> PyResult<Option<Bound<'py, PyAny>>> {
    map.get_item(key)
}

fn need<'py>(map: &Bound<'py, PyDict>, key: &str, kind: &str) -> PyResult<Bound<'py, PyAny>> {
    get(map, key)?
        .ok_or_else(|| value_err(format!("data_validations: a '{kind}' rule needs '{key}'")))
}

/// The comparison a numeric rule uses, before its values are narrowed.
#[derive(Clone, Copy)]
enum Cmp {
    EqualTo,
    NotEqualTo,
    GreaterThan,
    GreaterThanOrEqualTo,
    LessThan,
    LessThanOrEqualTo,
    Between,
    NotBetween,
}

/// Criteria and values, read once as `f64`.
///
/// The three numeric rule kinds differ only in the type they end up holding,
/// so the parsing happens here and each kind narrows the result itself —
/// which also gives a place to reject 2.5 for a whole-number rule instead of
/// quietly truncating it.
fn numeric_parts(map: &Bound<'_, PyDict>, kind: &str) -> PyResult<(Cmp, f64, f64)> {
    let criteria: String = need(map, "criteria", kind)?.extract()?;
    let cmp = match criteria.as_str() {
        "==" | "equal_to" => Cmp::EqualTo,
        "!=" | "not_equal_to" => Cmp::NotEqualTo,
        ">" | "greater_than" => Cmp::GreaterThan,
        ">=" => Cmp::GreaterThanOrEqualTo,
        "<" | "less_than" => Cmp::LessThan,
        "<=" => Cmp::LessThanOrEqualTo,
        "between" => Cmp::Between,
        "not_between" => Cmp::NotBetween,
        other => {
            return Err(value_err(format!(
                "data_validations: unknown criteria '{other}' for a '{kind}' rule \
                 (expected {CRITERIA})"
            )))
        }
    };
    match cmp {
        Cmp::Between | Cmp::NotBetween => Ok((
            cmp,
            need(map, "min", kind)?.extract()?,
            need(map, "max", kind)?.extract()?,
        )),
        _ => {
            let v: f64 = need(map, "value", kind)?.extract()?;
            Ok((cmp, v, v))
        }
    }
}

fn build_rule<T: rust_xlsxwriter::IntoDataValidationValue>(
    cmp: Cmp,
    a: T,
    b: T,
) -> DataValidationRule<T> {
    match cmp {
        Cmp::EqualTo => DataValidationRule::EqualTo(a),
        Cmp::NotEqualTo => DataValidationRule::NotEqualTo(a),
        Cmp::GreaterThan => DataValidationRule::GreaterThan(a),
        Cmp::GreaterThanOrEqualTo => DataValidationRule::GreaterThanOrEqualTo(a),
        Cmp::LessThan => DataValidationRule::LessThan(a),
        Cmp::LessThanOrEqualTo => DataValidationRule::LessThanOrEqualTo(a),
        Cmp::Between => DataValidationRule::Between(a, b),
        Cmp::NotBetween => DataValidationRule::NotBetween(a, b),
    }
}

/// Narrow to a whole number, refusing a fraction rather than truncating it.
fn whole(v: f64, kind: &str) -> PyResult<i64> {
    if v.fract() != 0.0 {
        return Err(value_err(format!(
            "data_validations: a '{kind}' rule needs whole numbers; got {v}"
        )));
    }
    Ok(v as i64)
}

/// The messages and flags every rule type accepts.
fn decorate(mut dv: DataValidation, map: &Bound<'_, PyDict>) -> PyResult<DataValidation> {
    if let Some(v) = get(map, "input_title")? {
        dv = dv
            .set_input_title(v.extract::<String>()?)
            .map_err(xlsx_err)?;
    }
    if let Some(v) = get(map, "input_message")? {
        dv = dv
            .set_input_message(v.extract::<String>()?)
            .map_err(xlsx_err)?;
    }
    if let Some(v) = get(map, "error_title")? {
        dv = dv
            .set_error_title(v.extract::<String>()?)
            .map_err(xlsx_err)?;
    }
    if let Some(v) = get(map, "error_message")? {
        dv = dv
            .set_error_message(v.extract::<String>()?)
            .map_err(xlsx_err)?;
    }
    if let Some(v) = get(map, "error_style")? {
        let name: String = v.extract()?;
        dv = dv.set_error_style(match name.as_str() {
            "stop" => DataValidationErrorStyle::Stop,
            "warning" => DataValidationErrorStyle::Warning,
            "information" => DataValidationErrorStyle::Information,
            other => {
                return Err(value_err(format!(
                    "data_validations: unknown error_style '{other}' \
                     (expected stop, warning, information)"
                )))
            }
        });
    }
    if let Some(v) = get(map, "ignore_blank")? {
        dv = dv.ignore_blank(v.extract()?);
    }
    if let Some(v) = get(map, "show_dropdown")? {
        dv = dv.show_dropdown(v.extract()?);
    }
    Ok(dv)
}

fn parse_rule(map: &Bound<'_, PyDict>) -> PyResult<DataValidation> {
    let kind: String = need(map, "type", "data validation")?.extract()?;
    let dv = DataValidation::new();

    let dv = match kind.as_str() {
        "list" => {
            let values: Vec<String> = need(map, "values", "list")?.extract().map_err(|_| {
                value_err(
                    "data_validations: a 'list' rule needs 'values' as a list of strings"
                        .to_string(),
                )
            })?;
            // Excel caps the inline list at 255 characters including
            // separators; the crate reports that rather than writing a file
            // Excel will refuse to open.
            dv.allow_list_strings(&values).map_err(xlsx_err)?
        }
        "whole_number" => {
            let (cmp, a, b) = numeric_parts(map, "whole_number")?;
            dv.allow_whole_number(build_rule(
                cmp,
                whole(a, "whole_number")? as i32,
                whole(b, "whole_number")? as i32,
            ))
        }
        "decimal" => {
            let (cmp, a, b) = numeric_parts(map, "decimal")?;
            dv.allow_decimal_number(build_rule(cmp, a, b))
        }
        "text_length" => {
            let (cmp, a, b) = numeric_parts(map, "text_length")?;
            dv.allow_text_length(build_rule(
                cmp,
                whole(a, "text_length")? as u32,
                whole(b, "text_length")? as u32,
            ))
        }
        "custom" => {
            let formula: String = need(map, "formula", "custom")?.extract()?;
            dv.allow_custom(Formula::new(formula))
        }
        "any" => dv.allow_any_value(),
        other => {
            return Err(value_err(format!(
                "data_validations: unknown type '{other}' (expected one of {})",
                TYPES.join(", ")
            )))
        }
    };
    decorate(dv, map)
}

impl DataValidations {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(DataValidations(out));
        };
        let map = spec.cast::<PyDict>().map_err(|_| {
            value_err("data_validations must be a dict keyed by column name".to_string())
        })?;

        for (key, value) in map.iter() {
            let column: String = key.extract()?;
            let rule = value.cast::<PyDict>().map_err(|_| {
                value_err(format!(
                    "data_validations: the rule for '{column}' must be a dict"
                ))
            })?;
            out.push((column, parse_rule(rule)?));
        }
        Ok(DataValidations(out))
    }

    /// An unknown column warns and is skipped, like the other column-keyed
    /// options: a missing dropdown is not worth failing an export over.
    pub fn apply(
        &self,
        worksheet: &mut Worksheet,
        headers: &[String],
        header_row: u32,
        data_rows: u32,
        py: Python,
    ) -> PyResult<()> {
        if self.0.is_empty() || data_rows == 0 {
            return Ok(());
        }
        let warnings = py.import("warnings")?;
        let first = header_row + 1;
        let last = header_row + data_rows;

        for (column, validation) in &self.0 {
            let Some(idx) = headers.iter().position(|h| h == column) else {
                warnings.call_method1(
                    "warn",
                    (format!(
                        "data_validations: unknown column '{column}', skipped"
                    ),),
                )?;
                continue;
            };
            let col = idx as u16;
            worksheet
                .add_data_validation(first, col, last, col, validation)
                .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
