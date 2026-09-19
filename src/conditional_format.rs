//! Conditional formatting, parsed from a `conditional_formats` mapping.
//!
//! Rules are given per column name and applied to that column's data rows —
//! never the header — which means the range is only known once the last row
//! has been written. So, like the autofilter and the totals row, these are
//! applied after the data rather than in [`crate::helpers::SheetLayout::apply`].
//! That is safe under constant-memory mode: a conditional format is worksheet
//! metadata written at save time, not a cell write, so no flushed row is
//! revisited.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatAverage,
    ConditionalFormatAverageRule, ConditionalFormatCell,
    ConditionalFormatCellRule, ConditionalFormatDataBar, ConditionalFormatDuplicate,
    ConditionalFormatText, ConditionalFormatTextRule, ConditionalFormatTop,
    ConditionalFormatTopRule, Format as XlsxFormat, Worksheet,
};

use crate::format::parse_color;
use crate::worksheet::xlsx_err;

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

const TYPES: [&str; 9] = [
    "cell",
    "data_bar",
    "2_color_scale",
    "3_color_scale",
    "text",
    "top",
    "average",
    "duplicate",
    "unique",
];

/// One parsed rule, ready to hand to `add_conditional_format`.
///
/// The crate's rule types do not share an object-safe trait, so they are held
/// as an enum rather than a `Box<dyn ConditionalFormat>`.
pub enum Rule {
    Cell(ConditionalFormatCell),
    DataBar(ConditionalFormatDataBar),
    TwoColor(ConditionalFormat2ColorScale),
    ThreeColor(ConditionalFormat3ColorScale),
    Text(ConditionalFormatText),
    Top(ConditionalFormatTop),
    Average(ConditionalFormatAverage),
    Duplicate(ConditionalFormatDuplicate),
}

/// Rules for one column, by header name.
#[derive(Default)]
pub struct ConditionalFormats(Vec<(String, Rule)>);

fn get<'py>(map: &Bound<'py, PyDict>, key: &str) -> PyResult<Option<Bound<'py, PyAny>>> {
    map.get_item(key)
}

fn need<'py>(map: &Bound<'py, PyDict>, key: &str, kind: &str) -> PyResult<Bound<'py, PyAny>> {
    get(map, key)?.ok_or_else(|| {
        value_err(format!(
            "conditional_formats: a '{kind}' rule needs '{key}'"
        ))
    })
}

fn format_of(map: &Bound<'_, PyDict>) -> PyResult<Option<XlsxFormat>> {
    match get(map, "format")? {
        Some(obj) => Ok(Some(
            obj.extract::<PyRef<'_, crate::format::Format>>()?.inner.clone(),
        )),
        None => Ok(None),
    }
}

fn color_of(map: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<rust_xlsxwriter::Color>> {
    match get(map, key)? {
        Some(obj) => Ok(Some(parse_color(&obj.extract::<String>()?)?)),
        None => Ok(None),
    }
}

/// Comparison values are numeric. A text comparison is the `text` type,
/// which has its own Excel rule rather than a string shoved through this one.
fn cell_rule(map: &Bound<'_, PyDict>) -> PyResult<ConditionalFormatCellRule<f64>> {
    let criteria: String = need(map, "criteria", "cell")?.extract()?;
    match criteria.as_str() {
        "between" | "not_between" => {
            let min: f64 = need(map, "min", "cell")?.extract()?;
            let max: f64 = need(map, "max", "cell")?.extract()?;
            Ok(if criteria == "between" {
                ConditionalFormatCellRule::Between(min, max)
            } else {
                ConditionalFormatCellRule::NotBetween(min, max)
            })
        }
        _ => {
            let v: f64 = need(map, "value", "cell")?.extract()?;
            Ok(match criteria.as_str() {
                "==" | "equal_to" => ConditionalFormatCellRule::EqualTo(v),
                "!=" | "not_equal_to" => ConditionalFormatCellRule::NotEqualTo(v),
                ">" | "greater_than" => ConditionalFormatCellRule::GreaterThan(v),
                ">=" => ConditionalFormatCellRule::GreaterThanOrEqualTo(v),
                "<" | "less_than" => ConditionalFormatCellRule::LessThan(v),
                "<=" => ConditionalFormatCellRule::LessThanOrEqualTo(v),
                other => {
                    return Err(value_err(format!(
                        "conditional_formats: unknown criteria '{other}' for a 'cell' rule \
                         (expected ==, !=, >, >=, <, <=, between, not_between)"
                    )))
                }
            })
        }
    }
}

fn text_rule(map: &Bound<'_, PyDict>) -> PyResult<ConditionalFormatTextRule> {
    let criteria: String = need(map, "criteria", "text")?.extract()?;
    let value: String = need(map, "value", "text")?.extract()?;
    Ok(match criteria.as_str() {
        "contains" => ConditionalFormatTextRule::Contains(value),
        "does_not_contain" => ConditionalFormatTextRule::DoesNotContain(value),
        "begins_with" => ConditionalFormatTextRule::BeginsWith(value),
        "ends_with" => ConditionalFormatTextRule::EndsWith(value),
        other => {
            return Err(value_err(format!(
                "conditional_formats: unknown criteria '{other}' for a 'text' rule \
                 (expected contains, does_not_contain, begins_with, ends_with)"
            )))
        }
    })
}

fn top_rule(map: &Bound<'_, PyDict>) -> PyResult<ConditionalFormatTopRule> {
    let criteria: String = match get(map, "criteria")? {
        Some(c) => c.extract()?,
        None => "top".to_string(),
    };
    let n: u16 = match get(map, "value")? {
        Some(v) => v.extract()?,
        None => 10,
    };
    Ok(match criteria.as_str() {
        "top" => ConditionalFormatTopRule::Top(n),
        "bottom" => ConditionalFormatTopRule::Bottom(n),
        "top_percent" => ConditionalFormatTopRule::TopPercent(n),
        "bottom_percent" => ConditionalFormatTopRule::BottomPercent(n),
        other => {
            return Err(value_err(format!(
                "conditional_formats: unknown criteria '{other}' for a 'top' rule \
                 (expected top, bottom, top_percent, bottom_percent)"
            )))
        }
    })
}

fn average_rule(map: &Bound<'_, PyDict>) -> PyResult<ConditionalFormatAverageRule> {
    let criteria: String = match get(map, "criteria")? {
        Some(c) => c.extract()?,
        None => "above".to_string(),
    };
    Ok(match criteria.as_str() {
        "above" => ConditionalFormatAverageRule::AboveAverage,
        "below" => ConditionalFormatAverageRule::BelowAverage,
        "equal_or_above" => ConditionalFormatAverageRule::EqualOrAboveAverage,
        "equal_or_below" => ConditionalFormatAverageRule::EqualOrBelowAverage,
        other => {
            return Err(value_err(format!(
                "conditional_formats: unknown criteria '{other}' for an 'average' rule \
                 (expected above, below, equal_or_above, equal_or_below)"
            )))
        }
    })
}

fn parse_rule(map: &Bound<'_, PyDict>) -> PyResult<Rule> {
    let kind: String = need(map, "type", "conditional format")?.extract()?;
    let fmt = format_of(map)?;

    Ok(match kind.as_str() {
        "cell" => {
            let mut r = ConditionalFormatCell::new().set_rule(cell_rule(map)?);
            if let Some(f) = fmt {
                r = r.set_format(f);
            }
            Rule::Cell(r)
        }
        "data_bar" => {
            let mut r = ConditionalFormatDataBar::new();
            if let Some(c) = color_of(map, "color")? {
                r = r.set_fill_color(c);
            }
            if let Some(v) = get(map, "bar_only")? {
                r = r.set_bar_only(v.extract()?);
            }
            Rule::DataBar(r)
        }
        "2_color_scale" => {
            let mut r = ConditionalFormat2ColorScale::new();
            if let Some(c) = color_of(map, "min_color")? {
                r = r.set_minimum_color(c);
            }
            if let Some(c) = color_of(map, "max_color")? {
                r = r.set_maximum_color(c);
            }
            Rule::TwoColor(r)
        }
        "3_color_scale" => {
            let mut r = ConditionalFormat3ColorScale::new();
            if let Some(c) = color_of(map, "min_color")? {
                r = r.set_minimum_color(c);
            }
            if let Some(c) = color_of(map, "mid_color")? {
                r = r.set_midpoint_color(c);
            }
            if let Some(c) = color_of(map, "max_color")? {
                r = r.set_maximum_color(c);
            }
            Rule::ThreeColor(r)
        }
        "text" => {
            let mut r = ConditionalFormatText::new().set_rule(text_rule(map)?);
            if let Some(f) = fmt {
                r = r.set_format(f);
            }
            Rule::Text(r)
        }
        "top" => {
            let mut r = ConditionalFormatTop::new().set_rule(top_rule(map)?);
            if let Some(f) = fmt {
                r = r.set_format(f);
            }
            Rule::Top(r)
        }
        "average" => {
            let mut r = ConditionalFormatAverage::new().set_rule(average_rule(map)?);
            if let Some(f) = fmt {
                r = r.set_format(f);
            }
            Rule::Average(r)
        }
        "duplicate" | "unique" => {
            let mut r = ConditionalFormatDuplicate::new();
            if kind == "unique" {
                r = r.invert();
            }
            if let Some(f) = fmt {
                r = r.set_format(f);
            }
            Rule::Duplicate(r)
        }
        other => {
            return Err(value_err(format!(
                "conditional_formats: unknown type '{other}' (expected one of {})",
                TYPES.join(", ")
            )))
        }
    })
}

impl ConditionalFormats {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(ConditionalFormats(out));
        };
        let map = spec.cast::<PyDict>().map_err(|_| {
            value_err("conditional_formats must be a dict keyed by column name".to_string())
        })?;

        for (key, value) in map.iter() {
            let column: String = key.extract()?;
            // One rule, or several stacked on the same column.
            let rules = match value.cast::<PyDict>() {
                Ok(single) => vec![parse_rule(single)?],
                Err(_) => {
                    let mut many = Vec::new();
                    for item in value.try_iter().map_err(|_| {
                        value_err(format!(
                            "conditional_formats: '{column}' must be a rule dict or a list of them"
                        ))
                    })? {
                        let item = item?;
                        let rule = item.cast::<PyDict>().map_err(|_| {
                            value_err(format!(
                                "conditional_formats: every rule for '{column}' must be a dict"
                            ))
                        })?;
                        many.push(parse_rule(rule)?);
                    }
                    many
                }
            };
            for rule in rules {
                out.push((column.clone(), rule));
            }
        }
        Ok(ConditionalFormats(out))
    }

    /// Apply every rule to its column's data rows.
    ///
    /// An unknown column name warns and is skipped, matching `column_formats`:
    /// a rule that cannot be placed costs some shading, not the export.
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

        for (column, rule) in &self.0 {
            let Some(idx) = headers.iter().position(|h| h == column) else {
                warnings.call_method1(
                    "warn",
                    (format!(
                        "conditional_formats: unknown column '{column}', skipped"
                    ),),
                )?;
                continue;
            };
            let col = idx as u16;
            match rule {
                Rule::Cell(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::DataBar(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::TwoColor(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::ThreeColor(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::Text(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::Top(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::Average(r) => worksheet.add_conditional_format(first, col, last, col, r),
                Rule::Duplicate(r) => worksheet.add_conditional_format(first, col, last, col, r),
            }
            .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
