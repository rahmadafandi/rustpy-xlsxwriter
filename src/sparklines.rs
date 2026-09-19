//! Sparklines — a one-cell chart per row, parsed from a `sparklines` mapping.
//!
//! The useful shape for a data table is a trend column: leave a column empty
//! in the records, point it at the span it should summarise, and every data
//! row gets its own tiny chart.
//!
//! ```python
//! rows = [{"q1": 1, "q2": 5, "q3": 3, "q4": 8, "trend": None}, ...]
//! sparklines={"trend": {"from": "q1", "to": "q4"}}
//! ```
//!
//! Pointing at a column that already exists is deliberate. Appending one would
//! mean reaching into the header assembly and the column accounting that
//! `formula_columns` uses, in both the records and the Arrow row loops — a far
//! larger change than the feature is worth, and an empty column in the data is
//! no hardship to add.
//!
//! Applied after the data, since the group spans the rows actually written.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{Sparkline, SparklineType, Worksheet};

use crate::format::parse_color;
use crate::options::value_err;
use crate::worksheet::xlsx_err;

const KEYS: [&str; 13] = [
    "from",
    "to",
    "type",
    "color",
    "style",
    "high_point",
    "low_point",
    "first_point",
    "last_point",
    "markers",
    "negative_points",
    "axis",
    "right_to_left",
];

/// A sparkline plus the columns it reads and lives in.
pub struct Spark {
    /// Column the charts are drawn in.
    column: String,
    /// First and last column of the span each row summarises.
    from: String,
    to: String,
    sparkline: Sparkline,
}

#[derive(Default)]
pub struct Sparklines(Vec<Spark>);

fn build(map: &Bound<'_, PyDict>, column: &str) -> PyResult<Spark> {
    for key in map.keys().iter() {
        let name: String = key.extract()?;
        if !KEYS.contains(&name.as_str()) {
            return Err(value_err(format!(
                "sparklines: unknown key '{name}' for '{column}' (expected one of {})",
                KEYS.join(", ")
            )));
        }
    }
    let need = |key: &str| -> PyResult<String> {
        map.get_item(key)?
            .ok_or_else(|| value_err(format!("sparklines: '{column}' needs '{key}'")))?
            .extract()
    };
    let from = need("from")?;
    let to = need("to")?;

    let mut sparkline = Sparkline::new();
    if let Some(v) = map.get_item("type")? {
        let name: String = v.extract()?;
        sparkline = sparkline.set_type(match name.as_str() {
            "line" => SparklineType::Line,
            "column" => SparklineType::Column,
            "win_lose" => SparklineType::WinLose,
            other => {
                return Err(value_err(format!(
                    "sparklines: unknown type '{other}' (expected line, column, win_lose)"
                )))
            }
        });
    }
    if let Some(v) = map.get_item("color")? {
        sparkline = sparkline.set_sparkline_color(parse_color(&v.extract::<String>()?)?);
    }
    if let Some(v) = map.get_item("style")? {
        sparkline = sparkline.set_style(v.extract()?);
    }
    for (key, on) in [
        ("high_point", 0),
        ("low_point", 1),
        ("first_point", 2),
        ("last_point", 3),
        ("markers", 4),
        ("negative_points", 5),
        ("axis", 6),
        ("right_to_left", 7),
    ] {
        let Some(v) = map.get_item(key)? else {
            continue;
        };
        let enable: bool = v.extract()?;
        sparkline = match on {
            0 => sparkline.show_high_point(enable),
            1 => sparkline.show_low_point(enable),
            2 => sparkline.show_first_point(enable),
            3 => sparkline.show_last_point(enable),
            4 => sparkline.show_markers(enable),
            5 => sparkline.show_negative_points(enable),
            6 => sparkline.show_axis(enable),
            _ => sparkline.set_right_to_left(enable),
        };
    }

    Ok(Spark {
        column: column.to_string(),
        from,
        to,
        sparkline,
    })
}

impl Sparklines {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(Sparklines(out));
        };
        let map = spec
            .cast::<PyDict>()
            .map_err(|_| value_err("sparklines must be a dict keyed by column name".to_string()))?;

        for (key, value) in map.iter() {
            let column: String = key.extract()?;
            let rule = value.cast::<PyDict>().map_err(|_| {
                value_err(format!(
                    "sparklines: the rule for '{column}' must be a dict"
                ))
            })?;
            out.push(build(rule, &column)?);
        }
        Ok(Sparklines(out))
    }

    /// An unknown column warns and is skipped, like the other column-keyed
    /// options.
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
        // The range is sheet-qualified, and setting it borrows the name while
        // the worksheet itself has to be borrowed mutably to take the group.
        let sheet = worksheet.name().to_string();
        let first = header_row + 1;
        let last = header_row + data_rows;

        for spark in &self.0 {
            let find = |name: &String| headers.iter().position(|h| h == name);
            let (Some(target), Some(from), Some(to)) =
                (find(&spark.column), find(&spark.from), find(&spark.to))
            else {
                warnings.call_method1(
                    "warn",
                    (format!(
                        "sparklines: unknown column in '{}' ('{}'..'{}'), skipped",
                        spark.column, spark.from, spark.to
                    ),),
                )?;
                continue;
            };
            if from > to {
                return Err(value_err(format!(
                    "sparklines: '{}' comes after '{}' in the data",
                    spark.from, spark.to
                )));
            }

            // A group takes the whole block — every data row by the span —
            // and splits it into one chart per row itself. Naming only the
            // first row is rejected as "must be a 2D range".
            let sparkline = spark.sparkline.clone().set_range((
                sheet.as_str(),
                first,
                from as u16,
                last,
                to as u16,
            ));
            worksheet
                .add_sparkline_group(first, target as u16, last, target as u16, &sparkline)
                .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
