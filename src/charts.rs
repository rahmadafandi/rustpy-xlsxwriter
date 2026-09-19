//! Charts anchored to a cell, parsed from a `charts` list.
//!
//! Series are named by column, like everything else keyed by column here, and
//! cover that column's data rows — so the ranges are only known once the rows
//! have been written, and these are applied after the data.
//!
//! ```python
//! charts=[{"type": "column", "series": ["q1", "q2"], "categories": "region"}]
//! ```
//!
//! By default the chart lands just to the right of the data rather than on top
//! of it, which is the only placement that is right more often than not.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{Chart, ChartType, Worksheet};

use crate::options::{opt, reject_unknown_keys, value_err};
use crate::worksheet::xlsx_err;

const KEYS: [&str; 12] = [
    "type",
    "series",
    "categories",
    "row",
    "col",
    "title",
    "x_axis",
    "y_axis",
    "width",
    "height",
    "style",
    "legend",
];

const TYPES: [&str; 20] = [
    "area",
    "area_stacked",
    "area_percent_stacked",
    "bar",
    "bar_stacked",
    "bar_percent_stacked",
    "column",
    "column_stacked",
    "column_percent_stacked",
    "doughnut",
    "line",
    "line_stacked",
    "line_percent_stacked",
    "pie",
    "radar",
    "radar_with_markers",
    "radar_filled",
    "scatter",
    "scatter_smooth",
    "stock",
];

fn chart_type(name: &str) -> PyResult<ChartType> {
    Ok(match name {
        "area" => ChartType::Area,
        "area_stacked" => ChartType::AreaStacked,
        "area_percent_stacked" => ChartType::AreaPercentStacked,
        "bar" => ChartType::Bar,
        "bar_stacked" => ChartType::BarStacked,
        "bar_percent_stacked" => ChartType::BarPercentStacked,
        "column" => ChartType::Column,
        "column_stacked" => ChartType::ColumnStacked,
        "column_percent_stacked" => ChartType::ColumnPercentStacked,
        "doughnut" => ChartType::Doughnut,
        "line" => ChartType::Line,
        "line_stacked" => ChartType::LineStacked,
        "line_percent_stacked" => ChartType::LinePercentStacked,
        "pie" => ChartType::Pie,
        "radar" => ChartType::Radar,
        "radar_with_markers" => ChartType::RadarWithMarkers,
        "radar_filled" => ChartType::RadarFilled,
        "scatter" => ChartType::Scatter,
        "scatter_smooth" => ChartType::ScatterSmooth,
        "stock" => ChartType::Stock,
        other => {
            return Err(value_err(format!(
                "charts: unknown type '{other}' (expected one of {})",
                TYPES.join(", ")
            )))
        }
    })
}

/// One series: the column it reads, and the name to show for it.
struct Series {
    column: String,
    /// `None` links the series name to the column's header cell, so Excel
    /// shows whatever the header says.
    name: Option<String>,
}

/// One chart, still unbuilt: the ranges need the row count.
struct Spec {
    kind: ChartType,
    series: Vec<Series>,
    categories: Option<String>,
    row: Option<u32>,
    col: Option<u16>,
    title: Option<String>,
    x_axis: Option<String>,
    y_axis: Option<String>,
    width: Option<u32>,
    height: Option<u32>,
    style: Option<u8>,
    legend: Option<bool>,
}

#[derive(Default)]
pub struct Charts(Vec<Spec>);

fn parse_series(value: &Bound<'_, PyAny>, index: usize) -> PyResult<Vec<Series>> {
    let mut out = Vec::new();
    for item in value.try_iter().map_err(|_| {
        value_err(format!(
            "charts[{index}]: 'series' must be a list of column names or dicts"
        ))
    })? {
        let item = item?;
        // A bare column name is the short form; the dict form is only needed
        // to override the label.
        if let Ok(column) = item.extract::<String>() {
            out.push(Series { column, name: None });
            continue;
        }
        let map = item.cast::<PyDict>().map_err(|_| {
            value_err(format!(
                "charts[{index}]: each series must be a column name or a dict with 'values'"
            ))
        })?;
        let column: String = map
            .get_item("values")?
            .ok_or_else(|| value_err(format!("charts[{index}]: a series dict needs 'values'")))?
            .extract()?;
        let name = match map.get_item("name")? {
            Some(v) => Some(v.extract::<String>()?),
            None => None,
        };
        out.push(Series { column, name });
    }
    if out.is_empty() {
        return Err(value_err(format!(
            "charts[{index}]: 'series' is empty, so the chart would have nothing to draw"
        )));
    }
    Ok(out)
}

fn parse(map: &Bound<'_, PyDict>, index: usize) -> PyResult<Spec> {
    reject_unknown_keys(map, &format!("charts[{index}]"), &KEYS)?;
    let get = |key: &str| map.get_item(key);
    let text = |key: &str| -> PyResult<Option<String>> {
        match map.get_item(key)? {
            Some(v) => Ok(Some(v.extract()?)),
            None => Ok(None),
        }
    };

    let kind = chart_type(
        &get("type")?
            .ok_or_else(|| value_err(format!("charts[{index}]: needs 'type'")))?
            .extract::<String>()?,
    )?;
    let series = parse_series(
        &get("series")?.ok_or_else(|| value_err(format!("charts[{index}]: needs 'series'")))?,
        index,
    )?;

    let categories = text("categories")?;
    // A scatter chart plots x against y, so its categories are the x values
    // rather than labels. The crate rejects a missing one too, but only once
    // every row has been written — this says so before any work is done, and
    // names the key the caller would have to add.
    if categories.is_none() && matches!(kind, ChartType::Scatter | ChartType::ScatterSmooth) {
        return Err(value_err(format!(
            "charts[{index}]: a scatter chart needs 'categories' — they are its \
             x values, not labels"
        )));
    }

    Ok(Spec {
        kind,
        series,
        categories,
        row: opt(map, "row")?,
        col: opt(map, "col")?,
        title: text("title")?,
        x_axis: text("x_axis")?,
        y_axis: text("y_axis")?,
        width: opt(map, "width")?,
        height: opt(map, "height")?,
        style: opt(map, "style")?,
        legend: opt(map, "legend")?,
    })
}

impl Charts {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(Charts(out));
        };
        for (index, item) in spec
            .try_iter()
            .map_err(|_| value_err("charts must be a list of dicts".to_string()))?
            .enumerate()
        {
            let item = item?;
            let map = item
                .cast::<PyDict>()
                .map_err(|_| value_err(format!("charts[{index}]: each chart must be a dict")))?;
            out.push(parse(map, index)?);
        }
        Ok(Charts(out))
    }

    /// An unknown column warns and the chart is skipped: a chart with a
    /// missing series would draw a misleading picture, which is worse than no
    /// chart, but still not worth failing the export over.
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
        let sheet = worksheet.name().to_string();
        let first = header_row + 1;
        let last = header_row + data_rows;

        for (index, spec) in self.0.iter().enumerate() {
            let find = |name: &String| headers.iter().position(|h| h == name);

            let mut columns = Vec::with_capacity(spec.series.len());
            let mut missing = None;
            for series in &spec.series {
                match find(&series.column) {
                    Some(idx) => columns.push((idx as u16, series.name.as_deref())),
                    None => {
                        missing = Some(series.column.clone());
                        break;
                    }
                }
            }
            let category_col = match &spec.categories {
                Some(name) => match find(name) {
                    Some(idx) => Some(idx as u16),
                    None => {
                        missing = missing.or_else(|| Some(name.clone()));
                        None
                    }
                },
                None => None,
            };
            if let Some(name) = missing {
                warnings.call_method1(
                    "warn",
                    (format!(
                        "charts[{index}]: unknown column '{name}', chart skipped"
                    ),),
                )?;
                continue;
            }

            let mut chart = Chart::new(spec.kind);
            for (col, name) in columns {
                let series = chart.add_series();
                series.set_values((sheet.as_str(), first, col, last, col));
                match name {
                    Some(text) => series.set_name(text),
                    // Linked to the header cell rather than copied, so the
                    // legend follows the header if it is ever edited.
                    None => series.set_name((sheet.as_str(), header_row, col, header_row, col)),
                };
                if let Some(cat) = category_col {
                    series.set_categories((sheet.as_str(), first, cat, last, cat));
                }
            }
            if let Some(title) = &spec.title {
                chart.title().set_name(title.as_str());
            }
            if let Some(name) = &spec.x_axis {
                chart.x_axis().set_name(name.as_str());
            }
            if let Some(name) = &spec.y_axis {
                chart.y_axis().set_name(name.as_str());
            }
            if let Some(w) = spec.width {
                chart.set_width(w);
            }
            if let Some(h) = spec.height {
                chart.set_height(h);
            }
            if let Some(style) = spec.style {
                chart.set_style(style);
            }
            if spec.legend == Some(false) {
                chart.legend().set_hidden();
            }

            // Default placement: one column clear of the data, level with the
            // header, so a chart never lands on top of the table.
            let row = spec.row.unwrap_or(header_row);
            let col = spec.col.unwrap_or_else(|| headers.len() as u16 + 1);
            worksheet.insert_chart(row, col, &chart).map_err(xlsx_err)?;
        }
        Ok(())
    }
}
