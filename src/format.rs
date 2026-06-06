use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use rust_xlsxwriter::{
    Color, Format as XlsxFormat, FontScheme, FormatAlign, FormatBorder, FormatDiagonalBorder,
    FormatPattern, FormatScript, FormatUnderline,
};

/// Parse a color string: `#RRGGBB` / `RRGGBB` hex, or a common color name.
pub fn parse_color(s: &str) -> PyResult<Color> {
    let t = s.trim();
    let named = match t.to_ascii_lowercase().as_str() {
        "black" => Some(Color::Black),
        "blue" => Some(Color::Blue),
        "brown" => Some(Color::Brown),
        "cyan" => Some(Color::Cyan),
        "gray" | "grey" => Some(Color::Gray),
        "green" => Some(Color::Green),
        "lime" => Some(Color::Lime),
        "magenta" => Some(Color::Magenta),
        "navy" => Some(Color::Navy),
        "orange" => Some(Color::Orange),
        "pink" => Some(Color::Pink),
        "purple" => Some(Color::Purple),
        "red" => Some(Color::Red),
        "silver" => Some(Color::Silver),
        "white" => Some(Color::White),
        "yellow" => Some(Color::Yellow),
        "automatic" => Some(Color::Automatic),
        _ => None,
    };
    if let Some(c) = named {
        return Ok(c);
    }
    let hex = t.strip_prefix('#').unwrap_or(t);
    if hex.len() == 6 {
        if let Ok(rgb) = u32::from_str_radix(hex, 16) {
            return Ok(Color::RGB(rgb));
        }
    }
    Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
        "invalid color '{s}' (expected '#RRGGBB', 'RRGGBB', or a color name)"
    )))
}

fn parse_align(s: &str) -> PyResult<FormatAlign> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "general" => FormatAlign::General,
        "left" => FormatAlign::Left,
        "center" => FormatAlign::Center,
        "right" => FormatAlign::Right,
        "fill" => FormatAlign::Fill,
        "justify" => FormatAlign::Justify,
        "center_across" => FormatAlign::CenterAcross,
        "distributed" => FormatAlign::Distributed,
        "top" => FormatAlign::Top,
        "bottom" => FormatAlign::Bottom,
        "vcenter" | "vertical_center" => FormatAlign::VerticalCenter,
        "vjustify" | "vertical_justify" => FormatAlign::VerticalJustify,
        "vdistributed" | "vertical_distributed" => FormatAlign::VerticalDistributed,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid align '{other}' (valid: general, left, center, right, fill, \
                 justify, center_across, distributed, top, bottom, vcenter, vjustify, \
                 vdistributed; also: vertical_center, vertical_justify, vertical_distributed)"
            )))
        }
    })
}

fn parse_border(s: &str) -> PyResult<FormatBorder> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "none" => FormatBorder::None,
        "thin" => FormatBorder::Thin,
        "medium" => FormatBorder::Medium,
        "dashed" => FormatBorder::Dashed,
        "dotted" => FormatBorder::Dotted,
        "thick" => FormatBorder::Thick,
        "double" => FormatBorder::Double,
        "hair" => FormatBorder::Hair,
        "medium_dashed" => FormatBorder::MediumDashed,
        "dash_dot" => FormatBorder::DashDot,
        "medium_dash_dot" => FormatBorder::MediumDashDot,
        "dash_dot_dot" => FormatBorder::DashDotDot,
        "medium_dash_dot_dot" => FormatBorder::MediumDashDotDot,
        "slant_dash_dot" => FormatBorder::SlantDashDot,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid border '{other}' (valid: none, thin, medium, dashed, dotted, \
                 thick, double, hair, medium_dashed, dash_dot, medium_dash_dot, \
                 dash_dot_dot, medium_dash_dot_dot, slant_dash_dot)"
            )))
        }
    })
}

fn parse_pattern(s: &str) -> PyResult<FormatPattern> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "none" => FormatPattern::None,
        "solid" => FormatPattern::Solid,
        "medium_gray" => FormatPattern::MediumGray,
        "dark_gray" => FormatPattern::DarkGray,
        "light_gray" => FormatPattern::LightGray,
        "dark_horizontal" => FormatPattern::DarkHorizontal,
        "dark_vertical" => FormatPattern::DarkVertical,
        "dark_down" => FormatPattern::DarkDown,
        "dark_up" => FormatPattern::DarkUp,
        "dark_grid" => FormatPattern::DarkGrid,
        "dark_trellis" => FormatPattern::DarkTrellis,
        "light_horizontal" => FormatPattern::LightHorizontal,
        "light_vertical" => FormatPattern::LightVertical,
        "light_down" => FormatPattern::LightDown,
        "light_up" => FormatPattern::LightUp,
        "light_grid" => FormatPattern::LightGrid,
        "light_trellis" => FormatPattern::LightTrellis,
        "gray125" => FormatPattern::Gray125,
        "gray0625" => FormatPattern::Gray0625,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid pattern '{other}' (valid: none, solid, medium_gray, dark_gray, \
                 light_gray, dark_horizontal, dark_vertical, dark_down, dark_up, dark_grid, \
                 dark_trellis, light_horizontal, light_vertical, light_down, light_up, \
                 light_grid, light_trellis, gray125, gray0625)"
            )))
        }
    })
}

fn parse_underline(s: &str) -> PyResult<FormatUnderline> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "none" => FormatUnderline::None,
        "single" => FormatUnderline::Single,
        "double" => FormatUnderline::Double,
        "single_accounting" => FormatUnderline::SingleAccounting,
        "double_accounting" => FormatUnderline::DoubleAccounting,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid underline '{other}' (valid: none, single, double, \
                 single_accounting, double_accounting)"
            )))
        }
    })
}

fn parse_script(s: &str) -> PyResult<FormatScript> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "none" => FormatScript::None,
        "super" | "superscript" => FormatScript::Superscript,
        "sub" | "subscript" => FormatScript::Subscript,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid script '{other}' (valid: none, super, superscript, sub, subscript)"
            )))
        }
    })
}

fn parse_diagonal_type(s: &str) -> PyResult<FormatDiagonalBorder> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "none" => FormatDiagonalBorder::None,
        "border_up" => FormatDiagonalBorder::BorderUp,
        "border_down" => FormatDiagonalBorder::BorderDown,
        "border_up_down" => FormatDiagonalBorder::BorderUpDown,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid diagonal_type '{other}' (valid: none, border_up, border_down, border_up_down)"
            )))
        }
    })
}

fn parse_font_scheme(s: &str) -> PyResult<FontScheme> {
    Ok(match s.to_ascii_lowercase().as_str() {
        "none" => FontScheme::None,
        "body" => FontScheme::Body,
        "headings" | "heading" => FontScheme::Headings,
        other => {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "invalid font_scheme '{other}' (valid: none, body, headings)"
            )))
        }
    })
}

/// Python-facing cell format. Chainable; each setter returns `self`.
#[pyclass(from_py_object)]
#[derive(Clone)]
pub struct Format {
    pub inner: XlsxFormat,
}

/// Generate the chainable `#[pymethods]` for [`Format`]. Every setter has one
/// of four identical shapes, so they are produced from compact lists instead of
/// ~40 hand-written near-duplicate methods. (The whole `#[pymethods] impl` is
/// emitted by this macro because the pyo3 attribute macro must see literal
/// method items — it cannot see through an inner declarative macro.)
///
/// - `flags`: no-arg toggles → `self.inner.set_x()`
/// - `values`: one primitive arg passed straight through
/// - `strs`: one `&str` arg passed straight through
/// - `parsed`: one `&str` arg validated through a `parse_*` fn first
macro_rules! format_methods {
    (
        flags: [$($flag:ident),* $(,)?],
        values: [$(($vname:ident, $vty:ty)),* $(,)?],
        strs: [$($sname:ident),* $(,)?],
        parsed: [$(($pname:ident, $pparse:ident)),* $(,)?],
    ) => {
        #[pymethods]
        impl Format {
            #[new]
            fn new() -> Self {
                Format { inner: XlsxFormat::new() }
            }

            $(
                fn $flag(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
                    slf.inner = std::mem::take(&mut slf.inner).$flag();
                    slf
                }
            )*

            $(
                fn $vname(mut slf: PyRefMut<'_, Self>, value: $vty) -> PyRefMut<'_, Self> {
                    slf.inner = std::mem::take(&mut slf.inner).$vname(value);
                    slf
                }
            )*

            $(
                fn $sname<'p>(mut slf: PyRefMut<'p, Self>, value: &str) -> PyRefMut<'p, Self> {
                    slf.inner = std::mem::take(&mut slf.inner).$sname(value);
                    slf
                }
            )*

            $(
                fn $pname<'p>(
                    mut slf: PyRefMut<'p, Self>,
                    value: &str,
                ) -> PyResult<PyRefMut<'p, Self>> {
                    let parsed = $pparse(value)?;
                    slf.inner = std::mem::take(&mut slf.inner).$pname(parsed);
                    Ok(slf)
                }
            )*

            // Only setter with a defaulted argument, kept explicit.
            #[pyo3(signature = (style = "single"))]
            fn set_underline<'p>(
                mut slf: PyRefMut<'p, Self>,
                style: &str,
            ) -> PyResult<PyRefMut<'p, Self>> {
                let u = parse_underline(style)?;
                slf.inner = std::mem::take(&mut slf.inner).set_underline(u);
                Ok(slf)
            }
        }
    };
}

format_methods! {
    flags: [
        set_bold,
        set_italic,
        set_text_wrap,
        set_shrink,
        set_font_strikethrough,
        set_locked,
        set_unlocked,
        set_hidden,
        set_quote_prefix,
        set_checkbox,
        set_hyperlink,
    ],
    values: [
        (set_font_size, f64),
        (set_font_family, u8),
        (set_font_charset, u8),
        (set_rotation, i16),
        (set_indent, u8),
        (set_reading_direction, u8),
        (set_num_format_index, u8),
    ],
    strs: [
        set_num_format,
        set_font_name,
    ],
    parsed: [
        (set_font_color, parse_color),
        (set_background_color, parse_color),
        (set_foreground_color, parse_color),
        (set_border_color, parse_color),
        (set_border_top_color, parse_color),
        (set_border_bottom_color, parse_color),
        (set_border_left_color, parse_color),
        (set_border_right_color, parse_color),
        (set_border_diagonal_color, parse_color),
        (set_align, parse_align),
        (set_border, parse_border),
        (set_border_top, parse_border),
        (set_border_bottom, parse_border),
        (set_border_left, parse_border),
        (set_border_right, parse_border),
        (set_border_diagonal, parse_border),
        (set_pattern, parse_pattern),
        (set_font_scheme, parse_font_scheme),
        (set_font_script, parse_script),
        (set_border_diagonal_type, parse_diagonal_type),
    ],
}

/// Borrow the inner `rust_xlsxwriter::Format` for column `idx`, if one was
/// resolved. Used per-cell to let an explicit column format win over the
/// sheet-wide float/datetime format.
pub fn col_override(col_formats: &[Option<Format>], idx: usize) -> Option<&XlsxFormat> {
    col_formats.get(idx).and_then(|o| o.as_ref()).map(|f| &f.inner)
}

/// Apply resolved per-column formats to the worksheet via `set_column_format`.
/// Must be called BEFORE data rows are written (constant-memory mode).
pub fn apply_column_formats(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    col_formats: &[Option<Format>],
) -> PyResult<()> {
    for (col, fmt) in col_formats.iter().enumerate() {
        if let Some(f) = fmt {
            worksheet
                .set_column_format(col as u16, &f.inner)
                .map_err(crate::worksheet::xlsx_err)?;
        }
    }
    Ok(())
}

/// Resolve `column_formats` (dict by header name or positional list of Format)
/// into a per-column `Vec<Option<Format>>` aligned to `headers`. Unknown dict
/// names warn and are skipped. A non-Format value raises `TypeError`.
pub fn resolve_column_formats(
    spec: Option<&Bound<'_, PyAny>>,
    headers: &[String],
    py: Python,
) -> PyResult<Vec<Option<Format>>> {
    let mut out: Vec<Option<Format>> = vec![None; headers.len()];
    let Some(spec) = spec else { return Ok(out) };

    let warnings = py.import("warnings")?;

    if let Ok(dict) = spec.cast::<PyDict>() {
        for (key, val) in dict.iter() {
            let name: String = key.extract()?;
            let fmt: Format = val.extract().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "column_formats values must be Format objects",
                )
            })?;
            match headers.iter().position(|h| h == &name) {
                Some(idx) => out[idx] = Some(fmt),
                None => {
                    warnings.call_method1(
                        "warn",
                        (format!("column_formats: unknown column '{name}', skipped"),),
                    )?;
                }
            }
        }
    } else if let Ok(list) = spec.cast::<PyList>() {
        for (idx, item) in list.iter().enumerate() {
            if idx >= out.len() {
                warnings.call_method1(
                    "warn",
                    (format!(
                        "column_formats: index {idx} out of range ({} columns), skipped",
                        out.len()
                    ),),
                )?;
                continue;
            }
            if item.is_none() {
                continue;
            }
            let fmt: Format = item.extract().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "column_formats values must be Format objects or None",
                )
            })?;
            out[idx] = Some(fmt);
        }
    } else {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "column_formats must be a dict (by column name) or a list (positional)",
        ));
    }
    Ok(out)
}
