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

#[pymethods]
impl Format {
    #[new]
    fn new() -> Self {
        Format {
            inner: XlsxFormat::new(),
        }
    }

    fn set_bold(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_bold();
        slf
    }

    fn set_italic(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_italic();
        slf
    }

    fn set_font_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_font_color(c);
        Ok(slf)
    }

    fn set_background_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_background_color(c);
        Ok(slf)
    }

    fn set_num_format<'p>(mut slf: PyRefMut<'p, Self>, format: &str) -> PyRefMut<'p, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_num_format(format);
        slf
    }

    fn set_align<'p>(
        mut slf: PyRefMut<'p, Self>,
        align: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let a = parse_align(align)?;
        slf.inner = std::mem::take(&mut slf.inner).set_align(a);
        Ok(slf)
    }

    fn set_border<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let b = parse_border(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border(b);
        Ok(slf)
    }

    // ── No-arg bool (infallible) ──────────────────────────────────────────────

    fn set_text_wrap(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_text_wrap();
        slf
    }

    fn set_shrink(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_shrink();
        slf
    }

    fn set_font_strikethrough(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_font_strikethrough();
        slf
    }

    fn set_locked(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_locked();
        slf
    }

    fn set_unlocked(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_unlocked();
        slf
    }

    fn set_hidden(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_hidden();
        slf
    }

    fn set_quote_prefix(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_quote_prefix();
        slf
    }

    fn set_checkbox(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_checkbox();
        slf
    }

    fn set_hyperlink(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_hyperlink();
        slf
    }

    // ── Numeric arg (infallible) ──────────────────────────────────────────────

    fn set_font_size(mut slf: PyRefMut<'_, Self>, size: f64) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_font_size(size);
        slf
    }

    fn set_font_family(mut slf: PyRefMut<'_, Self>, n: u8) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_font_family(n);
        slf
    }

    fn set_font_charset(mut slf: PyRefMut<'_, Self>, n: u8) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_font_charset(n);
        slf
    }

    fn set_rotation(mut slf: PyRefMut<'_, Self>, degrees: i16) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_rotation(degrees);
        slf
    }

    fn set_indent(mut slf: PyRefMut<'_, Self>, n: u8) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_indent(n);
        slf
    }

    fn set_reading_direction(mut slf: PyRefMut<'_, Self>, n: u8) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_reading_direction(n);
        slf
    }

    fn set_num_format_index(mut slf: PyRefMut<'_, Self>, i: u8) -> PyRefMut<'_, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_num_format_index(i);
        slf
    }

    // ── String passthrough (infallible) ──────────────────────────────────────

    fn set_font_name<'p>(mut slf: PyRefMut<'p, Self>, name: &str) -> PyRefMut<'p, Self> {
        slf.inner = std::mem::take(&mut slf.inner).set_font_name(name);
        slf
    }

    fn set_font_scheme<'p>(
        mut slf: PyRefMut<'p, Self>,
        scheme: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let s = parse_font_scheme(scheme)?;
        slf.inner = std::mem::take(&mut slf.inner).set_font_scheme(s);
        Ok(slf)
    }

    // ── Color arg (fallible) ──────────────────────────────────────────────────

    fn set_foreground_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_foreground_color(c);
        Ok(slf)
    }

    fn set_border_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_color(c);
        Ok(slf)
    }

    fn set_border_top_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_top_color(c);
        Ok(slf)
    }

    fn set_border_bottom_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_bottom_color(c);
        Ok(slf)
    }

    fn set_border_left_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_left_color(c);
        Ok(slf)
    }

    fn set_border_right_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_right_color(c);
        Ok(slf)
    }

    fn set_border_diagonal_color<'p>(
        mut slf: PyRefMut<'p, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let c = parse_color(color)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_diagonal_color(c);
        Ok(slf)
    }

    // ── Enum arg (fallible) ───────────────────────────────────────────────────

    #[pyo3(signature = (style = "single"))]
    fn set_underline<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let u = parse_underline(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_underline(u);
        Ok(slf)
    }

    fn set_font_script<'p>(
        mut slf: PyRefMut<'p, Self>,
        script: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let s = parse_script(script)?;
        slf.inner = std::mem::take(&mut slf.inner).set_font_script(s);
        Ok(slf)
    }

    fn set_pattern<'p>(
        mut slf: PyRefMut<'p, Self>,
        pattern: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let p = parse_pattern(pattern)?;
        slf.inner = std::mem::take(&mut slf.inner).set_pattern(p);
        Ok(slf)
    }

    fn set_border_top<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let b = parse_border(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_top(b);
        Ok(slf)
    }

    fn set_border_bottom<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let b = parse_border(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_bottom(b);
        Ok(slf)
    }

    fn set_border_left<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let b = parse_border(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_left(b);
        Ok(slf)
    }

    fn set_border_right<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let b = parse_border(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_right(b);
        Ok(slf)
    }

    fn set_border_diagonal<'p>(
        mut slf: PyRefMut<'p, Self>,
        style: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let b = parse_border(style)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_diagonal(b);
        Ok(slf)
    }

    fn set_border_diagonal_type<'p>(
        mut slf: PyRefMut<'p, Self>,
        dtype: &str,
    ) -> PyResult<PyRefMut<'p, Self>> {
        let d = parse_diagonal_type(dtype)?;
        slf.inner = std::mem::take(&mut slf.inner).set_border_diagonal_type(d);
        Ok(slf)
    }
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
