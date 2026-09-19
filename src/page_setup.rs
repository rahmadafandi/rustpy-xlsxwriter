//! Page and print setup, parsed from a single `page_setup` dict.
//!
//! `rust_xlsxwriter` exposes about twenty separate setters for this. Giving
//! each one its own keyword would take `write_worksheet` from 27 parameters to
//! nearly 50, so they arrive together in one mapping and are validated here.
//!
//! An unknown key raises. Unlike a column name — which comes from data that
//! varies and so only warns — a key comes from the programmer, and a
//! misspelled one silently doing nothing is never what was wanted.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::Worksheet;

use crate::options::value_err;
use crate::worksheet::xlsx_err;

/// Excel's own defaults, in inches, used for any margin left unset.
const DEFAULT_MARGINS: [f64; 6] = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
const MARGIN_KEYS: [&str; 6] = ["left", "right", "top", "bottom", "header", "footer"];

const KEYS: [&str; 15] = [
    "landscape",
    "paper_size",
    "margins",
    "print_area",
    "repeat_rows",
    "repeat_columns",
    "fit_to_pages",
    "scale",
    "center_horizontally",
    "center_vertically",
    "print_gridlines",
    "print_headings",
    "first_page_number",
    "header",
    "footer",
];

#[derive(Default)]
pub struct PageSetup {
    landscape: Option<bool>,
    paper_size: Option<u8>,
    margins: Option<[f64; 6]>,
    print_area: Option<(u32, u16, u32, u16)>,
    repeat_rows: Option<(u32, u32)>,
    repeat_columns: Option<(u16, u16)>,
    fit_to_pages: Option<(u16, u16)>,
    scale: Option<u16>,
    center_horizontally: Option<bool>,
    center_vertically: Option<bool>,
    print_gridlines: Option<bool>,
    print_headings: Option<bool>,
    first_page_number: Option<u16>,
    header: Option<String>,
    footer: Option<String>,
}

/// A span given either as one index (`0`) or as a pair (`(0, 2)`).
fn span<T>(value: &Bound<'_, PyAny>, key: &str) -> PyResult<(T, T)>
where
    T: for<'a, 'p> FromPyObject<'a, 'p> + Copy,
{
    if let Ok(pair) = value.extract::<(T, T)>() {
        return Ok(pair);
    }
    let one: T = value.extract().map_err(|_| {
        value_err(format!(
            "page_setup: '{key}' must be an index or a (first, last) pair"
        ))
    })?;
    Ok((one, one))
}

fn margins(value: &Bound<'_, PyAny>) -> PyResult<[f64; 6]> {
    let map = value.cast::<PyDict>().map_err(|_| {
        value_err(
            "page_setup: 'margins' must be a dict of left/right/top/bottom/header/footer"
                .to_string(),
        )
    })?;
    for key in map.keys().iter() {
        let name: String = key.extract()?;
        if !MARGIN_KEYS.contains(&name.as_str()) {
            return Err(value_err(format!(
                "page_setup: unknown margin '{name}' (expected one of {})",
                MARGIN_KEYS.join(", ")
            )));
        }
    }
    let mut out = DEFAULT_MARGINS;
    for (i, name) in MARGIN_KEYS.iter().enumerate() {
        if let Some(v) = map.get_item(name)? {
            out[i] = v.extract()?;
        }
    }
    Ok(out)
}

impl PageSetup {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut setup = PageSetup::default();
        let Some(spec) = spec else {
            return Ok(setup);
        };
        let map = spec
            .cast::<PyDict>()
            .map_err(|_| value_err("page_setup must be a dict".to_string()))?;

        for (key, value) in map.iter() {
            let name: String = key.extract()?;
            match name.as_str() {
                "landscape" => setup.landscape = Some(value.extract()?),
                "paper_size" => setup.paper_size = Some(value.extract()?),
                "margins" => setup.margins = Some(margins(&value)?),
                "print_area" => {
                    setup.print_area = Some(value.extract().map_err(|_| {
                        value_err(
                            "page_setup: 'print_area' must be \
                         (first_row, first_col, last_row, last_col)"
                                .to_string(),
                        )
                    })?)
                }
                "repeat_rows" => setup.repeat_rows = Some(span(&value, "repeat_rows")?),
                "repeat_columns" => setup.repeat_columns = Some(span(&value, "repeat_columns")?),
                "fit_to_pages" => {
                    setup.fit_to_pages = Some(value.extract().map_err(|_| {
                        value_err(
                            "page_setup: 'fit_to_pages' must be (width, height); \
                         0 lets that dimension run to as many pages as it needs"
                                .to_string(),
                        )
                    })?)
                }
                "scale" => setup.scale = Some(value.extract()?),
                "center_horizontally" => setup.center_horizontally = Some(value.extract()?),
                "center_vertically" => setup.center_vertically = Some(value.extract()?),
                "print_gridlines" => setup.print_gridlines = Some(value.extract()?),
                "print_headings" => setup.print_headings = Some(value.extract()?),
                "first_page_number" => setup.first_page_number = Some(value.extract()?),
                "header" => setup.header = Some(value.extract()?),
                "footer" => setup.footer = Some(value.extract()?),
                _ => {
                    return Err(value_err(format!(
                        "page_setup: unknown key '{name}' (expected one of {})",
                        KEYS.join(", ")
                    )))
                }
            }
        }

        if setup.scale.is_some() && setup.fit_to_pages.is_some() {
            return Err(value_err(
                "page_setup: 'scale' and 'fit_to_pages' are mutually exclusive in Excel — \
                 setting both makes the file open with only one of them applied"
                    .to_string(),
            ));
        }
        Ok(setup)
    }

    /// Applied before the first data row, like the rest of the per-sheet
    /// settings, so constant-memory mode never has a flushed row to revisit.
    pub fn apply(&self, worksheet: &mut Worksheet) -> PyResult<()> {
        if let Some(landscape) = self.landscape {
            if landscape {
                worksheet.set_landscape();
            } else {
                worksheet.set_portrait();
            }
        }
        if let Some(size) = self.paper_size {
            worksheet.set_paper_size(size);
        }
        if let Some(m) = self.margins {
            worksheet.set_margins(m[0], m[1], m[2], m[3], m[4], m[5]);
        }
        if let Some((r1, c1, r2, c2)) = self.print_area {
            worksheet.set_print_area(r1, c1, r2, c2).map_err(xlsx_err)?;
        }
        if let Some((first, last)) = self.repeat_rows {
            worksheet.set_repeat_rows(first, last).map_err(xlsx_err)?;
        }
        if let Some((first, last)) = self.repeat_columns {
            worksheet
                .set_repeat_columns(first, last)
                .map_err(xlsx_err)?;
        }
        if let Some((w, h)) = self.fit_to_pages {
            worksheet.set_print_fit_to_pages(w, h);
        }
        if let Some(scale) = self.scale {
            worksheet.set_print_scale(scale);
        }
        if let Some(on) = self.center_horizontally {
            worksheet.set_print_center_horizontally(on);
        }
        if let Some(on) = self.center_vertically {
            worksheet.set_print_center_vertically(on);
        }
        if let Some(on) = self.print_gridlines {
            worksheet.set_print_gridlines(on);
        }
        if let Some(on) = self.print_headings {
            worksheet.set_print_headings(on);
        }
        if let Some(n) = self.first_page_number {
            worksheet.set_print_first_page_number(n);
        }
        if let Some(text) = &self.header {
            worksheet.set_header(text);
        }
        if let Some(text) = &self.footer {
            worksheet.set_footer(text);
        }
        Ok(())
    }
}
