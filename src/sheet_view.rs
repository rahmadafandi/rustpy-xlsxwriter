//! How a sheet presents on screen, parsed from a `sheet_view` mapping.
//!
//! Kept apart from [`crate::page_setup`] on purpose: that module is about
//! paper, this one is about the window. Sharing one mapping would put `zoom`
//! next to `margins`, which reads like a mistake.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{Color, Worksheet};

use crate::format::parse_color;

const KEYS: [&str; 6] = [
    "tab_color",
    "gridlines",
    "zoom",
    "right_to_left",
    "hidden",
    "selected",
];

#[derive(Default)]
pub struct SheetView {
    tab_color: Option<Color>,
    gridlines: Option<bool>,
    zoom: Option<u16>,
    right_to_left: Option<bool>,
    hidden: Option<bool>,
    selected: Option<bool>,
}

fn value_err(msg: String) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(msg)
}

impl SheetView {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut view = SheetView::default();
        let Some(spec) = spec else {
            return Ok(view);
        };
        let map = spec
            .cast::<PyDict>()
            .map_err(|_| value_err("sheet_view must be a dict".to_string()))?;

        for (key, value) in map.iter() {
            let name: String = key.extract()?;
            match name.as_str() {
                "tab_color" => view.tab_color = Some(parse_color(&value.extract::<String>()?)?),
                "gridlines" => view.gridlines = Some(value.extract()?),
                "zoom" => view.zoom = Some(value.extract()?),
                "right_to_left" => view.right_to_left = Some(value.extract()?),
                "hidden" => view.hidden = Some(value.extract()?),
                "selected" => view.selected = Some(value.extract()?),
                _ => {
                    // A key comes from the programmer, so a misspelled one that
                    // quietly does nothing gives no other signal — same call as
                    // page_setup.
                    return Err(value_err(format!(
                        "sheet_view: unknown key '{name}' (expected one of {})",
                        KEYS.join(", ")
                    )));
                }
            }
        }

        if view.hidden == Some(true) && view.selected == Some(true) {
            return Err(value_err(
                "sheet_view: a sheet cannot be both 'hidden' and 'selected' — \
                 Excel rejects a workbook whose active sheet is hidden"
                    .to_string(),
            ));
        }
        Ok(view)
    }

    pub fn apply(&self, worksheet: &mut Worksheet) {
        if let Some(color) = self.tab_color {
            worksheet.set_tab_color(color);
        }
        if let Some(on) = self.gridlines {
            worksheet.set_screen_gridlines(on);
        }
        if let Some(zoom) = self.zoom {
            worksheet.set_zoom(zoom);
        }
        if let Some(on) = self.right_to_left {
            worksheet.set_right_to_left(on);
        }
        if let Some(on) = self.hidden {
            worksheet.set_hidden(on);
        }
        if let Some(on) = self.selected {
            worksheet.set_selected(on);
        }
    }
}
