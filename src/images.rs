//! Images anchored to cells, parsed from an `images` list.
//!
//! Unlike everything else in a sheet here, an image is not tied to a column of
//! data — it floats above the grid — so it is placed by row and column index
//! rather than by header name, and given as a list rather than a mapping.
//!
//! The source is either a `path` or raw `data`, the second being what a web
//! handler has: a logo already in memory with no file to point at.

use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyDict};
use rust_xlsxwriter::{Image, Worksheet};

use crate::options::{reject_unknown_keys, value_err};
use crate::worksheet::xlsx_err;

const KEYS: [&str; 11] = [
    "path",
    "data",
    "row",
    "col",
    "scale",
    "scale_x",
    "scale_y",
    "fit_to_cell",
    "keep_aspect_ratio",
    "alt_text",
    "url",
];

/// One placed image: where it goes, and how.
pub struct Placed {
    row: u32,
    col: u16,
    image: Image,
    fit_to_cell: bool,
    keep_aspect_ratio: bool,
}

#[derive(Default)]
pub struct Images(Vec<Placed>);

fn build(map: &Bound<'_, PyDict>, index: usize) -> PyResult<Placed> {
    reject_unknown_keys(map, &format!("images[{index}]"), &KEYS)?;

    let mut image = match (map.get_item("path")?, map.get_item("data")?) {
        (Some(_), Some(_)) => {
            return Err(value_err(format!(
                "images[{index}]: give 'path' or 'data', not both"
            )))
        }
        (Some(path), None) => Image::new(path.extract::<std::path::PathBuf>()?),
        (None, Some(data)) => Image::new_from_buffer(&data.extract::<Vec<u8>>()?),
        (None, None) => {
            return Err(value_err(format!(
                "images[{index}]: needs 'path' or 'data'"
            )))
        }
    }
    .map_err(xlsx_err)?;

    // `scale` is the shorthand for the common case of resizing both axes
    // together; the per-axis keys win where they are given.
    if let Some(v) = map.get_item("scale")? {
        let s: f64 = v.extract()?;
        image = image.set_scale_width(s).set_scale_height(s);
    }
    if let Some(v) = map.get_item("scale_x")? {
        image = image.set_scale_width(v.extract()?);
    }
    if let Some(v) = map.get_item("scale_y")? {
        image = image.set_scale_height(v.extract()?);
    }
    if let Some(v) = map.get_item("alt_text")? {
        image = image.set_alt_text(v.extract::<String>()?);
    }
    if let Some(v) = map.get_item("url")? {
        image = image
            .set_url(rust_xlsxwriter::Url::new(v.extract::<String>()?))
            .map_err(xlsx_err)?;
    }

    let fit_to_cell = match map.get_item("fit_to_cell")? {
        Some(v) => v.extract()?,
        None => false,
    };
    let keep_aspect_ratio = match map.get_item("keep_aspect_ratio")? {
        Some(v) => v.extract()?,
        None => true,
    };
    if !fit_to_cell && map.get_item("keep_aspect_ratio")?.is_some() {
        return Err(value_err(format!(
            "images[{index}]: 'keep_aspect_ratio' only applies with 'fit_to_cell'"
        )));
    }

    Ok(Placed {
        row: match map.get_item("row")? {
            Some(v) => v.extract()?,
            None => 0,
        },
        col: match map.get_item("col")? {
            Some(v) => v.extract()?,
            None => 0,
        },
        image,
        fit_to_cell,
        keep_aspect_ratio,
    })
}

impl Images {
    pub fn from_py(spec: Option<&Bound<'_, PyAny>>) -> PyResult<Self> {
        let mut out = Vec::new();
        let Some(spec) = spec else {
            return Ok(Images(out));
        };
        for (index, item) in spec
            .try_iter()
            .map_err(|_| value_err("images must be a list of dicts".to_string()))?
            .enumerate()
        {
            let item = item?;
            let map = item
                .cast::<PyDict>()
                .map_err(|_| value_err(format!("images[{index}]: each image must be a dict")))?;
            out.push(build(map, index)?);
        }
        Ok(Images(out))
    }

    /// Applied before the data, with the rest of the layout.
    ///
    /// An image is a drawing anchored to a cell rather than a cell write, so
    /// constant-memory mode has nothing to revisit and the sheet keeps it.
    pub fn apply(&self, worksheet: &mut Worksheet) -> PyResult<()> {
        for placed in &self.0 {
            if placed.fit_to_cell {
                worksheet.insert_image_fit_to_cell(
                    placed.row,
                    placed.col,
                    &placed.image,
                    placed.keep_aspect_ratio,
                )
            } else {
                worksheet.insert_image(placed.row, placed.col, &placed.image)
            }
            .map_err(xlsx_err)?;
        }
        Ok(())
    }
}
