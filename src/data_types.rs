use indexmap::IndexMap;
use pyo3::conversion::{FromPyObject, IntoPyObjectExt};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDict};
use pyo3::Borrowed;
use pyo3::Py;

#[derive(Debug)]
pub enum WorksheetData {
    Records(Py<PyAny>),
    ArrowDataFrame(Py<PyAny>),
    PandasDataFrame(Py<PyAny>),
    PolarsDataFrame(Py<PyAny>),
}

impl<'a, 'py> FromPyObject<'a, 'py> for WorksheetData {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        if ob.hasattr("__arrow_c_stream__")? {
            return Ok(WorksheetData::ArrowDataFrame(ob.into_py_any(ob.py())?));
        }
        if ob.hasattr("get_column")? && ob.hasattr("schema")? {
            return Ok(WorksheetData::PolarsDataFrame(ob.into_py_any(ob.py())?));
        }
        if ob.hasattr("columns")? {
            return Ok(WorksheetData::PandasDataFrame(ob.into_py_any(ob.py())?));
        }
        Ok(WorksheetData::Records(ob.into_py_any(ob.py())?))
    }
}

/// Row/col offsets used when calling `set_freeze_panes`.
#[derive(Debug, Clone, Copy, Default)]
pub struct FreezePane {
    pub row: Option<u32>,
    pub col: Option<u16>,
}

/// Freeze-pane configuration. `default` applies to every sheet; `per_sheet`
/// overrides by sheet name. Python wire format remains the existing
/// `Dict[str, Dict[str, int]]` with the sentinel key `"general"`.
#[derive(Debug, Clone, Default)]
pub struct FreezePanesConfig {
    pub default: FreezePane,
    pub per_sheet: IndexMap<String, FreezePane>,
}

impl FreezePanesConfig {
    pub fn resolve(&self, sheet_name: &str) -> FreezePane {
        let mut pane = self.default;
        if let Some(override_pane) = self.per_sheet.get(sheet_name) {
            if override_pane.row.is_some() {
                pane.row = override_pane.row;
            }
            if override_pane.col.is_some() {
                pane.col = override_pane.col;
            }
        }
        pane
    }
}

impl<'a, 'py> FromPyObject<'a, 'py> for FreezePanesConfig {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        let dict = ob.cast::<PyDict>()?;
        let mut config = FreezePanesConfig::default();

        for (key, value) in dict.iter() {
            let key = key.extract::<String>()?;
            let Ok(inner_dict) = value.cast::<PyDict>() else {
                continue;
            };
            let mut pane = FreezePane::default();
            for (inner_key, inner_value) in inner_dict.iter() {
                let inner_key = inner_key.extract::<String>()?;
                let inner_value = inner_value.extract::<usize>()?;
                match inner_key.as_str() {
                    "row" => pane.row = Some(inner_value as u32),
                    "col" => pane.col = Some(inner_value as u16),
                    _ => {}
                }
            }
            if key == "general" {
                config.default = pane;
            } else {
                config.per_sheet.insert(key, pane);
            }
        }

        Ok(config)
    }
}
