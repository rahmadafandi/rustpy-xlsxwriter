use indexmap::IndexMap;
use pyo3::conversion::{FromPyObject, IntoPyObjectExt};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDict};
use pyo3::Py;
use pyo3::Borrowed;

#[derive(Debug)]
pub enum WorksheetData {
    Records(Py<PyAny>),        // Holds the list of dicts or iterable
    ArrowStream(Py<PyAny>),    // Holds an object supporting __arrow_c_stream__ (Pandas/Polars)
    PandasDataFrame(Py<PyAny>), // Fallback for Pandas without Arrow
    PolarsDataFrame(Py<PyAny>), // Fallback for Polars without Arrow
}

impl<'a, 'py> FromPyObject<'a, 'py> for WorksheetData {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        // Prefer Arrow zero-copy path if __arrow_c_stream__ is available
        if ob.getattr("__arrow_c_stream__").is_ok() {
            return Ok(WorksheetData::ArrowStream(ob.into_py_any(ob.py())?));
        }

        // Detect Polars DataFrame: has "get_column" method (Polars-specific)
        if ob.getattr("get_column").is_ok() && ob.getattr("schema").is_ok() {
            return Ok(WorksheetData::PolarsDataFrame(ob.into_py_any(ob.py())?));
        }

        // Detect Pandas DataFrame: has "columns" attribute
        if ob.getattr("columns").is_ok() {
            return Ok(WorksheetData::PandasDataFrame(ob.into_py_any(ob.py())?));
        }

        // Assume it's an iterable of records
        Ok(WorksheetData::Records(ob.into_py_any(ob.py())?))
    }
}


#[derive(Debug, Clone)]
pub struct FreezePaneConfig {
    pub config: IndexMap<String, IndexMap<String, usize>>,
}

impl<'a, 'py> FromPyObject<'a, 'py> for FreezePaneConfig {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        let dict = ob.cast::<PyDict>()?;
        let mut map = IndexMap::new();

        for (key, value) in dict.iter() {
            let key = key.extract::<String>()?;
            if let Some(inner_dict) = value.cast::<PyDict>().ok() {
                let mut inner_map = IndexMap::new();
                for (inner_key, inner_value) in inner_dict.iter() {
                    let inner_key = inner_key.extract::<String>()?;
                    let inner_value = inner_value.extract::<usize>()?;
                    inner_map.insert(inner_key, inner_value);
                }
                map.insert(key, inner_map);
            }
        }

        Ok(FreezePaneConfig { config: map })
    }
}
