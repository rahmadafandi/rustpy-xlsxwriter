use indexmap::IndexMap;
use pyo3::conversion::{FromPyObject, IntoPyObject};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDict};
use pyo3::Py;
use pyo3::Borrowed;

#[derive(Debug)]
pub struct WorksheetRow {
    pub hash: IndexMap<String, Option<Py<PyAny>>>,
}

impl<'a, 'py> FromPyObject<'a, 'py> for WorksheetRow {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        let dict = ob.cast::<PyDict>()?;
        let mut map = IndexMap::new();

        for (key, value) in dict.iter() {
            let key = key.extract::<String>()?;
            let value = if value.is_none() {
                None
            } else {
                let obj = value.into_pyobject(dict.py()).clone().unwrap();
                let val = obj.extract::<Py<PyAny>>().unwrap();
                Some(val)
            };
            map.insert(key, value);
        }

        Ok(WorksheetRow { hash: map })
    }
}

#[derive(Debug)]
pub struct WorksheetData {
    pub records: Vec<WorksheetRow>,
}

impl<'a, 'py> FromPyObject<'a, 'py> for WorksheetData {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        let list = ob.extract::<Vec<WorksheetRow>>()?;
        Ok(WorksheetData { records: list })
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
                map.insert(key.clone(), inner_map);
            }
        }

        Ok(FreezePaneConfig { config: map })
    }
}
