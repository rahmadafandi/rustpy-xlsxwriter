use indexmap::IndexMap;
use pyo3::conversion::{FromPyObject, IntoPyObject};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDict};

#[derive(Debug)]
pub struct WorksheetRow {
    pub hash: IndexMap<String, Option<PyObject>>,
}

impl<'source> FromPyObject<'source> for WorksheetRow {
    fn extract_bound(ob: &Bound<'source, PyAny>) -> PyResult<Self> {
        let dict = ob.downcast::<PyDict>()?;
        let mut map = IndexMap::new();

        for (key, value) in dict.iter() {
            let key = key.extract::<String>()?;
            let value = if value.is_none() {
                None
            } else {
                let obj = value.into_pyobject(dict.py()).clone().unwrap();
                let val = obj.extract::<PyObject>().unwrap();
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

impl<'source> FromPyObject<'source> for WorksheetData {
    fn extract_bound(ob: &Bound<'source, PyAny>) -> PyResult<Self> {
        let list = ob.extract::<Vec<WorksheetRow>>()?;
        Ok(WorksheetData { records: list })
    }
}

#[derive(Debug, Clone)]
pub struct FreezePaneConfig {
    pub config: IndexMap<String, IndexMap<String, usize>>,
}

impl<'source> FromPyObject<'source> for FreezePaneConfig {
    fn extract_bound(ob: &Bound<'source, PyAny>) -> PyResult<Self> {
        let dict = ob.downcast::<PyDict>()?;
        let mut map = IndexMap::new();

        for (key, value) in dict.iter() {
            let key = key.extract::<String>()?;
            if let Some(inner_dict) = value.downcast::<PyDict>().ok() {
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
