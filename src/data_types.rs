use indexmap::IndexMap;
use pyo3::conversion::{FromPyObject, IntoPyObject, IntoPyObjectExt};
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
pub enum WorksheetData {
    Records(Vec<WorksheetRow>),
    DataFrame(Py<PyAny>), // Holds the DataFrame object
}

impl<'a, 'py> FromPyObject<'a, 'py> for WorksheetData {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        // Check if it's a list of dicts first (existing logic)
        if let Ok(list) = ob.extract::<Vec<WorksheetRow>>() {
            return Ok(WorksheetData::Records(list));
        }
        
        // Check if it looks like a DataFrame (has "columns" and "values" attributes)
        if ob.getattr("columns").is_ok() && ob.getattr("values").is_ok() {
             return Ok(WorksheetData::DataFrame(ob.into_py_any(ob.py())?));
        }

        Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
            "Argument 'records' must be a list of dictionaries or a pandas DataFrame",
        ))
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
