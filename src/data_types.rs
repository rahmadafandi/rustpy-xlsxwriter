use indexmap::IndexMap;
use pyo3::conversion::{FromPyObject, IntoPyObjectExt};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDict};
use pyo3::Py;
use pyo3::Borrowed;

#[derive(Debug)]
pub enum WorksheetData {
    Records(Py<PyAny>),   // Holds the list of dicts or iterable
    DataFrame(Py<PyAny>), // Holds the DataFrame object
}

impl<'a, 'py> FromPyObject<'a, 'py> for WorksheetData {
    type Error = PyErr;

    fn extract(ob: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        // Check if it looks like a DataFrame (has "columns" attributes)
        if ob.getattr("columns").is_ok() {
             return Ok(WorksheetData::DataFrame(ob.into_py_any(ob.py())?));
        }

        // Assume it's an iterable of records if not a dataframe
        // We can check if it has __iter__ but almost everything does. 
        // Let's just wrap it. The iteration logic will fail later if it's not iterable.
        return Ok(WorksheetData::Records(ob.into_py_any(ob.py())?));
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
