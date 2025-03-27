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
