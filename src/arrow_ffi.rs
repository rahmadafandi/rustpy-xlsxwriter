//! Arrow C Stream Interface bridge — converts Python objects with
//! `__arrow_c_stream__` into Rust `arrow_array::RecordBatchReader`
//! without the `pyo3-arrow` crate (avoids chrono-tz compile issues).

use arrow_array::ffi_stream::{ArrowArrayStreamReader, FFI_ArrowArrayStream};
use arrow_array::RecordBatchReader;
use arrow_schema::ArrowError;
use pyo3::prelude::*;
use pyo3::Py;

/// Call `obj.__arrow_c_stream__()` and consume the returned PyCapsule
/// to produce a `Box<dyn RecordBatchReader + Send>`.
pub fn stream_to_reader(
    obj: &Py<PyAny>,
    py: Python<'_>,
) -> PyResult<Box<dyn RecordBatchReader + Send>> {
    // 1. Call __arrow_c_stream__(None) → PyCapsule
    let capsule = obj.call_method1(py, "__arrow_c_stream__", (py.None(),))?;
    let capsule_bound = capsule.bind(py);

    // 2. Extract the raw pointer from the PyCapsule
    let ptr = unsafe {
        let cap_ptr = pyo3::ffi::PyCapsule_GetPointer(
            capsule_bound.as_ptr(),
            b"arrow_array_stream\0".as_ptr() as *const _,
        );
        if cap_ptr.is_null() {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Failed to get pointer from Arrow PyCapsule",
            ));
        }
        cap_ptr as *mut FFI_ArrowArrayStream
    };

    // 3. Convert FFI stream to Rust RecordBatchReader
    let stream =
        unsafe { ArrowArrayStreamReader::from_raw(ptr) }.map_err(|e: ArrowError| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                "Failed to create Arrow reader from stream: {}",
                e
            ))
        })?;

    Ok(Box::new(stream))
}
