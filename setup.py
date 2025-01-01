from setuptools import setup
from setuptools_rust import RustExtension

setup(
    name="rustpy_xlsxwriter",
    version="0.1.0",
    description="A Rust-powered Excel library for Python",
    rust_extensions=[RustExtension("rustpy_xlsxwriter.rustpy_xlsxwriter", "Cargo.toml", binding="pyo3")],
    packages=["rustpy_xlsxwriter"],
    zip_safe=False,
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Rust",
        "Operating System :: OS Independent",
    ],
)
