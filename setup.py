from setuptools import setup
from setuptools_rust import RustExtension

setup(
    name="rustpy_xlsxwriter",
    version="0.0.6",
    description="Rust Python bindings for rust_xlsxwriter",
    rust_extensions=[
        RustExtension(
            "rustpy_xlsxwriter.rustpy_xlsxwriter", "Cargo.toml", binding="pyo3"
        )
    ],
    packages=["rustpy_xlsxwriter"],
    zip_safe=False,
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Rust",
        "Operating System :: OS Independent",
    ],
)
