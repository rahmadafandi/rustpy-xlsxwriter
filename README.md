# RustPy-XlsxWriter

[![PyPI version](https://badge.fury.io/py/rustpy-xlsxwriter.svg)](https://badge.fury.io/py/rustpy-xlsxwriter)
[![Python Versions](https://img.shields.io/pypi/pyversions/rustpy-xlsxwriter.svg)](https://pypi.org/project/rustpy-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://pepy.tech/badge/rustpy-xlsxwriter)](https://pepy.tech/project/rustpy-xlsxwriter)
[![CI](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions/workflows/CI.yml/badge.svg)](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions/workflows/CI.yml)

RustPy-XlsxWriter is a high-performance library for generating Excel files in Python, powered by Rust and integrated using PyO3. This library is ideal for creating Excel files with large datasets efficiently while maintaining a simple and Pythonic interface.

## Installation

```bash
pip install rustpy-xlsxwriter
```

## Quick Start

```python
from rustpy_xlsxwriter import FastExcel

# One-liner
FastExcel("output.xlsx").sheet("Sheet1", records).save()

# Multiple sheets with options
(
    FastExcel("report.xlsx", password="secret")
    .format(float_format="0.00", index_columns=["Name"])
    .freeze(row=1, col=1)
    .sheet("Users", user_records)
    .sheet("Orders", order_records)
    .save()
)
```

## Features

- **7x–7.8x faster** than Python's xlsxwriter
- Fluent builder API via `FastExcel` class
- Support for `str`, `int`, `float`, `bool`, `None`, `datetime` values
- Multiple sheets in a single file
- Password protection
- Freeze panes (rows & columns)
- Float formatting & bold index columns
- Pandas DataFrame support
- `io.BytesIO` in-memory buffer support
- Generator/iterator streaming for memory-efficient large datasets

## Usage Examples

### Single Sheet

```python
from rustpy_xlsxwriter import FastExcel
from datetime import datetime

records = [
    {"Name": "Alice", "Age": 30, "City": "New York", "Active": True, "Join Date": datetime(2023, 1, 15)},
    {"Name": "Bob", "Age": 25, "City": "San Francisco", "Active": False, "Join Date": datetime(2023, 2, 1)},
]

FastExcel("output.xlsx").sheet("Employees", records).save()
```

### Multiple Sheets

```python
from rustpy_xlsxwriter import FastExcel

employees = [{"Name": "Alice", "Age": 30}, {"Name": "Bob", "Age": 25}]
inventory = [{"Product": "Laptop", "Price": 1000.0}, {"Product": "Phone", "Price": 500.0}]

(
    FastExcel("report.xlsx")
    .sheet("Employees", employees)
    .sheet("Inventory", inventory)
    .save()
)
```

### Freeze Panes

```python
from rustpy_xlsxwriter import FastExcel

# Freeze first row on all sheets
(
    FastExcel("frozen.xlsx")
    .freeze(row=1)
    .sheet("Sheet1", records)
    .sheet("Sheet2", more_records)
    .save()
)

# Per-sheet freeze pane configuration
(
    FastExcel("custom_freeze.xlsx")
    .freeze(row=1, col=0)                  # general (all sheets)
    .freeze(row=1, col=2, sheet="Sheet1")  # override for Sheet1
    .freeze(row=2, col=1, sheet="Sheet2")  # override for Sheet2
    .sheet("Sheet1", data1)
    .sheet("Sheet2", data2)
    .save()
)
```

### Pandas DataFrame

```python
import pandas as pd
from rustpy_xlsxwriter import FastExcel

df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25], "Score": [88.5, 92.3]})

# Basic
FastExcel("dataframe.xlsx").sheet("Data", df).save()

# With styling
(
    FastExcel("styled.xlsx")
    .format(float_format="0.00", index_columns=["Name"])
    .sheet("Data", df)
    .save()
)
```

### In-Memory Buffer

```python
import io
from rustpy_xlsxwriter import FastExcel

buf = io.BytesIO()
FastExcel(buf).sheet("Sheet1", [{"Name": "Alice", "Age": 30}]).save()

xlsx_data = buf.getvalue()
```

### Generator Streaming (Memory-Efficient)

```python
from rustpy_xlsxwriter import FastExcel

def rows():
    for i in range(1_000_000):
        yield {"id": i, "value": f"row_{i}"}

FastExcel("streamed.xlsx").sheet("Data", rows()).save()
```

### Functional API

The lower-level functional API is also available:

```python
from rustpy_xlsxwriter import write_worksheet, write_worksheets

# Single sheet
write_worksheet(records, "output.xlsx", sheet_name="Sheet1", password="secret")

# Multiple sheets
write_worksheets(
    [{"Sheet1": records1}, {"Sheet2": records2}],
    "output.xlsx",
    freeze_panes={"general": {"row": 1, "col": 0}},
)
```

## API Reference

### `FastExcel` Class

| Method | Description |
|---|---|
| `FastExcel(target, *, password=None)` | Create writer for file path or `BytesIO` buffer |
| `.format(*, float_format=None, index_columns=None)` | Set number format & bold index columns |
| `.freeze(*, row=None, col=None, sheet=None)` | Configure freeze panes (general or per-sheet) |
| `.sheet(name, data)` | Add a worksheet (list of dicts, generator, or DataFrame) |
| `.save()` | Write all sheets and save |

### Functional API

| Function | Description |
|---|---|
| `write_worksheet(records, file_name, ...)` | Write single sheet |
| `write_worksheets(records_with_sheet_name, file_name, ...)` | Write multiple sheets |
| `validate_sheet_name(name)` | Check if sheet name is valid for Excel |

### Metadata

| Function | Description |
|---|---|
| `get_version()` | Package version |
| `get_name()` | Package name |
| `get_authors()` | Package authors |
| `get_description()` | Package description |
| `get_repository()` | Repository URL |
| `get_homepage()` | Homepage URL |
| `get_license()` | License identifier |

## Performance

![Test Result](image.png)

RustPy-XlsxWriter delivers exceptional speed improvements compared to traditional Python solutions, achieving up to **7.8x faster** processing speeds while maintaining optimal memory usage.

Based on performance testing with 1 million records:

| Operation         | Records   | Time (seconds) | Comparison      |
| ----------------- | --------- | -------------- | --------------- |
| Single Sheet      | 1,000,000 | ~13.91s        | **7x faster**   |
| Multiple Sheets   | 1,000,000 | ~12.54s        | **7.8x faster** |
| Python xlsxwriter | 1,000,000 | ~97.40s        | baseline        |

Key optimizations:

1. Rust's zero-cost abstractions and memory management
2. Native machine code compilation
3. Constant memory mode for large files
4. Lazy processing of Python iterables (including generators)
5. High-precision floating point with ryu
6. Efficient zlib compression
7. Memory safety via Rust's ownership system

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## License

This project is licensed under the MIT ![License](LICENSE).

## Acknowledgements

This project is inspired by [Rust-XlsxWriter](https://github.com/jmcnamara/rust_xlsxwriter) and [PyO3](https://github.com/pyo3/pyo3) with the help of [maturin](https://github.com/PyO3/maturin).

## Contributors

- [Rahmad Afandi](https://github.com/rahmadafandi)
