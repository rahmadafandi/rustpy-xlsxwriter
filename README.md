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

# Context manager (auto-saves on exit)
with FastExcel("output.xlsx") as f:
    f.sheet("Users", user_records)
    f.sheet("Orders", order_records)
```

## Features

- **~7.3x–9.0x faster** than Python's xlsxwriter
- Fluent builder API via `FastExcel` class
- Context manager support (`with` statement) for auto-save
- Support for `str`, `int`, `float`, `bool`, `None`, `datetime` values
- Numpy scalar types (`numpy.int64`, `numpy.float64`, `numpy.bool_`) handled correctly
- Multiple sheets in a single file
- Password protection
- Freeze panes (rows & columns)
- Float formatting, custom datetime formatting & bold index columns
- Bold headers option
- Pandas and **Polars** DataFrame support with **Arrow zero-copy** optimization
- `io.BytesIO` in-memory buffer support
- Generator/iterator streaming for memory-efficient large datasets
- Optional `autofit` control for column widths

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

### Context Manager

```python
from rustpy_xlsxwriter import FastExcel

# Auto-saves when exiting the block; skips save if an exception occurs
with FastExcel("report.xlsx", password="secret") as f:
    f.format(float_format="0.00")
    f.freeze(row=1)
    f.sheet("Users", user_records)
    f.sheet("Orders", order_records)
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

### Polars DataFrame

```python
import polars as pl
from rustpy_xlsxwriter import FastExcel

df = pl.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25], "Score": [88.5, 92.3]})

# Basic
FastExcel("polars.xlsx").sheet("Data", df).save()

# With styling
(
    FastExcel("styled.xlsx")
    .format(float_format="0.00", bold_headers=True)
    .sheet("Data", df)
    .save()
)
```

### Custom Datetime Format

```python
from rustpy_xlsxwriter import FastExcel

# Default: "yyyy-mm-ddThh:mm:ss"
# Custom format:
FastExcel("report.xlsx").format(datetime_format="dd/mm/yyyy").sheet("Sheet1", records).save()
```

### Bold Headers

```python
from rustpy_xlsxwriter import FastExcel

FastExcel("report.xlsx").format(bold_headers=True).sheet("Sheet1", records).save()
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

### Disable Autofit (Performance)

For large datasets, disabling autofit can improve write performance:

```python
from rustpy_xlsxwriter import FastExcel

FastExcel("large.xlsx", autofit=False).sheet("Data", large_records).save()
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
| `FastExcel(target, *, password=None, autofit=True)` | Create writer for file path or `BytesIO` buffer. Set `autofit=False` to skip column width auto-adjustment. |
| `.format(*, float_format=None, datetime_format=None, index_columns=None, bold_headers=None)` | Set number/datetime format, bold index columns, and bold headers |
| `.freeze(*, row=None, col=None, sheet=None)` | Configure freeze panes (general or per-sheet) |
| `.sheet(name, data)` | Add a worksheet (list of dicts, generator, or DataFrame) |
| `.save()` | Write all sheets and save |

`FastExcel` also supports the context manager protocol (`with` statement). When used as a context manager, `.save()` is called automatically on exit unless an exception occurred.

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

RustPy-XlsxWriter delivers exceptional speed improvements compared to traditional Python solutions, achieving up to **~9x faster** processing speeds while maintaining optimal memory usage.

Benchmarked via [`benchmark.py`](benchmark.py) (`python benchmark.py`):

### Records (list of dicts)

| Records   | RustPy-XlsxWriter | Python xlsxwriter | Speedup          |
| --------- | ------------------ | ----------------- | ---------------- |
| 500,000   | ~2.95s             | ~26.67s           | **9.0x faster**  |
| 1,000,000 | ~5.86s             | ~52.36s           | **8.9x faster**  |

### Pandas DataFrame (Arrow zero-copy)

| Records   | RustPy-XlsxWriter | Python xlsxwriter | Speedup          |
| --------- | ------------------ | ----------------- | ---------------- |
| 500,000   | ~1.18s             | ~9.01s            | **7.6x faster**  |
| 1,000,000 | ~2.36s             | ~18.00s           | **7.6x faster**  |

### Polars DataFrame (Arrow zero-copy)

| Records   | RustPy-XlsxWriter | Python xlsxwriter | Speedup          |
| --------- | ------------------ | ----------------- | ---------------- |
| 500,000   | ~1.18s             | ~8.63s            | **7.3x faster**  |
| 1,000,000 | ~2.35s             | ~17.21s           | **7.3x faster**  |

### Key optimizations

1. **Arrow zero-copy** for DataFrames — reads memory buffers directly, no Python object conversion
2. Rust's zero-cost abstractions and memory management
3. LTO (Link-Time Optimization) and single codegen unit for maximum inlining
4. Constant memory mode for large files
5. Lazy processing of Python iterables (including generators)
6. Pre-allocated Format objects (created once, reused across all cells)
7. Dict `values()` iteration instead of per-key hash lookups
8. Correct numpy scalar type handling (no string fallback)
9. High-precision floating point with ryu
10. Efficient zlib compression

## Testing

The test suite uses `pytest` with content verification via `openpyxl`:

```bash
# Run unit tests only (fast, ~1 second)
pytest tests/ -m "not benchmark"

# Run all tests including benchmarks (~2 minutes)
pytest tests/

# Run a specific test file
pytest tests/test_dataframe.py -v
```

Test structure:

| File | Tests |
|---|---|
| `test_metadata.py` | Package metadata functions |
| `test_validation.py` | Sheet name validation (unicode, length, special chars) |
| `test_write_single.py` | Single sheet: types, generator, context manager, autofit |
| `test_write_multi.py` | Multiple sheets |
| `test_write_functional.py` | `write_worksheet()`, `write_worksheets()` |
| `test_freeze_panes.py` | Freeze panes (single & multi-sheet) |
| `test_password.py` | Password protection |
| `test_bytesio.py` | In-memory buffer I/O |
| `test_dataframe.py` | Pandas DataFrame, numpy scalar types |
| `test_polars.py` | Polars DataFrame: types, datetime, date, null, styling |
| `test_styling.py` | Float format, datetime format, bold headers, bold index columns |
| `test_benchmark.py` | 1M row benchmarks (rustpy vs xlsxwriter) |

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## License

This project is licensed under the MIT ![License](LICENSE).

## Acknowledgements

This project is inspired by [Rust-XlsxWriter](https://github.com/jmcnamara/rust_xlsxwriter) and [PyO3](https://github.com/pyo3/pyo3) with the help of [maturin](https://github.com/PyO3/maturin).

## Contributors

- [Rahmad Afandi](https://github.com/rahmadafandi)
