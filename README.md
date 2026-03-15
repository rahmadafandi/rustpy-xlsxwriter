# RustPy-XlsxWriter

[![PyPI version](https://badge.fury.io/py/rustpy-xlsxwriter.svg)](https://badge.fury.io/py/rustpy-xlsxwriter)
[![Python Versions](https://img.shields.io/pypi/pyversions/rustpy-xlsxwriter.svg)](https://pypi.org/project/rustpy-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://pepy.tech/badge/rustpy-xlsxwriter)](https://pepy.tech/project/rustpy-xlsxwriter)
[![CI](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions/workflows/CI.yml/badge.svg)](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions/workflows/CI.yml)

High-performance Excel and CSV file generation for Python, powered by Rust. **~7x-9x faster** than [XlsxWriter](https://github.com/jmcnamara/XlsxWriter), **~5x faster** CSV than Python's csv module.

```python
from rustpy_xlsxwriter import FastExcel

FastExcel("report.xlsx").sheet("Sheet1", records).save()
```

## Installation

```bash
pip install rustpy-xlsxwriter
```

## Performance

Benchmarked via [`benchmark.py`](benchmark.py) — run `python benchmark.py` to reproduce:

| Output | Input type | Records | RustPy | Baseline | Speedup |
|---|---|---|---|---|---|
| **Excel** | Records (list of dicts) | 500K | ~2.99s | ~26.72s | **8.9x** |
| | | 1M | ~5.94s | ~51.92s | **8.7x** |
| | Pandas DataFrame | 500K | ~1.21s | ~9.11s | **7.6x** |
| | | 1M | ~2.41s | ~18.17s | **7.5x** |
| | Polars DataFrame | 500K | ~1.20s | ~8.59s | **7.1x** |
| | | 1M | ~2.42s | ~17.07s | **7.1x** |
| **CSV** | Records (generator) | 500K | ~0.16s | ~0.77s | **4.8x** |
| | | 1M | ~0.32s | ~1.53s | **4.8x** |

*Excel baseline: Python xlsxwriter. CSV baseline: Python csv module.*

<details>
<summary>Key optimizations</summary>

1. **Arrow zero-copy** for DataFrames — reads memory buffers directly via Arrow C Data Interface
2. **First-row type caching** for Records — detect column types once, skip type cascade
3. LTO (Link-Time Optimization) and single codegen unit
4. Constant memory mode for large files
5. Pre-allocated Format objects (created once, reused across all cells)
6. Dict `values()` iteration instead of per-key hash lookups
7. Lazy processing of Python iterables (including generators)
8. High-precision floating point with ryu
9. Efficient zlib compression

</details>

## Features

**Data Sources**
- List of dicts, generators/iterators, Pandas DataFrame, Polars DataFrame
- All Python types: `str`, `int`, `float`, `bool`, `None`, `datetime`, `date`
- Numpy scalar types (`numpy.int64`, `numpy.float64`, `numpy.bool_`)

**Formatting & Styling**
- Float number format (e.g. `"0.00"`)
- Custom datetime format (e.g. `"dd/mm/yyyy"`)
- Bold headers and bold index columns
- Freeze panes (rows, columns, per-sheet overrides)

**Output Options**
- `.xlsx` (Excel) — auto-detected from file extension
- `.csv` / `.tsv` — auto-detected, ~5x faster than Python csv module
- `io.BytesIO` in-memory buffer
- Password protection (Excel only)
- Optional column auto-fit (`autofit=True/False`)
- Multiple sheets in a single file (Excel only)

**API**
- Fluent builder via `FastExcel` class
- Context manager (`with` statement) for auto-save
- Lower-level functional API (`write_worksheet`, `write_worksheets`)

## Quick Start

```python
from rustpy_xlsxwriter import FastExcel

# Simple
FastExcel("output.xlsx").sheet("Users", [{"Name": "Alice", "Age": 30}]).save()

# Full-featured with context manager
with FastExcel("report.xlsx", password="secret") as f:
    f.format(
        float_format="0.00",
        datetime_format="dd/mm/yyyy",
        bold_headers=True,
        index_columns=["ID"],
    )
    f.freeze(row=1)
    f.sheet("Employees", employee_records)
    f.sheet("Departments", dept_records)
```

## Usage Examples

### Pandas & Polars DataFrames

```python
import pandas as pd
import polars as pl
from rustpy_xlsxwriter import FastExcel

# Pandas — Arrow zero-copy, dtype-aware
df_pd = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [88.5, 92.3]})
FastExcel("pandas.xlsx").sheet("Data", df_pd).save()

# Polars — native support, no .to_pandas() needed
df_pl = pl.DataFrame({"Name": ["Alice", "Bob"], "Score": [88.5, 92.3]})
FastExcel("polars.xlsx").sheet("Data", df_pl).save()
```

### Freeze Panes

```python
# Freeze header row on all sheets
FastExcel("frozen.xlsx").freeze(row=1).sheet("Sheet1", data).save()

# Per-sheet freeze configuration
(
    FastExcel("custom.xlsx")
    .freeze(row=1)                             # all sheets
    .freeze(row=1, col=2, sheet="Details")     # override for Details
    .sheet("Summary", summary_data)
    .sheet("Details", detail_data)
    .save()
)
```

### Generator Streaming

```python
def rows():
    for i in range(1_000_000):
        yield {"id": i, "value": f"row_{i}"}

FastExcel("streamed.xlsx").sheet("Data", rows()).save()
```

### In-Memory Buffer (Web Frameworks)

```python
import io
from rustpy_xlsxwriter import FastExcel

buf = io.BytesIO()
FastExcel(buf).sheet("Sheet1", records).save()
xlsx_bytes = buf.getvalue()  # send as HTTP response
```

### CSV / TSV Output

```python
# Auto-detected from file extension — same API, no code change
FastExcel("output.csv").sheet("Sheet1", records).save()
FastExcel("output.tsv").sheet("Sheet1", records).save()

# Or use write_csv directly
from rustpy_xlsxwriter import write_csv

write_csv(records, "output.csv")
write_csv(records, "output.csv", delimiter=";")  # custom delimiter
```

### Functional API

```python
from rustpy_xlsxwriter import write_worksheet, write_worksheets

write_worksheet(records, "output.xlsx", sheet_name="Sheet1", password="secret")

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
| `FastExcel(target, *, password=None, autofit=True)` | Create writer for file path or `BytesIO` buffer |
| `.format(*, float_format, datetime_format, index_columns, bold_headers)` | Set number/datetime format and styling |
| `.freeze(*, row=None, col=None, sheet=None)` | Configure freeze panes (general or per-sheet) |
| `.sheet(name, data)` | Add a worksheet (list of dicts, generator, or DataFrame) |
| `.save()` | Write all sheets and save |

Supports context manager (`with` statement) — auto-saves on exit, skips save on exception.

### Functional API

| Function | Description |
|---|---|
| `write_worksheet(records, file_name, ...)` | Write single Excel sheet |
| `write_worksheets(records_with_sheet_name, file_name, ...)` | Write multiple Excel sheets |
| `write_csv(records, file_name, delimiter=",")` | Write CSV/TSV file |
| `validate_sheet_name(name)` | Check if sheet name is valid for Excel |

### Supported Data Types

| Python Type | Excel Output |
|---|---|
| `str` | Text |
| `int` | Number |
| `float` | Number (with optional format) |
| `bool` | Boolean |
| `None` | Empty cell |
| `datetime.datetime` | DateTime (with optional format) |
| `datetime.date` | Date (with optional format) |
| `numpy.int64` / `numpy.float64` | Number |
| `numpy.bool_` | Boolean |
| `dict`, other | String representation |

## Examples

See [`examples/`](examples/) for 13 runnable scripts + a Jupyter notebook:

| File | Description |
|---|---|
| [`01_basic.py`](examples/01_basic.py) | Single sheet from list of dicts |
| [`02_multiple_sheets.py`](examples/02_multiple_sheets.py) | Multiple sheets in one file |
| [`03_dataframe.py`](examples/03_dataframe.py) | Pandas DataFrame with styling |
| [`04_freeze_panes.py`](examples/04_freeze_panes.py) | Freeze rows, columns, per-sheet config |
| [`05_bytesio.py`](examples/05_bytesio.py) | In-memory buffer for web frameworks |
| [`06_generator.py`](examples/06_generator.py) | Memory-efficient streaming (100K rows) |
| [`07_password.py`](examples/07_password.py) | Password-protected workbook |
| [`08_full_featured.py`](examples/08_full_featured.py) | All features combined |
| [`09_polars.py`](examples/09_polars.py) | Polars DataFrame (native support) |
| [`10_context_manager.py`](examples/10_context_manager.py) | Auto-save with `with` statement |
| [`11_datetime_format.py`](examples/11_datetime_format.py) | Custom datetime/date formatting |
| [`12_bold_headers.py`](examples/12_bold_headers.py) | Bold header row |
| [`13_autofit.py`](examples/13_autofit.py) | Column auto-fit toggle |
| [`14_csv_tsv.py`](examples/14_csv_tsv.py) | CSV/TSV output (~5x faster) |
| [`quickstart.ipynb`](examples/quickstart.ipynb) | Jupyter notebook walkthrough |

## Testing

```bash
# Unit tests (~1 second)
pytest tests/ -m "not benchmark"

# All tests including benchmarks
pytest tests/

# Benchmark only
python benchmark.py
```

<details>
<summary>Test structure (86 tests)</summary>

| File | Coverage |
|---|---|
| `test_metadata.py` | Package metadata functions |
| `test_validation.py` | Sheet name validation (unicode, length, special chars) |
| `test_write_single.py` | Single sheet: all types, generator, context manager, autofit |
| `test_write_multi.py` | Multiple sheets |
| `test_write_functional.py` | Functional API |
| `test_freeze_panes.py` | Freeze panes (single & multi-sheet) |
| `test_password.py` | Password protection |
| `test_bytesio.py` | In-memory buffer I/O |
| `test_dataframe.py` | Pandas DataFrame, numpy scalar types |
| `test_polars.py` | Polars DataFrame: types, datetime, date, null, styling |
| `test_styling.py` | Float format, datetime format, bold headers, index columns |
| `test_benchmark.py` | Performance benchmarks (Records + Pandas + Polars vs xlsxwriter) |

</details>

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## License

This project is licensed under the MIT [License](LICENSE).

## Acknowledgements

This project is powered by [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter), [PyO3](https://github.com/pyo3/pyo3), and [maturin](https://github.com/PyO3/maturin).
