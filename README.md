# RustPy-XlsxWriter

[![PyPI version](https://badge.fury.io/py/rustpy-xlsxwriter.svg)](https://badge.fury.io/py/rustpy-xlsxwriter)
[![Python Versions](https://img.shields.io/pypi/pyversions/rustpy-xlsxwriter.svg)](https://pypi.org/project/rustpy-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://pepy.tech/badge/rustpy-xlsxwriter)](https://pepy.tech/project/rustpy-xlsxwriter)
[![CI](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions/workflows/CI.yml/badge.svg)](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions/workflows/CI.yml)
[![Docs](https://img.shields.io/badge/docs-API%20reference-blue)](https://rahmadafandi.github.io/rustpy-xlsxwriter/)
[![Donate](https://img.shields.io/badge/donate-Saweria-orange)](https://saweria.co/rahmadafandi)

High-performance Excel and CSV file generation for Python, powered by Rust. **~7x-9x faster** than [XlsxWriter](https://github.com/jmcnamara/XlsxWriter), **~5x faster** CSV (records) than Python's `csv` module, and **~12x faster** Pandas DataFrame → CSV than `pandas.to_csv` via zero-copy Arrow. On free-threaded Python the gap widens to **~9.4x**, because it parallelises better than a pure-Python writer can — see [Concurrency](#concurrency).

```python
from rustpy_xlsxwriter import FastExcel

FastExcel("report.xlsx").sheet("Sheet1", records).save()
```

**[Documentation](https://rahmadafandi.github.io/rustpy-xlsxwriter/)** · [Guide](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/dataframes/) · [API reference](https://rahmadafandi.github.io/rustpy-xlsxwriter/api/)

## Installation

```bash
pip install rustpy-xlsxwriter
```

Prebuilt wheels cover CPython on Linux (glibc and musl), macOS and Windows.

### Free-threaded Python

Python 3.14 makes the free-threaded build officially supported. It is a
*separate* interpreter — `python3.14t` — not the default one, so you have to
install it deliberately:

```bash
uv python install 3.14t          # or your distro's python3.14-freethreading package
uv venv --python 3.14t
uv pip install rustpy-xlsxwriter # installs the cp314t wheel
```

Check which one you are on:

```python
import sys
sys.version                # "... free-threading build ..." on 3.14t
sys._is_gil_enabled()      # False on a free-threaded build
```

Nothing in the API changes. See [Concurrency](#concurrency) for what it buys.

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
| | Pandas DataFrame | 1M | — | — | **~12x**† |
| | Polars DataFrame | 1M | — | — | (use Polars' native `write_csv` — already faster)† |

*Baselines: Excel → Python `xlsxwriter`; Records CSV → Python `csv` module; Pandas DataFrame CSV → `DataFrame.to_csv()`.*

**Every row above is single-threaded, and the GIL keeps it that way.** On
free-threaded Python both writers spread across threads — but RustPy spreads
further, so the speedup against `xlsxwriter` grows from **7.4x to 9.4x** at
eight threads. Full numbers in [Concurrency](#concurrency).

† *DataFrame → CSV rows measured on a separate machine — only speedup ratio shown. The Pandas path goes through zero-copy Arrow C Data Interface.*

### Concurrency

Every row below writes **the same 1,000,000 records** — the Records row from the
table above — just spread over more workers. `xlsxwriter` is measured at each
worker count too. Reproduce with `python benchmark.py --concurrent` under each
interpreter:

**Python 3.14 — standard build**

| Workers | RustPy | xlsxwriter | Speedup |
|---|---|---|---|
| 1 | 5.40s | 43.47s | 8.0x |
| 2 | 5.72s | 40.72s | 7.1x |
| 4 | 5.70s | 39.79s | 7.0x |
| 8 | 5.69s | 39.55s | 6.9x |

**Python 3.14t — free-threaded**

| Workers | RustPy | xlsxwriter | Speedup |
|---|---|---|---|
| 1 | 6.10s | 45.00s | 7.4x |
| 2 | 3.20s | 27.17s | 8.5x |
| 4 | 2.09s | 18.83s | 9.0x |
| 8 | **1.72s** | 16.21s | **9.4x** |

> **The standard-build table above predates 0.7.1** and shows RustPy flat
> across workers. It no longer is: the save — XML assembly and deflate, about
> two thirds of a write — now runs with the GIL released, so concurrent writers
> overlap there. Measured after the change on 8 physical cores: 8.84s at one
> worker to 5.42s at eight, a 1.6x gain where there used to be none. Re-run
> `python benchmark.py --concurrent` to refresh these numbers for your machine.

With the GIL, RustPy now spreads part of the work: the rows are read through
Python objects and stay serialised, but the save does not. `xlsxwriter` is pure
Python throughout, so its threads take turns.

Without the GIL both writers spread out — `xlsxwriter` is pure Python and gets faster
too, from 45.00s to 16.21s (**2.8x**). RustPy goes from 6.10s to 1.72s
(**3.5x**), because more of its work is Rust rather than interpreted bytecode.
That is why the advantage widens rather than staying put: 7.4x at one worker,
9.4x at eight.

The cost is a single write being ~10% slower on 3.14t, the usual price of the
free-threaded interpreter. So: one big export favours the standard build, many
concurrent exports favour the free-threaded one.

Measured on 10 physical cores, on a different machine than the table above — read
each row's columns against each other, not against the rows above.

Two caveats: Polars has no free-threaded wheel yet, so that input path is
unavailable on 3.14t (Pandas, Arrow and records all work). And the free-threaded
build is younger — treat it as the newer option it is.

<details>
<summary>Key optimizations</summary>

1. **Arrow zero-copy** for DataFrames — reads memory buffers directly via Arrow C Data Interface (Excel and CSV paths)
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
- Page and print setup (`page_setup=`): orientation, margins, repeat rows, fit-to-pages, headers/footers
- Conditional formatting (`conditional_formats=`): data bars, colour scales, cell/text/top/average rules
- Data validation (`data_validations=`): dropdowns, numeric and text-length rules
- Outline grouping (`outline=`): collapsible row and column groups
- Header notes (`notes=`) and cell-anchored images (`images=`, path or bytes)
- Sparklines (`sparklines=`): a one-cell trend chart per row
- Charts (`charts=`): column, bar, line, pie, scatter and more, series by column name
- Sheet view (`sheet_view=`): tab colour, gridlines, zoom, hidden
- Suppress error triangles (`ignore_errors=`), e.g. numbers stored as text

**Output Options**
- `.xlsx` (Excel) — auto-detected from file extension
- `.csv` / `.tsv` — auto-detected; ~5x faster than Python `csv` (records), ~12x faster than `pandas.to_csv` (Pandas DataFrame, via Arrow zero-copy)
- `io.BytesIO` in-memory buffer, with `output_format` to put CSV in one
- CSV byte order mark (`bom=True`) so Excel on Windows reads UTF-8
- Column selection and ordering (`columns=`), header row toggle (`header=`)
- Text for missing values and infinity (`na_rep=`, `inf_value=`)
- Password protection (Excel only)
- Optional column auto-fit (`autofit=True/False`)
- Multiple sheets in a single file (Excel only)

**Runtime**
- CPython 3.9+ — prebuilt wheels for Linux (glibc/musl), macOS, Windows
- Free-threaded builds (`python3.14t`) — parallel writes, see [Concurrency](#concurrency)

**API**
- Typed: ships `py.typed`, so mypy and Pyright check calls into it
- Fluent builder via `FastExcel` class
- Context manager (`with` statement) for auto-save
- Lower-level functional API (`write_worksheet`, `write_worksheets`)

## Quick Start

```python
from rustpy_xlsxwriter import FastExcel

# One line
FastExcel("report.xlsx").sheet("Sheet1", records).save()

# Or with options
(
    FastExcel("report.xlsx")
    .format(float_format="0.00", bold_headers=True)
    .freeze(row=1)
    .sheet("Users", users)
    .sheet("Orders", orders)
    .save()
)
```

Everything else — formatting, charts, validation, printing, CSV options — is in
the **[documentation](https://rahmadafandi.github.io/rustpy-xlsxwriter/)**:

| | |
|---|---|
| [Data sources](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/dataframes/) | DataFrames, generators, in-memory buffers |
| [Formatting](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/formatting/) | fonts, widths, banding, conditional formats |
| [Formulas and links](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/formulas/) | computed columns, totals rows, hyperlinks |
| [Charts and visuals](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/visuals/) | charts, sparklines, notes, images |
| [Sheet layout](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/layout/) | printing, outline groups, sheet view |
| [Data integrity](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/data-integrity/) | validation, missing values |
| [CSV and TSV](https://rahmadafandi.github.io/rustpy-xlsxwriter/guide/csv/) | delimiters, BOM, column selection |
| [API reference](https://rahmadafandi.github.io/rustpy-xlsxwriter/api/) | every function and option |

## Testing

```bash
pytest tests/ -m "not benchmark"   # unit tests
pytest tests/                      # including benchmarks
python benchmark.py                # standalone benchmark
```

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## Support

If this project saves you time, consider supporting its development via [Saweria](https://saweria.co/rahmadafandi) ☕ — or use the **Sponsor** button at the top of the repository.

## License

This project is licensed under the MIT [License](LICENSE).

## Acknowledgements

This project is powered by [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter), [PyO3](https://github.com/pyo3/pyo3), and [maturin](https://github.com/PyO3/maturin).
