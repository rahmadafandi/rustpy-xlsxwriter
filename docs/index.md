# RustPy-XlsxWriter

High-performance Excel and CSV generation for Python, written in Rust.

```python
from rustpy_xlsxwriter import FastExcel

FastExcel("report.xlsx").sheet("Sheet1", records).save()
```

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

Benchmarked via [`benchmark.py`](https://github.com/rahmadafandi/rustpy-xlsxwriter/blob/master/benchmark.py) — run `python benchmark.py` to reproduce:

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
further, so the speedup against `xlsxwriter` grows from **7.9x to 12.1x** at
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
| 1 | 8.88s | 65.40s | 7.4x |
| 2 | 6.56s | 72.79s | 11.1x |
| 4 | 5.79s | 168.62s | 29.1x |
| 8 | **5.26s** | 159.41s | **30.3x** |

**Python 3.14t — free-threaded**

| Workers | RustPy | xlsxwriter | Speedup |
|---|---|---|---|
| 1 | 9.00s | 71.54s | 7.9x |
| 2 | 4.61s | 43.07s | 9.3x |
| 4 | 2.63s | 27.67s | 10.5x |
| 8 | **1.63s** | 19.72s | **12.1x** |

**With the GIL, RustPy still spreads.** It goes from 8.88s to 5.26s, a 1.7x
gain, because the save — XML assembly and deflate, about two thirds of a write
— runs with the GIL released. The rows are read through Python objects and stay
serialised, which is why it is 1.7x and not 8x.

`xlsxwriter` is pure Python throughout and gets *worse* past two workers, from
65.40s to 159.41s. That is thread contention, not memory: free RAM never dropped
below 7.9 GB during the run and swap stayed at zero.

**Without the GIL both spread properly.** RustPy goes from 9.00s to 1.63s
(**5.5x**) and `xlsxwriter` from 71.54s to 19.72s (**3.6x**) — more of RustPy's
work is Rust rather than interpreted bytecode, so its advantage widens from 7.9x
at one worker to 12.1x at eight.

The cost is a single write being ~1% slower on 3.14t, the usual price of the
free-threaded interpreter. So: one big export is a wash, many concurrent exports
favour the free-threaded build — and either way more workers now helps.

Measured on 8 physical cores (16 threads), 16 GB, on a different machine than
the table above — read each row's columns against each other, not against the
rows above.

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

## Functional API

```python
from rustpy_xlsxwriter import write_worksheet, write_worksheets

write_worksheet(records, "output.xlsx", sheet_name="Sheet1", password="secret")

write_worksheets(
    [("Sheet1", records1), ("Sheet2", records2)],
    "output.xlsx",
    freeze_panes={"general": {"row": 1, "col": 0}},
)
```

## Type checking

The package ships `py.typed`, so mypy and Pyright check calls into it without
any extra stub package.

## Where to go next

- **[Data sources](guide/dataframes.md)** — DataFrames, generators, buffers
- **[Formatting](guide/formatting.md)** — fonts, colours, widths, banding, conditional formats
- **[Formulas and links](guide/formulas.md)** — computed columns, totals, hyperlinks
- **[Charts and visuals](guide/visuals.md)** — charts, sparklines, notes, images
- **[Sheet layout](guide/layout.md)** — printing, outline groups, sheet view
- **[Data integrity](guide/data-integrity.md)** — validation, missing values
- **[CSV and TSV](guide/csv.md)** — delimiters, BOM, column selection
- **[API reference](api.md)** — every function and option
