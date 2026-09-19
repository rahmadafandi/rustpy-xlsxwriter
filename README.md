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

With the GIL both columns are flat: threads take turns, so the 1M records cost
the same whether one worker writes them or eight.

Without it both writers spread out — `xlsxwriter` is pure Python and gets faster
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
- CPython 3.8+ — prebuilt wheels for Linux (glibc/musl), macOS, Windows
- Free-threaded builds (`python3.14t`) — parallel writes, see [Concurrency](#concurrency)

**API**
- Typed: ships `py.typed`, so mypy and Pyright check calls into it
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

### Column widths

Override `autofit` with explicit widths (Excel character units):

```python
from rustpy_xlsxwriter import FastExcel

(
    FastExcel("out.xlsx")
    .sheet("RawData", rows, column_width=15)                    # uniform
    .sheet("Meta", meta, column_widths={"row": 7, "var": 22})   # per-column (by name)
    .sheet("Other", rows, column_widths=[7, 22, 40])            # per-column (positional)
    .save()
)
```

`column_widths` overrides `column_width` per named/positional column, and explicit
widths win over `autofit=True`. Unknown column names and out-of-range list indices
emit a warning and are skipped. For `write_worksheets`, pass `column_width` /
`column_widths` as dicts keyed by sheet name (with a `"general"` fallback key).

### Cell formatting

Build a reusable `Format` and attach it per column or to the header row:

```python
from rustpy_xlsxwriter import FastExcel, Format

money = Format().set_num_format("$#,##0.00").set_font_color("#006600")
header = Format().set_bold().set_background_color("#1F4E78").set_font_color("white")

(
    FastExcel("out.xlsx")
    .sheet("Products", rows, header_format=header, column_formats={"price": money})
    .save()
)
```

`Format` chains setters (font, fill, border, alignment, number format). Colors
accept `"#RRGGBB"` or names (`"red"`); enum-valued setters take lowercase strings
(`set_align("center")`, `set_border("thin")`). A column's format wins over
`float_format` / `datetime_format`. For `write_worksheets`, pass `column_formats` /
`header_format` as dicts keyed by sheet name (with a `"general"` fallback key).

> **Note:** a column format wins *entirely* over the automatic number format. On
> a date/datetime column, a `Format` without `set_num_format(...)` makes cells
> show the raw Excel serial number — chain `.set_num_format("yyyy-mm-dd")` (or
> similar) to keep a date display.

### Row Layout: Merged Headers, Borders, Banding

Crosstab and summary reports need structure above and across the data rows.
Everything here is declared up front, per sheet:

```python
from rustpy_xlsxwriter import FastExcel, Format

banner = Format().set_bold().set_align("center").set_background_color("#1F4E78")
under_header = Format().set_border_bottom("thin")

(
    FastExcel("crosstab.xlsx")
    .sheet(
        "Survey",
        rows,
        header_row=1,                                    # leave row 0 for banners
        merge_ranges=[(0, 1, 0, 2, "Gender", banner)],   # "Gender" spans B:C
        row_heights={0: 28, 1: 22},
        row_formats={1: under_header},                   # rule under the header
        banded_rows="#F2F2F2",                           # alternating fill
    )
    .save()
)
```

| Option | Effect |
|---|---|
| `header_row` | 0-based row for headers; data starts on the next row |
| `merge_ranges` | `(first_row, first_col, last_row, last_col, value[, format])` |
| `row_heights` | `{row_index: height}` in points |
| `row_formats` | `{row_index: Format}` — borders under headers, above totals |
| `banded_rows` | Background colour for every other data row |
| `autofilter` | Filter dropdowns over the header row and its data |

| `url_columns` | Column names whose text cells become clickable links |
| `totals_row` | Aggregate formulas in a row below the data |
| `formula_columns` | Computed columns appended after the data |

`autofilter=True` sizes its own range from the rows actually written, so it
follows `header_row` and needs no manual bounds:

```python
FastExcel("report.xlsx").sheet("Data", rows, autofilter=True, freeze_row=1).save()
```

### Formulas

Append computed columns. `{row}` becomes that row's sheet row, `{first}` the
first data row:

```python
(
    FastExcel("report.xlsx")
    .sheet(
        "Sales",
        rows,
        formula_columns={
            "total":   "=A{row}*B{row}",
            "running": "=SUM(B${first}:B{row})",
        },
    )
    .save()
)
```

The text goes to Excel unchanged, so **anything Excel accepts works** — nested
calls, `SUMIFS`, `INDEX`/`MATCH`, cross-sheet references. Modern functions are
handled for you: the 131 "future" functions (`IFS`, `TEXTJOIN`, `MAXIFS`,
`STDEV.P`, …) get their required `_xlfn.` prefix, and the 30 dynamic-array ones
(`XLOOKUP`, `UNIQUE`, `SORT`, `FILTER`, …) additionally get array-formula markup
and the `xl/metadata.xml` part. Writing those by hand is the usual way to end up
with a file Excel refuses to open.

`totals_row` values may also be raw formulas — a value starting with `=`, with
`{col}` the column letter and `{first}`/`{last}` the data range:

```python
totals_row={"qty": "sum", "price": "=ROUND(AVERAGE({col}{first}:{col}{last}),2)"}
```

Two things to know:

**Structure is validated, names are not.** Unbalanced parentheses or quotes and
an empty formula raise at write time, naming the column and echoing the
formula. Function names are not checked: `LAMBDA` and `LET` bind their own,
workbooks carry user-defined functions, and Excel keeps adding to the list, so a
whitelist would reject valid formulas. `=NOTAFUNC(A1)` therefore reaches the
file and shows `#NAME?` in that cell.

For context, a malformed formula never corrupts the file — every case tested
opens fine and shows an error value in the one cell. Validation just moves the
discovery from "when someone opens the report" to "when the export runs", which
is why it stops at the checks that cannot produce a false positive.

**There is no `{last}` in `formula_columns`.** Those cells are written while
rows are still streaming, so the final row is unknown; the placeholder raises
with that explanation. Use `totals_row` for whole-column formulas.

### Totals Row

Aggregate formulas in a row below the data. The ranges are derived from the
rows actually written:

```python
(
    FastExcel("report.xlsx")
    .sheet(
        "Sales",
        rows,
        totals_row={"amount": "sum", "qty": "sum"},
        totals_label="Total",
        totals_format=Format().set_bold().set_border_top("thin"),
    )
    .save()
)
# amount column gets  =SUM(C2:C101)
```

Aggregates: `sum`, `average` (`avg`/`mean`), `count`, `min`, `max`, `product`,
`stdev`. An unknown one raises rather than writing a broken formula. The row is
skipped entirely when there are no data rows, since the range would be empty,
and `autofilter` deliberately stops above it so sorting never drags the total
into the data.

`totals_format` exists because the totals row index depends on how much data
there was, so `row_formats` cannot reach it.

> **The formulas carry no computed result.** This library does not evaluate
> them; Excel and LibreOffice do, on open. Readers that trust the cached value —
> `pandas.read_excel`, `openpyxl` with `data_only=True` — get `None`, not a
> number. That is deliberate: the cached result is written empty rather than
> left at the underlying crate's default of `0`, which would look like a real
> total of zero. Use the totals row for files people will open, not for a
> machine-readable handoff.

### Hyperlinks

Name the columns that hold links; the cell text stays the URL:

```python
FastExcel("report.xlsx").sheet("Docs", rows, url_columns=["homepage"]).save()
```

Accepts what Excel accepts — `http(s)://`, `mailto:`, and `internal:Sheet2!A1`
to jump to another sheet. Anything Excel would reject (ordinary text, a blank,
or a URL past its 2083-character limit) is written as plain text instead, so one
stray value in a column of thousands never aborts the export. Links keep their
banding and column format.

**Ordering is enforced, not assumed.** Sheets are written row by row and a row
that has been flushed cannot be revisited — `rust_xlsxwriter` would drop a late
`merge_range` with only a message on stderr. So a merge range that reaches
`header_row` or below raises `ValueError` telling you what to raise
`header_row` to, rather than silently losing the banner.

**Banding is applied per cell, not per row.** A cell carrying its own format
ignores the row's, so shading a row with `set_row_format` leaves holes in
exactly the columns that have a number format. `banded_rows` instead shades
each cell, so float, integer, boolean, datetime and explicitly-formatted
columns all stay banded. That costs roughly 20% write time; leave it off when
you don't need it.

For `write_worksheets`, every one of these takes a dict keyed by sheet name
(with a `"general"` fallback key).

#### Showing something other than the URL

Pass a mapping instead of a list to link one column and display another:

```python
write_worksheet(
    rows,
    "out.xlsx",
    url_columns={"url": "product_name"},   # cell reads "Widget", links to the URL
)
```

An unknown display-text column warns and falls back to showing the URL, so a
typo costs a label rather than the export.

### Notes and Images

```python
write_worksheet(
    rows,
    "report.xlsx",
    notes={"revenue": "Net of returns and credit notes"},
    images=[{"path": "logo.png", "row": 0, "col": 5, "scale": 0.5}],
)
```

A note lands on the column's **header** cell — where you say what a column
means without widening it or adding a legend sheet. Pass a dict instead of a
string for `author`, `width`, `height`, `visible` or `background_color`.

An image is placed by `row`/`col` rather than by column name, since it floats
above the grid. Give it a `path` or raw `data` — the second is what a web
handler has, a logo already in memory. Identical images are stored once.
Neither costs constant-memory mode.

### Outline Grouping

The collapsible `+`/`-` brackets in Excel's margin:

```python
write_worksheet(
    rows,
    "out.xlsx",
    outline={
        "columns": [{"from": "q1", "to": "q3", "collapsed": True}],
        "rows": [{"from": 2, "to": 8}],
    },
)
```

Rows are given by 0-based sheet index, like `row_heights`; columns by header
name, like everything else keyed by column. `symbols_above` and
`symbols_to_left` choose which side the summary sits on.

> **A row group turns off constant-memory mode for that sheet**, the same
> trade-off as `dedupe_strings`. The constant-memory row writer emits `hidden`
> but not `outlineLevel`, so a collapsed group there would leave hidden rows
> with no bracket to reopen them. Column groups carry no such cost.

### Data Validation

```python
write_worksheet(
    rows,
    "out.xlsx",
    data_validations={
        "status": {"type": "list", "values": ["open", "closed", "blocked"],
                   "input_message": "Pick one"},
        "qty":    {"type": "whole_number", "criteria": ">=", "value": 0,
                   "error_message": "Quantity cannot be negative"},
        "score":  {"type": "decimal", "criteria": "between", "min": 0, "max": 100},
    },
)
```

Types: `list` (the dropdown), `whole_number`, `decimal`, `text_length`,
`custom`, `any`. Every rule also takes `input_title`, `input_message`,
`error_title`, `error_message`, `error_style` (`stop`, `warning`,
`information`), `ignore_blank` and `show_dropdown`.

Rules cover the column's data rows, never the header. Two limits Excel
imposes are raised rather than written into a file it would refuse to open: an
inline `list` is capped at 255 characters including separators, and a fraction
given to `whole_number` or `text_length` is rejected instead of truncated.

### Sheet View and Error Indicators

```python
write_worksheet(
    rows,
    "out.xlsx",
    sheet_view={"tab_color": "#C00000", "gridlines": False, "zoom": 120},
    ignore_errors=["sku"],
)
```

`sheet_view` covers the screen — `tab_color`, `gridlines`, `zoom`,
`right_to_left`, `hidden`, `selected` — and is kept apart from `page_setup`,
which covers paper.

`ignore_errors` removes Excel's green triangles. A list of column names means
`number_stored_as_text`, which is why anyone wants it: an ID, SKU or postcode
column is digits stored as text on purpose and Excel flags every cell. Pass a
dict to name another error — one per column, since Excel allows a single
ignore rule per cell.

### Conditional Formatting

Rules are given per column and cover that column's data rows — never the
header — with the range taken from the rows actually written, so there are no
bounds to compute:

```python
write_worksheet(
    rows,
    "report.xlsx",
    conditional_formats={
        "revenue": {"type": "data_bar", "color": "#638EC6"},
        "margin":  {"type": "3_color_scale"},
        "overdue": {"type": "cell", "criteria": ">", "value": 30,
                    "format": Format().set_background_color("#FFC7CE")},
        "status":  {"type": "text", "criteria": "contains", "value": "FAIL",
                    "format": Format().set_bold()},
    },
)
```

Types: `cell`, `data_bar`, `2_color_scale`, `3_color_scale`, `text`, `top`,
`average`, `duplicate`, `unique`. Pass a list to stack several on one column:

```python
conditional_formats={"score": [{"type": "data_bar"},
                               {"type": "top", "value": 3, "format": gold}]}
```

An unknown column warns and is skipped — a rule that cannot be placed costs
shading, not the export. An unknown type or criteria raises, since that is a
mistake in the code rather than in the data.

### Printing

Excel has about twenty page-setup settings, so they arrive as one mapping
rather than twenty keywords:

```python
write_worksheet(
    rows,
    "report.xlsx",
    page_setup={
        "landscape": True,
        "fit_to_pages": (1, 0),      # one page wide, as many tall as needed
        "repeat_rows": 0,            # header on every printed page
        "margins": {"left": 0.5},    # omitted sides keep Excel's defaults
        "footer": "&RPage &P of &N",
    },
)
```

Keys: `landscape`, `paper_size`, `margins`, `print_area`, `repeat_rows`,
`repeat_columns`, `fit_to_pages`, `scale`, `center_horizontally`,
`center_vertically`, `print_gridlines`, `print_headings`,
`first_page_number`, `header`, `footer`.

`repeat_rows` and `repeat_columns` take an index or a `(first, last)` pair —
the first is what stops a long report losing its header after page one. An
unknown key raises, and so does `scale` together with `fit_to_pages`, which
Excel cannot honour at once.

### String Deduplication

By default every sheet is written in constant-memory mode: strings go inline
into the sheet XML and nothing is buffered. Passing `dedupe_strings=True` takes
that sheet out of constant-memory mode and stores each distinct string once in
the workbook's shared-string table instead.

```python
(
    FastExcel("report.xlsx")
    .sheet("Events", events, dedupe_strings=True)   # lots of repeated text
    .sheet("Raw", raw_rows)                         # default: streamed
    .save()
)
```

Measured on 50k rows — the uncompressed XML shrinks a lot, but `.xlsx` is a zip
and deflate already collapses repetition, so the **on-disk** win is modest and
can even be negative:

| Data | Uncompressed | On disk | Write time |
|---|---|---|---|
| Short repeats (4 distinct values) | −11% | **+2%** | 1.7x |
| Long repeats (20 distinct, 120 chars) | −56% | −5% | 1.1x |
| All-unique strings | +11% | −1% | 1.5x |
| 20 repeated text columns | −39% | −9% | 1.1x |

Worth enabling for sheets with many repeated *long* strings, or when the
consumer parses the uncompressed XML. Not worth it for short categorical values.
Off by default because it buffers the whole sheet in memory — measure on your
own data before turning it on for a large export.

For `write_worksheets`, pass `dedupe_strings` as a dict keyed by sheet name
(with a `"general"` fallback key).

### Generator Streaming

```python
def rows():
    for i in range(1_000_000):
        yield {"id": i, "value": f"row_{i}"}

FastExcel("streamed.xlsx").sheet("Data", rows()).save()
```

> **Note:** `dedupe_strings=True` buffers the sheet, so it defeats the point of
> generator streaming. Leave it off for very large streamed exports.

### In-Memory Buffer (Web Frameworks)

```python
import io
from rustpy_xlsxwriter import FastExcel

buf = io.BytesIO()
FastExcel(buf).sheet("Sheet1", records).save()
xlsx_bytes = buf.getvalue()  # send as HTTP response
```

### Type checking

The package ships `py.typed`, so mypy and Pyright check calls into it without
any extra stub package.

### CSV / TSV Output

```python
# Auto-detected from file extension
FastExcel("output.csv").sheet("Sheet1", records).save()
FastExcel("output.tsv").sheet("Sheet1", records).save()

# A buffer has no extension, so name the format — this is how you get CSV
# out of a web handler without touching the filesystem
buf = io.BytesIO()
FastExcel(buf, output_format="csv").sheet("Sheet1", records).save()

# Or use write_csv directly
from rustpy_xlsxwriter import write_csv

write_csv(records, "output.csv")
write_csv(records, "output.csv", delimiter=";")  # custom delimiter
```

`output_format` accepts `"xlsx"`, `"csv"` or `"tsv"` and overrides the
extension, so a `.txt` target can hold CSV. For any other delimiter, call
`write_csv` directly.

#### Excel on Windows and the byte order mark

A UTF-8 CSV without a BOM opens in Excel as the system code page, which turns
every non-ASCII character into mojibake. `bom=True` fixes it:

```python
write_csv(records, "out.csv", bom=True)      # "Café" stays "Café" in Excel
```

Off by default, so output stays byte-identical for anything that parses it.

#### Selecting columns

```python
write_csv(records, "out.csv", columns=["sku", "name"])  # subset and order
write_csv(records, "out.csv", header=False)             # no header row
```

`columns` works on every input type and stays zero-copy on the Arrow path, so
you no longer have to slice a DataFrame in Python first — which is what threw
away the speed this library is for. A name that is not in the data raises
`ValueError`: unlike the styling options this one decides the shape of the
file, so a silent drop would hand back something that looks complete.

### Missing values, NaN and infinity

Excel has no cell type for `NaN` or `inf`, so both are written as an **empty
cell**. That is the default, and it means a column of missing data is
indistinguishable from a column of blanks. Name a representation to keep them
apart:

```python
FastExcel("out.xlsx").format(na_rep="N/A", inf_value="INF").sheet("S", df).save()

write_csv(df, "out.csv", na_rep="N/A", inf_value="INF")
```

`na_rep` covers `None`, an Arrow null and a float `NaN` as one setting, on
purpose: pandas turns `NaN` in a float column into an Arrow null, so a knob
that caught only true `NaN` would do nothing on the most common input there is.
`inf_value` is separate — infinity is a value, not a missing one — and `-inf`
takes the same text with a leading `-`, matching Excel's own `INF`/`-INF`.

CSV carries no formatting, so every Excel-only option is dropped —
`float_format`, `column_formats`, `header_format`, freeze panes, merges,
banding, row heights and formats, `password`, `dedupe_strings`. Only
`delimiter` and `sanitize_formulas` apply. Switching a target from `.xlsx` to
`.csv` therefore silently changes the output, so the builder warns and names
what it discarded:

```python
FastExcel("out.csv").format(float_format="0.00").sheet("S", rows).save()
# UserWarning: CSV/TSV output ignores Excel-only options: float_format. …
```

The data is still written correctly — only the styling is gone.

### Functional API

```python
from rustpy_xlsxwriter import write_worksheet, write_worksheets

write_worksheet(records, "output.xlsx", sheet_name="Sheet1", password="secret")

write_worksheets(
    [("Sheet1", records1), ("Sheet2", records2)],
    "output.xlsx",
    freeze_panes={"general": {"row": 1, "col": 0}},
)
```

## API Reference

### `FastExcel` Class

| Method | Description |
|---|---|
| `FastExcel(target, *, output_format=None, password=None, autofit=True, sanitize_formulas=False, bom=False, columns=None, header=True)` | Create writer for file path or `BytesIO` buffer |
| `.format(*, float_format, datetime_format, index_columns, bold_headers, na_rep, inf_value)` | Set number/datetime format and styling |
| `.freeze(*, row=None, col=None, sheet=None)` | Configure freeze panes (general or per-sheet) |
| `.sheet(name, data)` | Add a worksheet (list of dicts, generator, or DataFrame) |
| `.save()` | Write all sheets and save |

Supports context manager (`with` statement) — auto-saves on exit, skips save on exception.

### Functional API

| Function | Description |
|---|---|
| `write_worksheet(records, file_name, ...)` | Write single Excel sheet |
| `write_worksheets(records_with_sheet_name, file_name, ...)` | Write multiple Excel sheets |
| `write_csv(records, file_name, delimiter=",", bom=False, columns=None, header=True, na_rep=None, inf_value=None)` | Write CSV/TSV file |
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

See [`examples/`](examples/) for 16 runnable scripts + a Jupyter notebook:

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
| [`15_column_widths.py`](examples/15_column_widths.py) | Uniform & per-column widths |
| [`16_cell_formats.py`](examples/16_cell_formats.py) | Cell formatting (Format class) |
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
<summary>Test structure</summary>

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
| `test_output_format.py` | Explicit `output_format`, CSV/TSV into a buffer |
| `test_type_stubs.py` | `.pyi` kept in step with the compiled extension |
| `test_csv_options.py` | CSV `bom`, `columns`, `header` across all four input paths |
| `test_nan_inf.py` | `na_rep` / `inf_value` on every write path |
| `test_page_setup.py` | Page and print settings, and their validation |
| `test_conditional_formats.py` | Rule types, ranges, and validation |
| `test_sheet_view.py` | Screen presentation and error indicators |
| `test_data_validation.py` | Dropdowns, numeric rules, messages, limits |
| `test_outline.py` | Row and column groups, and the constant-memory swap |
| `test_notes_images.py` | Header notes, images from path or bytes |
| `test_benchmark.py` | Performance benchmarks (Records + Pandas + Polars vs xlsxwriter) |

</details>

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## Support

If this project saves you time, consider supporting its development via [Saweria](https://saweria.co/rahmadafandi) ☕ — or use the **Sponsor** button at the top of the repository.

## License

This project is licensed under the MIT [License](LICENSE).

## Acknowledgements

This project is powered by [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter), [PyO3](https://github.com/pyo3/pyo3), and [maturin](https://github.com/PyO3/maturin).
