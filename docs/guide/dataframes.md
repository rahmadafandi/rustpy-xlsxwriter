# Data sources

## Pandas & Polars DataFrames

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

## Generator Streaming

```python
def rows():
    for i in range(1_000_000):
        yield {"id": i, "value": f"row_{i}"}

FastExcel("streamed.xlsx").sheet("Data", rows()).save()
```

> **Note:** `dedupe_strings=True` buffers the sheet, so it defeats the point of
> generator streaming. Leave it off for very large streamed exports.

## In-Memory Buffer (Web Frameworks)

```python
import io
from rustpy_xlsxwriter import FastExcel

buf = io.BytesIO()
FastExcel(buf).sheet("Sheet1", records).save()
xlsx_bytes = buf.getvalue()  # send as HTTP response
```

## String Deduplication

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
