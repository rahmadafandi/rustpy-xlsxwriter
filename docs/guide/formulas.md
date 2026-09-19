# Formulas and links

## Formulas

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

## Totals Row

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

## Hyperlinks

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

## Showing something other than the URL

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
