# Sheet layout

## Printing

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

## Outline Grouping

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

## Sheet View and Error Indicators

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
