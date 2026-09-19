# Charts and visuals

## Charts

Series are named by column and cover that column's data rows, so there are no
ranges to compute:

```python
write_worksheet(
    rows,
    "report.xlsx",
    charts=[{
        "type": "column",
        "series": ["q1", "q2", "q3", "q4"],
        "categories": "region",
        "title": "Quarterly revenue",
        "y_axis": "USD",
    }],
)
```

Types: `area`, `bar`, `column`, `line` (each with `_stacked` and
`_percent_stacked` too), `pie`, `doughnut`, `radar`, `radar_with_markers`,
`radar_filled`, `scatter`, `scatter_smooth`, `stock`. Also takes `row`/`col`,
`title`, `x_axis`, `y_axis`, `width`, `height`, `style` and `legend`.

A series name links to the column's **header cell**, so the legend follows the
header if it is ever edited; pass `{"values": "q1", "name": "Quarter 1"}` to
set it outright. Left unplaced, a chart lands one column clear of the data
rather than on top of it.

A scatter chart needs `categories` — they are its x values, not labels — and
saying so is refused up front rather than after the rows are written.

## Sparklines

A one-cell chart per data row. Leave a column empty in the records and point
it at the span it should summarise:

```python
rows = [{"name": "a", "q1": 1, "q2": 5, "q3": 3, "q4": 8, "trend": None}, ...]

write_worksheet(
    rows,
    "report.xlsx",
    sparklines={"trend": {"from": "q1", "to": "q4", "type": "column",
                          "high_point": True}},
)
```

Types: `line` (default), `column`, `win_lose`. Also takes `color`, `style` and
the toggles `high_point`, `low_point`, `first_point`, `last_point`, `markers`,
`negative_points`, `axis`, `right_to_left`.

The target column has to be one that already exists — appending one would mean
reaching into the header assembly and column accounting that `formula_columns`
uses, in both row loops, which costs far more than leaving a key out of your
records does.

## Notes and Images

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
