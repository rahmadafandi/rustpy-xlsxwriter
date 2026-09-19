# Formatting

## Cell formatting

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

## Column widths

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

## Freeze Panes

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

## Row Layout: Merged Headers, Borders, Banding

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

## Conditional Formatting

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
