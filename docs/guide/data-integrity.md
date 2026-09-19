# Data integrity

## Data Validation

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

## Missing values, NaN and infinity

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
