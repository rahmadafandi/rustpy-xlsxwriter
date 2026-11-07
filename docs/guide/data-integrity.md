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

## Records must carry the same keys

The header comes from the first record, because a generator gives no way to
look ahead. Later records are matched to it **by key name**, so one whose keys
arrive in another order still lands in the right columns:

```python
rows = [{"id": 1, "name": "Alice"}, {"name": "Bob", "id": 2}]
# Bob stays under "name", 2 stays under "id".
```

A record missing a key leaves that cell empty. A record carrying a key the
first one did not have raises, because the header row has already been written
and a column cannot be given a title after the fact:

```python
FastExcel("out.xlsx").sheet("S", [{"a": 1}, {"a": 2, "b": 3}]).save()
# ValueError: record 1: no column 'b'. The header comes from the first record (a) …
```

Normalise the records first, or hand over a DataFrame, which carries its own
schema and has no first-record problem.

## Whole numbers above 2^53

Excel stores every number as a 64-bit float, so an integer larger than
9,007,199,254,740,992 loses its last digits, and nothing reports it: the file
is valid either way:

```python
write_worksheet([{"id": 9007199254740993}], "out.xlsx")
# reads back as 9007199254740992
```

This is the file format, not this library; `xlsxwriter` and Excel itself
behave the same. Long identifiers, account numbers and phone numbers belong in
a text column, where they also keep any leading zeros:

```python
write_worksheet([{"id": str(9007199254740993)}], "out.xlsx")
```

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
