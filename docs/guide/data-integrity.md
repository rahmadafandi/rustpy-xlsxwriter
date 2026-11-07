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

## Decimals: exact in CSV, floats in xlsx

Excel has no decimal type: every number is a 64-bit float. A Python `Decimal`,
and a decimal column in a DataFrame, are written to xlsx as the nearest float.
Trailing zeros go, and so does every digit past the 15th or 16th significant
one:

```python
write_worksheet([{"price": Decimal("1.10")}], "out.xlsx")   # cell holds 1.1
```

For money that is usually fine, provided the cell is formatted to show the
cents: `column_formats={"price": Format().set_num_format("0.00")}`. When the
exact digits matter in xlsx, pass the text instead: `str(value)`.

CSV is text, so nothing has to be lost there. A `Decimal` is written as its
own text, which `Decimal(text)` reads back exactly, trailing zeros and
exponent included. A DataFrame decimal column is written with as many
fractional digits as its scale:

```python
write_csv([{"price": Decimal("1.10")}], "out.csv")            # price\n1.10
write_csv([{"n": Decimal("12345678901234567890.12")}], "out.csv")
# n\n12345678901234567890.12
```

A `Decimal` NaN or infinity takes `na_rep` and `inf_value`, as a float does.
Before 0.9 CSV got the nearest float too, so `1.10` came out as `1.1`; a
pipeline that compared the text will see the trailing zeros now.

## Dates, times and timezones

**Fractions of a second** are kept. Excel stores a datetime as a day count
with the time as the fraction, which holds about a microsecond; its number
formats show at most milliseconds, so ask for them with `hh:mm:ss.000`. CSV
writes `.ffffff`, the way `isoformat()` does, and nanoseconds from an Arrow
column as `.fffffffff`.

**A `datetime.time`** is written as an Excel time, formatted `hh:mm:ss`,
so it can be formatted, summed and compared like any time. CSV writes
`HH:MM:SS`. A `column_formats` entry overrides the time format.

**Timezones** survive in CSV and not in xlsx. Excel has no timezone, so
something has to give:

| Value | xlsx | CSV |
|---|---|---|
| naive `datetime` | as given | as given, no offset |
| aware Python `datetime` | its wall-clock time; the offset is dropped | wall-clock time and offset, as `isoformat()` writes it |
| aware DataFrame column | the **UTC** time | the UTC time with `+00:00` |

So in xlsx, `12:00+00:00` and `12:00+07:00` both become `12:00`, two instants
seven hours apart written as one value. And a polars column converted with
`dt.convert_time_zone("Asia/Jakarta")`, which shows 19:00, is written as
12:00, because a DataFrame stores the UTC instant and only names its zone.
CSV keeps the instant in both cases: `datetime.fromisoformat` reads back the
value you wrote.

To control what lands in xlsx, convert first:

```python
# Python: everything in UTC, or everything in one zone
rows = [{**r, "at": r["at"].astimezone(timezone.utc)} for r in rows]

# polars: keep the wall-clock time the column shows
df = df.with_columns(pl.col("at").dt.replace_time_zone(None))

# pandas: the same
df["at"] = df["at"].dt.tz_localize(None)
```

**Dates before 1900** cannot be stored in xlsx; Excel's calendar starts on
1900-01-01 and ends on 9999-12-31. A date outside it raises `ValueError`, from
a list of dicts and from a DataFrame alike; the DataFrame error names the
column and the value. Earlier versions wrote such a DataFrame cell empty
without an error. To keep old dates in xlsx, write the column as text. CSV
has no such limit and writes them as they are.

## Column types a DataFrame can hold

A DataFrame column whose type has no cell to go in raises `ValueError`
naming the column, instead of writing it empty:

| Column type | Written as |
|---|---|
| pandas `category`, polars `Categorical` and `Enum` | their values |
| time of day | an Excel time / `HH:MM:SS` |
| decimal | the nearest float in xlsx, the exact digits in CSV |
| an all-null column | empty cells |
| duration (`timedelta`) | `ValueError`: cast it to a string, or to a number of seconds |
| binary, list, struct, and anything else | `ValueError` |

Earlier versions wrote all of these, categories and times included, as empty
cells and reported success. A duration is refused rather than guessed, because
a guess at the unit would be silently wrong.

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
