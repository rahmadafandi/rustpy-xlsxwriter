# CSV and TSV

## CSV / TSV Output

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

## Large files

Writing to a path streams: the rows go to disk in 64 KiB pieces as they are
produced, so a multi-gigabyte export needs no more memory than a small one.
A list of dicts is of course already in memory; pass a generator to keep
the input flat as well.

```python
write_csv((row_to_dict(r) for r in cursor), "export.csv")
```

If the write fails part-way, for example on a record with a key the header
does not have, the partial file is removed rather than left looking like a
complete export with fewer rows. A failure in the first 64 KiB never opens
the file, so whatever was at the path before is still there.

A buffer such as `io.BytesIO` gets the whole file in one `write()` call, so
the output is held in memory before it is handed over.

## Excel on Windows and the byte order mark

A UTF-8 CSV without a BOM opens in Excel as the system code page, which turns
every non-ASCII character into mojibake. `bom=True` fixes it:

```python
write_csv(records, "out.csv", bom=True)      # "Café" stays "Café" in Excel
```

Off by default, so output stays byte-identical for anything that parses it.

## Selecting columns

```python
write_csv(records, "out.csv", columns=["sku", "name"])  # subset and order
write_csv(records, "out.csv", header=False)             # no header row
```

`columns` works on every input type and stays zero-copy on the Arrow path, so
you no longer have to slice a DataFrame in Python first — which is what threw
away the speed this library is for. A name that is not in the data raises
`ValueError`: unlike the styling options this one decides the shape of the
file, so a silent drop would hand back something that looks complete. The
Excel writers take the same option; see
[Data sources](dataframes.md#selecting-columns).
