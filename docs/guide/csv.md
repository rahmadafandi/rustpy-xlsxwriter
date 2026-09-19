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
file, so a silent drop would hand back something that looks complete.
