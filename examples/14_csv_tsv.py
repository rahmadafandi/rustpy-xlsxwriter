"""Write CSV and TSV files — auto-detected from file extension.

Same FastExcel API, just change the extension. ~5x faster than Python's csv module.
"""

from rustpy_xlsxwriter import FastExcel, write_csv

records = [
    {"Name": "Alice", "Age": 30, "Score": 95.5},
    {"Name": "Bob", "Age": 25, "Score": 88.0},
    {"Name": "Charlie", "Age": 35, "Score": 72.3},
]

# Auto-detected from extension — same API as Excel
FastExcel("output.csv").sheet("Sheet1", records).save()
print("output.csv created")

# TSV (tab-separated)
FastExcel("output.tsv").sheet("Sheet1", records).save()
print("output.tsv created")

# Or use write_csv directly
write_csv(records, "output_direct.csv")
print("output_direct.csv created")

# Custom delimiter
write_csv(records, "output_semicolon.csv", delimiter=";")
print("output_semicolon.csv created (semicolon delimiter)")

# Works with generators too
def large_data():
    for i in range(100_000):
        yield {"id": i, "value": f"row_{i}"}

FastExcel("output_large.csv").sheet("Data", large_data()).save()
print("output_large.csv created (100K rows via generator)")
