"""Basic usage: write a single sheet from a list of dicts."""

from rustpy_xlsxwriter import FastExcel

records = [
    {"Name": "Alice", "Age": 30, "City": "New York"},
    {"Name": "Bob", "Age": 25, "City": "San Francisco"},
    {"Name": "Charlie", "Age": 35, "City": "Chicago"},
]

FastExcel("output_basic.xlsx").sheet("Employees", records).save()
print("✅ output_basic.xlsx created")
