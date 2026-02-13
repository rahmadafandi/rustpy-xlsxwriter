"""Protect an Excel file with a password."""

from rustpy_xlsxwriter import FastExcel

records = [
    {"Account": "Savings", "Balance": 12500.50},
    {"Account": "Checking", "Balance": 3200.75},
]

(
    FastExcel("output_protected.xlsx", password="s3cret!")
    .format(float_format="0.00")
    .freeze(row=1)
    .sheet("Accounts", records)
    .save()
)
print("✅ output_protected.xlsx created (password protected)")
