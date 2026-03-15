"""Customize datetime format in Excel output."""

from datetime import date, datetime

from rustpy_xlsxwriter import FastExcel

records = [
    {"Event": "Launch", "When": datetime(2024, 6, 15, 10, 30, 0)},
    {"Event": "Review", "When": datetime(2024, 7, 1, 14, 0, 0)},
    {"Event": "Release", "When": datetime(2024, 8, 20, 9, 15, 0)},
]

# Default format: yyyy-mm-ddThh:mm:ss
FastExcel("output_datetime_default.xlsx").sheet("Events", records).save()
print("✅ output_datetime_default.xlsx (default: yyyy-mm-ddThh:mm:ss)")

# Custom format: dd/mm/yyyy
(
    FastExcel("output_datetime_custom.xlsx")
    .format(datetime_format="dd/mm/yyyy hh:mm")
    .sheet("Events", records)
    .save()
)
print("✅ output_datetime_custom.xlsx (custom: dd/mm/yyyy hh:mm)")

# Date objects (not datetime)
date_records = [
    {"Name": "Alice", "Birthday": date(1994, 3, 20)},
    {"Name": "Bob", "Birthday": date(1999, 7, 10)},
]

(
    FastExcel("output_dates.xlsx")
    .format(datetime_format="dd-mmm-yyyy")
    .sheet("Birthdays", date_records)
    .save()
)
print("✅ output_dates.xlsx (date objects with dd-mmm-yyyy format)")
