"""Make header row bold."""

from rustpy_xlsxwriter import FastExcel

records = [
    {"Name": "Alice", "Score": 95.5, "Grade": "A"},
    {"Name": "Bob", "Score": 88.0, "Grade": "B+"},
    {"Name": "Charlie", "Score": 72.3, "Grade": "B-"},
]

(
    FastExcel("output_bold_headers.xlsx")
    .format(bold_headers=True)
    .sheet("Results", records)
    .save()
)
print("✅ output_bold_headers.xlsx created (bold header row)")
