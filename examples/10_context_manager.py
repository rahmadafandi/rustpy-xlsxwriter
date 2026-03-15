"""Use FastExcel as a context manager for auto-save.

The file is automatically saved when exiting the `with` block.
If an exception occurs, the file is NOT saved.
"""

from rustpy_xlsxwriter import FastExcel

records = [
    {"Name": "Alice", "Age": 30},
    {"Name": "Bob", "Age": 25},
]

# Auto-saves on exit
with FastExcel("output_context.xlsx") as f:
    f.format(bold_headers=True)
    f.freeze(row=1)
    f.sheet("Users", records)

print("✅ output_context.xlsx created (auto-saved via context manager)")

# Exception safety: file is NOT saved if an error occurs
try:
    with FastExcel("output_should_not_exist.xlsx") as f:
        f.sheet("Data", records)
        raise ValueError("Something went wrong!")
except ValueError:
    pass

import os

assert not os.path.exists("output_should_not_exist.xlsx")
print("✅ Confirmed: file not created when exception occurred")
