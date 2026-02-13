"""Write a Pandas DataFrame to Excel."""

import pandas as pd

from rustpy_xlsxwriter import FastExcel

df = pd.DataFrame(
    {
        "Name": ["Alice", "Bob", "Charlie"],
        "Score": [95.678, 88.123, 72.456],
        "Passed": [True, True, False],
    }
)

# Basic
FastExcel("output_dataframe.xlsx").sheet("Results", df).save()
print("✅ output_dataframe.xlsx created")

# With float formatting and bold index columns
(
    FastExcel("output_dataframe_styled.xlsx")
    .format(float_format="0.00", index_columns=["Name"])
    .sheet("Results", df)
    .save()
)
print("✅ output_dataframe_styled.xlsx created (float_format + bold index)")
