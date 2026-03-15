"""Write a Polars DataFrame to Excel (native support, no .to_pandas() needed)."""

import polars as pl

from rustpy_xlsxwriter import FastExcel

df = pl.DataFrame(
    {
        "Name": ["Alice", "Bob", "Charlie"],
        "Score": [95.678, 88.123, 72.456],
        "Passed": [True, True, False],
    }
)

# Basic
FastExcel("output_polars.xlsx").sheet("Results", df).save()
print("✅ output_polars.xlsx created")

# With styling
(
    FastExcel("output_polars_styled.xlsx")
    .format(float_format="0.00", bold_headers=True, index_columns=["Name"])
    .sheet("Results", df)
    .save()
)
print("✅ output_polars_styled.xlsx created (float_format + bold headers + bold index)")
