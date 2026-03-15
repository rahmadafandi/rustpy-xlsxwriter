"""Control column auto-fit behavior.

By default, columns are auto-fitted to content width.
Disable autofit for large datasets to improve write performance.
"""

from rustpy_xlsxwriter import FastExcel


def generate_rows(n: int):
    for i in range(n):
        yield {"id": i, "value": f"row_{i}", "score": i * 0.1}


# Default: autofit enabled (columns adjust to content width)
FastExcel("output_autofit_on.xlsx").sheet("Data", generate_rows(100)).save()
print("✅ output_autofit_on.xlsx (autofit enabled — columns fit content)")

# Disabled: faster for large datasets
FastExcel("output_autofit_off.xlsx", autofit=False).sheet(
    "Data", generate_rows(100_000)
).save()
print("✅ output_autofit_off.xlsx (autofit disabled — faster for large data)")
