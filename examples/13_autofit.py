"""Control column auto-fit behavior.

Columns keep Excel's default width unless you ask for auto-fit.
Leave it off for large datasets; sizing scans every cell.
"""

from rustpy_xlsxwriter import FastExcel


def generate_rows(n: int):
    for i in range(n):
        yield {"id": i, "value": f"row_{i}", "score": i * 0.1}


# Opt in: columns adjust to content width
FastExcel("output_autofit_on.xlsx", autofit=True).sheet(
    "Data", generate_rows(100)
).save()
print("output_autofit_on.xlsx (autofit enabled, columns fit content)")

# Default: Excel's own widths, and nothing scans the data
FastExcel("output_autofit_off.xlsx").sheet("Data", generate_rows(100_000)).save()
print("output_autofit_off.xlsx (default, faster for large data)")
