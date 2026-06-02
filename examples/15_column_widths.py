"""Per-column and uniform column widths (issue #17)."""

from rustpy_xlsxwriter import FastExcel, write_worksheets

meta = [
    {"row": 1, "var": "age", "label": "Age in years", "type": "int", "format": "0"},
    {"row": 2, "var": "name", "label": "Full legal name", "type": "str", "format": "@"},
]
raw = [{"a": 1, "b": 2, "c": 3}, {"a": 4, "b": 5, "c": 6}]

# Builder: uniform on one sheet, per-column (by name) on another.
(
    FastExcel("output_15_column_widths.xlsx", autofit=False)
    .sheet("RawData", raw, column_width=15)
    .sheet(
        "Meta",
        meta,
        column_widths={"row": 7, "var": 22, "label": 91, "type": 14, "format": 12},
    )
    .save()
)

# Functional multi-sheet form, keyed by sheet name with a "general" fallback.
write_worksheets(
    [("RawData", raw), ("Meta", meta)],
    "output_15_column_widths_functional.xlsx",
    autofit=False,
    column_width={"general": 15},
    column_widths={"Meta": [7, 22, 91, 14, 12]},
)

print("Wrote output_15_column_widths*.xlsx")
