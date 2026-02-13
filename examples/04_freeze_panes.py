"""Freeze panes to keep headers visible while scrolling."""

from rustpy_xlsxwriter import FastExcel

records = [{"ID": i, "Name": f"Item {i}", "Value": i * 1.5} for i in range(100)]

# Freeze first row (headers stay visible)
FastExcel("output_freeze_row.xlsx").freeze(row=1).sheet("Data", records).save()
print("✅ output_freeze_row.xlsx created (frozen header row)")

# Freeze first row + first column
(
    FastExcel("output_freeze_both.xlsx")
    .freeze(row=1, col=1)
    .sheet("Data", records)
    .save()
)
print("✅ output_freeze_both.xlsx created (frozen row + column)")

# Per-sheet freeze config
(
    FastExcel("output_freeze_custom.xlsx")
    .freeze(row=1)  # general: all sheets
    .freeze(row=2, col=1, sheet="Summary")  # override for Summary
    .sheet("Data", records)
    .sheet("Summary", records[:10])
    .save()
)
print("✅ output_freeze_custom.xlsx created (per-sheet freeze config)")
