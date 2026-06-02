"""Cell-level formatting via the Format class (issue #18)."""

from rustpy_xlsxwriter import FastExcel, Format

records = [
    {"product": "Widget", "price": 19.99, "qty": 100},
    {"product": "Gadget", "price": 49.50, "qty": 25},
]

money = Format().set_num_format("$#,##0.00").set_font_color("#006600")
header = Format().set_bold().set_background_color("#1F4E78").set_font_color("white")
qty = Format().set_align("center").set_border("thin")

(
    FastExcel("output_16_cell_formats.xlsx", autofit=True)
    .sheet(
        "Products",
        records,
        header_format=header,
        column_formats={"price": money, "qty": qty},
    )
    .save()
)

print("Wrote output_16_cell_formats.xlsx")
