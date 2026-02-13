"""Write multiple sheets in a single file."""

from rustpy_xlsxwriter import FastExcel

employees = [
    {"Name": "Alice", "Age": 30, "Department": "Engineering"},
    {"Name": "Bob", "Age": 25, "Department": "Marketing"},
]

inventory = [
    {"Product": "Laptop", "Price": 1200.00, "Stock": 50},
    {"Product": "Phone", "Price": 800.00, "Stock": 120},
    {"Product": "Tablet", "Price": 450.00, "Stock": 80},
]

(
    FastExcel("output_multi.xlsx")
    .sheet("Employees", employees)
    .sheet("Inventory", inventory)
    .save()
)
print("✅ output_multi.xlsx created with 2 sheets")
