"""Full-featured example combining all options."""

from datetime import datetime

from rustpy_xlsxwriter import FastExcel

employees = [
    {
        "ID": 1,
        "Name": "Alice",
        "Salary": 95000.50,
        "Hired": datetime(2021, 3, 15),
        "Active": True,
    },
    {
        "ID": 2,
        "Name": "Bob",
        "Salary": 82000.75,
        "Hired": datetime(2022, 7, 1),
        "Active": True,
    },
    {
        "ID": 3,
        "Name": "Charlie",
        "Salary": 110000.00,
        "Hired": datetime(2020, 1, 10),
        "Active": False,
    },
]

departments = [
    {"Dept": "Engineering", "Budget": 500000.00, "Headcount": 25},
    {"Dept": "Marketing", "Budget": 200000.00, "Headcount": 12},
    {"Dept": "Sales", "Budget": 350000.00, "Headcount": 18},
]

(
    FastExcel("output_full.xlsx", password="admin123")
    .format(float_format="0.00", index_columns=["ID", "Dept"])
    .freeze(row=1)
    .freeze(row=1, col=2, sheet="Employees")  # extra freeze for Employees
    .sheet("Employees", employees)
    .sheet("Departments", departments)
    .save()
)
print("✅ output_full.xlsx created")
print("   - 2 sheets: Employees, Departments")
print("   - Password protected")
print("   - Float format: 0.00")
print("   - Bold index columns: ID, Dept")
print("   - Frozen header row on all sheets")
print("   - Extra column freeze on Employees sheet")
