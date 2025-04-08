# RustPy-XlsxWriter

[![PyPI version](https://badge.fury.io/py/rustpy-xlsxwriter.svg)](https://badge.fury.io/py/rustpy-xlsxwriter)
[![Python Versions](https://img.shields.io/pypi/pyversions/rustpy-xlsxwriter.svg)](https://pypi.org/project/rustpy-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://pepy.tech/badge/rustpy-xlsxwriter)](https://pepy.tech/project/rustpy-xlsxwriter)
[![Tests](https://github.com/rahmadafandi/rustpy-xlsxwriter/workflows/Tests/badge.svg)](https://github.com/rahmadafandi/rustpy-xlsxwriter/actions)

RustPy-XlsxWriter is a high-performance library for generating Excel files in Python, powered by Rust and integrated using PyO3. This library is ideal for creating Excel files with large datasets efficiently while maintaining a simple and Pythonic interface.

## Installation

Install RustPy-XlsxWriter via pip:

```bash
pip install rustpy-xlsxwriter
```

## Features

- Create Excel files quickly and efficiently.
- Support for various data types including text, numbers, dates, and booleans.
- Save data into multiple sheets.
- Optionally protect Excel files with passwords.
- Sheet name validation utilities.
- Freeze panes support for better worksheet navigation.

## Freeze Panes

RustPy-XlsxWriter supports freezing rows and columns in worksheets to keep important data visible while scrolling. This feature is available for both single and multiple worksheet scenarios.

### Single Worksheet

For a single worksheet, you can freeze rows and columns using the `freeze_row` and `freeze_col` parameters:

```python
from rustpy_xlsxwriter import write_worksheet

# Freeze the first row (headers)
write_worksheet(records, "output.xlsx", freeze_row=1)

# Freeze both first row and first column
write_worksheet(records, "output.xlsx", freeze_row=1, freeze_col=1)
```

### Multiple Worksheets

For multiple worksheets, you can configure freeze panes globally and/or per sheet using the `freeze_pane` parameter:

```python
from rustpy_xlsxwriter import write_worksheets

# Configuration for freeze panes
freeze_config = {
    "general": {"row": 1, "col": 0},  # Apply to all sheets
    "Sheet1": {"row": 1, "col": 2},  # Override for Sheet1
    "Sheet2": {"row": 2, "col": 1}   # Override for Sheet2
}

write_worksheets(
    records_with_sheet_name,
    "output.xlsx",
    freeze_panes=freeze_config
)
```

The `freeze_panes` configuration allows you to:

- Set a general configuration that applies to all sheets
- Override the configuration for specific sheets
- Freeze rows and columns independently

## API Reference

Below is the API provided by `rustpy_xlsxwriter`:

### `get_version()`

```python
from rustpy_xlsxwriter import get_version

def get_version() -> str:
    """
    Get the version of the RustPy-XlsxWriter library.

    Returns:
        str: The version string.
    """
```

### `write_worksheet()`

```python
from rustpy_xlsxwriter import write_worksheet

def write_worksheet(
    records: List[Dict[str, Any]],
    file_name: str,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
    freeze_row: Optional[int] = None,
    freeze_col: Optional[int] = None,
):
    """
    Save records to a single sheet in an Excel file.

    Args:
        records: List of dictionaries where each dict represents a row of data.
                Dictionary keys become column headers and values become cell contents.
                Supported value types:
                - str: Text values
                - int/float: Numeric values
                - bool: Boolean values
                - None: Empty cells
                - datetime.date/datetime.datetime: Date values
        file_name: Full path including filename where the Excel file will be saved.
                  Must have .xlsx extension.
        sheet_name: Optional name for the worksheet. If not provided, defaults to 'Sheet1'.
                   Must be <= 31 chars and cannot contain [ ] : * ? / \.
        password: Optional password to protect the workbook from modifications.
        freeze_row: Optional row index to freeze.
        freeze_col: Optional column index to freeze.
    """
```

### `write_worksheets()`

```python
from rustpy_xlsxwriter import write_worksheets

def write_worksheets(
    records_with_sheet_name: List[Dict[str, List[Dict[str, Any]]]],
    file_name: str,
    password: Optional[str] = None,
    freeze_panes: Optional[Dict[str, Any]] = None,
):
    """
    Save records to multiple sheets in an Excel file.

    Args:
        records_with_sheet_name: List of dictionaries where each dict maps a sheet name to its records.
                                The records for each sheet follow the same format as write_worksheet().
                                Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \.
        file_name: Full path including filename where the Excel file will be saved.
                  Must have .xlsx extension.
        password: Optional password to protect the workbook from modifications.
        freeze_panes: Optional configuration for freeze panes.
                      If provided, it should be a dictionary with the following keys:
                      - 'general': A dictionary with 'row' and 'col' keys for general freeze panes.
                      - 'Sheet Name': A dictionary where keys are sheet names and values are dictionaries
                                          with 'row' and 'col' keys for specific sheet freeze panes.
    """
```

### `validate_sheet_name()`

```python
from rustpy_xlsxwriter import validate_sheet_name

def validate_sheet_name(name: str) -> bool:
    """
    Validate if a sheet name is valid for Excel.

    Args:
        name: Sheet name to validate. Excel has several restrictions on valid sheet names:
              - Maximum 31 characters
              - Cannot contain characters: [ ] : * ? / \
              - Cannot be empty
              - Cannot start or end with an apostrophe
              - Cannot be 'History' (reserved name)

    Returns:
        bool: True if the sheet name is valid for Excel, False otherwise
    """
```

## Performance

![Test Result](image.png)

RustPy-XlsxWriter has been extensively tested with large-scale datasets to measure its performance capabilities. Our benchmarks demonstrate that this Rust-powered implementation delivers exceptional speed improvements compared to traditional Python solutions. The library achieves up to 6x faster processing speeds while maintaining optimal memory usage, making it ideal for handling large datasets efficiently.

Based on performance testing with 1 million records:

| Operation         | Records   | Time (seconds) | Comparison  |
| ----------------- | --------- | -------------- | ----------- |
| Single Sheet      | 1,000,000 | ~67.80s        | 5.4x faster |
| Multiple Sheets   | 1,000,000 | ~61.19s        | 6x faster   |
| Python xlsxwriter | 1,000,000 | ~364.46s       | baseline    |

Key findings:

- Demonstrates superior performance with 6x faster processing compared to Python's xlsxwriter
- Efficiently handles single sheet operations for 1 million records
- Maintains consistent performance for multiple sheet operations
- Shows excellent scalability - performance improves proportionally with smaller datasets

The exceptional performance is achieved through several key optimizations:

1. Leveraging Rust's zero-cost abstractions and memory management system
2. Native machine code compilation for maximum efficiency
3. Advanced memory optimization using rust_xlsxwriter capabilities
4. High-precision floating point operations with ryu
5. Efficient large file handling through zlib compression
6. Memory safety guarantees via Rust's ownership system

These technical advantages ensure consistent high performance and reliability across varying workload sizes while maintaining optimal resource utilization.

## Usage Examples

### Write Records to a Single Sheet

```python
from rustpy_xlsxwriter import write_worksheet
from datetime import datetime

records = [
    {
        "Name": "Alice",
        "Age": 30,
        "City": "New York",
        "Active": True,
        "Join Date": datetime(2023, 1, 15)
    },
    {
        "Name": "Bob",
        "Age": 25,
        "City": "San Francisco",
        "Active": False,
        "Join Date": datetime(2023, 2, 1)
    },
]

# Basic usage
write_worksheet(records, "output.xlsx", sheet_name="Sheet1")

# With freeze panes - freeze first row and first column
write_worksheet(records, "output_frozen.xlsx", sheet_name="Sheet1", freeze_row=1, freeze_col=1)
```

### Write Records to Multiple Sheets

```python
from rustpy_xlsxwriter import write_worksheets

records_with_sheet_name = [
    {"Employees": [
        {
            "Name": "Alice",
            "Age": 30,
            "City": "New York",
            "Active": True
        },
        {
            "Name": "Bob",
            "Age": 25,
            "City": "San Francisco",
            "Active": False
        },
    ]},
    {"Inventory": [
        {
            "Product": "Laptop",
            "Price": 1000.0,
            "Stock": 50,
            "Available": True
        },
        {
            "Product": "Phone",
            "Price": 500.0,
            "Stock": 100,
            "Available": True
        },
    ]},
]

# Basic usage
write_worksheets(records_with_sheet_name, "output_multiple_sheets.xlsx")

# With freeze panes configuration
freeze_config = {
    "general": {"row": 1, "col": 0},  # Apply to all sheets
    "Employees": {"row": 1, "col": 2},  # Override for Employees sheet
    "Inventory": {"row": 2, "col": 1}   # Override for Inventory sheet
}
write_worksheets(
    records_with_sheet_name,
    "output_frozen.xlsx",
    freeze_pane=freeze_config
)
```

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## License

This project is licensed under the MIT ![License](LICENSE).

## Acknowledgements

This project is inspired by [Rust-XlsxWriter](https://github.com/jmcnamara/rust_xlsxwriter) and [PyO3](https://github.com/pyo3/pyo3) with the help of [maturin](https://github.com/PyO3/maturin).

## Contributors

- [Rahmad Afandi](https://github.com/rahmadafandi)
