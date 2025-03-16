# RustPy-XlsxWriter

RustPy-XlsxWriter is a high-performance library for generating Excel files in Python, powered by Rust and integrated using PyO3. This library is ideal for creating Excel files with large datasets efficiently while maintaining a simple and Pythonic interface.

## Installation

Install RustPy-XlsxWriter via pip:

```bash
pip install rustpy-xlsxwriter
```

## Features

- Create Excel files quickly and efficiently.
- Support for single and multi-threaded record saving.
- Save data into multiple sheets.
- Optionally protect Excel files with passwords.

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

### `save_records()`

```python
from rustpy_xlsxwriter import save_records

def save_records(
    records: List[Dict[str, str]],
    file_name: str,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
):
    """
    Save records to a single sheet in an Excel file.

    Args:
        records (List[Dict[str, str]]): A list of dictionaries containing data to save.
        file_name (str): The name of the Excel file to create.
        sheet_name (Optional[str], optional): The name of the sheet. Defaults to None.
        password (Optional[str], optional): The password to protect the file. Defaults to None.
    """
```

### `save_records_multiple_sheets()`

```python
from rustpy_xlsxwriter import save_records_multiple_sheets

def save_records_multiple_sheets(
    records_with_sheet_name: List[Dict[str, List[Dict[str, str]]]],
    file_name: str,
    password: Optional[str] = None,
):
    """
    Save records to multiple sheets in an Excel file.

    Args:
        records_with_sheet_name (List[Dict[str, List[Dict[str, str]]]]): A list of dictionaries with sheet names as keys and record lists as values.
        file_name (str): The name of the Excel file to create.
        password (Optional[str], optional): The password to protect the file. Defaults to None.
    """
```
## Performance
![Test Result](image.png)

The library has undergone rigorous performance testing with large-scale datasets to evaluate its speed and efficiency. Benchmarks show that this Rust-based implementation consistently outperforms the pure Python xlsxwriter library by a significant margin, delivering up to 6.6x faster processing speeds while maintaining memory efficiency. The performance metrics below demonstrate these substantial improvements.

Based on performance testing with 1 million records:

| Operation | Records | Time (seconds) |
|-----------|---------|----------------|
| Single Sheet | 1,000,000 | ~60.38s |
| Multiple Sheets | 1,000,000 | ~64.13s |
| Python xlsxwriter | 1,000,000 | ~398.43s |

Key findings:
- The Rust implementation processes data approximately 6.6x faster than Python's xlsxwriter library
- Single sheet operations complete in around 60 seconds for 1 million records
- Multiple sheet operations take slightly longer at ~64 seconds for the same volume
- Performance scales linearly with data size - smaller datasets process proportionally faster

The significant performance advantages are achieved through:

1. Rust's zero-cost abstractions and efficient memory management
2. Direct compilation to optimized native machine code
3. Constant memory usage optimization via rust_xlsxwriter features
4. Enhanced floating point handling with the ryu feature
5. Optimized large file handling using zlib compression
6. Rust's ownership model preventing memory leaks and race conditions

These optimizations ensure excellent performance across different dataset sizes while maintaining memory efficiency.


## Usage Examples

### Save Records to a Single Sheet

```python
from rustpy_xlsxwriter import save_records

records = [
    {"Name": "Alice", "Age": "30", "City": "New York"},
    {"Name": "Bob", "Age": "25", "City": "San Francisco"},
]

save_records(records, "output.xlsx", sheet_name="Sheet1")
```

### Save Records to Multiple Sheets

```python
from rustpy_xlsxwriter import save_records_multiple_sheets

records_with_sheet_name = [
    {"Sheet1": [
        {"Name": "Alice", "Age": "30", "City": "New York"},
        {"Name": "Bob", "Age": "25", "City": "San Francisco"},
    ]},
    {"Sheet2": [
        {"Product": "Laptop", "Price": "1000", "Stock": "50"},
        {"Product": "Phone", "Price": "500", "Stock": "100"},
    ]},
]

save_records_multiple_sheets(records_with_sheet_name, "output_multiple_sheets.xlsx")
```

## Contributing

Contributions are welcome! Please submit issues or pull requests on the [GitHub repository](https://github.com/rahmadafandi/rustpy-xlsxwriter).

## License

This project is licensed under the MIT ![License](LICENSE).

## Acknowledgements

This project is inspired by [Rust-XlsxWriter](https://github.com/jmcnamara/rust_xlsxwriter) and [PyO3](https://github.com/pyo3/pyo3) with the help of [maturin](https://github.com/PyO3/maturin).

## Contributors

- [Rahmad Afandi](https://github.com/rahmadafandi)