from typing import Any, Dict, List, Optional, Union

def get_version() -> str:
    """Get the version of the rustpy_xlsxwriter package.

    Returns:
        str: The version string (e.g. '0.0.5')
    """
    pass

def write_worksheet(
    records: List[Dict[str, Any]],
    file_name: str,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
) -> None:
    """Save records to an Excel file.

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
                   Must be <= 31 chars and cannot contain [ ] : * ? / \\.
        password: Optional password to protect the workbook from modifications.

    Raises:
        ValueError: If file_name doesn't end with .xlsx
        ValueError: If sheet_name is invalid according to Excel's requirements
        ValueError: If records contain unsupported data types
        OSError: If there are filesystem errors when writing the file
    """
    pass

# TODO: Add this back in when we have a better solution
# def write_worksheet_multithread(
#     records: List[Dict[str, Any]],
#     file_name: str,
#     sheet_name: Optional[str] = None,
#     password: Optional[str] = None,
# ) -> None:
#     pass

def write_worksheets(
    records_with_sheet_name: List[Dict[str, List[Dict[str, Any]]]],
    file_name: str,
    password: Optional[str] = None,
) -> None:
    """Save records to multiple sheets in an Excel file.

    Args:
        records_with_sheet_name: List of dictionaries where each dict maps a sheet name to its records.
                                The records for each sheet follow the same format as write_worksheet().
                                Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\.
        file_name: Full path including filename where the Excel file will be saved.
                  Must have .xlsx extension.
        password: Optional password to protect the workbook from modifications.

    Raises:
        ValueError: If file_name doesn't end with .xlsx
        ValueError: If any sheet name is invalid according to Excel's requirements
        ValueError: If records contain unsupported data types
        OSError: If there are filesystem errors when writing the file
    """
    pass

def get_name() -> str:
    """Get the name of the rustpy_xlsxwriter package.

    Returns:
        str: The package name ('rustpy-xlsxwriter')
    """
    pass

def get_authors() -> str:
    """Get the authors of the rustpy_xlsxwriter package.

    Returns:
        str: The authors string ('Rahmad Afandi <rahmadafandiii@gmail.com>')
    """
    pass

def get_description() -> str:
    """Get the description of the rustpy_xlsxwriter package.

    Returns:
        str: The package description ('Rust Python bindings for rust_xlsxwriter')
    """
    pass

def get_repository() -> str:
    """Get the repository URL of the rustpy_xlsxwriter package.

    Returns:
        str: The repository URL ('https://github.com/rahmadafandi/rustpy-xlsxwriter')
    """
    pass

def get_homepage() -> str:
    """Get the homepage URL of the rustpy_xlsxwriter package.

    Returns:
        str: The homepage URL ('https://github.com/rahmadafandi/rustpy-xlsxwriter')
    """
    pass

def get_license() -> str:
    """Get the license of the rustpy_xlsxwriter package.

    Returns:
        str: The package license ('MIT')
    """
    pass

def validate_sheet_name(name: str) -> bool:
    """Validate if a sheet name is valid for Excel.
    
    Args:
        name: Sheet name to validate. Excel has several restrictions on valid sheet names:
              - Maximum 31 characters
              - Cannot contain characters: [ ] : * ? / \\
              - Cannot be empty
              - Cannot start or end with an apostrophe
              - Cannot be 'History' (reserved name)
        
    Returns:
        bool: True if the sheet name is valid for Excel, False otherwise

    Examples:
        >>> validate_sheet_name("Sheet1")  # Valid
        True
        >>> validate_sheet_name("Sheet[1]")  # Invalid - contains brackets
        False
        >>> validate_sheet_name("A"*32)  # Invalid - too long
        False
    """
    pass

__all__ = [
    "get_version",
    "write_worksheet",
    "write_worksheets",
    "get_name",
    "get_authors",
    "get_description",
    "get_repository",
    "get_homepage",
    "get_license",
    "validate_sheet_name",
]
