from typing import Dict, List, Optional, Any

def get_version() -> str:
    """Get the version of the rustpy_xlsxwriter package.

    Returns:
        str: The version string (e.g. '0.0.3')
    """
    pass

def save_records(
    records: List[Dict[str, Any]],
    file_name: str,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
):
    """Save records to an Excel file.

    Args:
        records: List of dictionaries where each dict represents a row of data.
                Dictionary keys become column headers and values become cell contents.
                Values can be strings, numbers, None, or empty strings.
        file_name: Full path including filename where the Excel file will be saved.
                  Must have .xlsx extension.
        sheet_name: Optional name for the worksheet. If not provided, defaults to 'Sheet1'.
                   Must be <= 31 chars and cannot contain [ ] : * ? / \\.
        password: Optional password to protect the workbook from modifications.

    Raises:
        ValueError: If file_name doesn't end with .xlsx
        ValueError: If sheet_name is invalid according to Excel's requirements
    """
    pass

# TODO: Add this back in when we have a better solution
# def save_records_multithread(
#     records: List[Dict[str, str]],
#     file_name: str,
#     sheet_name: Optional[str] = None,
#     password: Optional[str] = None,
# ):
#     pass

def save_records_multiple_sheets(
    records_with_sheet_name: List[Dict[str, List[Dict[str, Any]]]],
    file_name: str,
    password: Optional[str] = None,
):
    """Save records to multiple sheets in an Excel file.

    Args:
        records_with_sheet_name: List of dictionaries where each dict maps a sheet name to its records.
                                The records for each sheet follow the same format as save_records().
                                Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\.
        file_name: Full path including filename where the Excel file will be saved.
                  Must have .xlsx extension.
        password: Optional password to protect the workbook from modifications.

    Raises:
        ValueError: If file_name doesn't end with .xlsx
        ValueError: If any sheet name is invalid according to Excel's requirements
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
        
    Returns:
        bool: True if the sheet name is valid for Excel, False otherwise
    """
    pass
