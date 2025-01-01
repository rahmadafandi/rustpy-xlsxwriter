from typing import Dict, List, Optional

def get_version() -> str:
    pass

def save_records(
    records: List[Dict[str, str]],
    file_name: str,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
):
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
    records_with_sheet_name: List[Dict[str, List[Dict[str, str]]]],
    file_name: str,
    password: Optional[str] = None,
):
    pass

def get_name() -> str:
    pass

def get_authors() -> str:
    pass

def get_description() -> str:
    pass

def get_repository() -> str:
    pass

def get_homepage() -> str:
    pass

def get_license() -> str:
    pass

def validate_sheet_name(name: str) -> bool:
    pass
