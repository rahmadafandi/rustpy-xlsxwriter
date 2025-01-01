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

def save_records_multi_thread(
    records: List[Dict[str, str]],
    file_name: str,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
):
    pass

def save_records_multiple_sheets(
    records_with_sheet_name: List[Dict[str, List[Dict[str, str]]]],
    file_name: str,
    password: Optional[str] = None,
):
    pass
