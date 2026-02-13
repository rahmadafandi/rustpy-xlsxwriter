
import io
from rustpy_xlsxwriter import write_worksheet, write_worksheets

def test_write_worksheet_bytesio():
    output = io.BytesIO()
    records = [{"Name": "Alice", "Age": 30}]
    write_worksheet(records, output, sheet_name="Sheet1")
    output.seek(0)
    assert output.read(4) == b"PK\x03\x04"  # Check for zip header

def test_write_worksheets_bytesio():
    output = io.BytesIO()
    records_with_sheet_name = [{"Sheet1": [{"Name": "Alice"}]}]
    write_worksheets(records_with_sheet_name, output)
    output.seek(0)
    assert output.read(4) == b"PK\x03\x04"
