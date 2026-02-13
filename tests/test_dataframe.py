import pandas as pd
from rustpy_xlsxwriter import write_worksheet, write_worksheets
import io

def test_write_worksheet_dataframe():
    output = io.BytesIO()
    df = pd.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [30, 25],
        "Score": [95.5, 88.0],
    })
    write_worksheet(df, output, sheet_name="Sheet1")
    output.seek(0)
    assert output.read(4) == b"PK\x03\x04"

def test_write_worksheets_dataframe():
    output = io.BytesIO()
    df1 = pd.DataFrame({"A": [1, 2]})
    records_with_sheet_name = [{"Sheet1": df1}]
    # Note: write_worksheets expects a list of dicts where keys are sheet names and values are WorksheetData (which can be DataFrame)
    write_worksheets(records_with_sheet_name, output)
    output.seek(0)
    assert output.read(4) == b"PK\x03\x04"
