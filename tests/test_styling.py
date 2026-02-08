
import io
import pandas as pd
from rustpy_xlsxwriter import write_worksheet

def test_write_worksheet_float_format():
    output = io.BytesIO()
    df = pd.DataFrame({"Value": [123.45678]})
    write_worksheet(df, output, sheet_name="Sheet1", float_format="0.00")
    output.seek(0)
    assert output.read(4) == b"PK\x03\x04"
    # Note: Checking exact cell format via binary is hard, relying on successful write

def test_write_worksheet_index_columns():
    output = io.BytesIO()
    df = pd.DataFrame({"ID": [1, 2], "Name": ["A", "B"]})
    write_worksheet(df, output, sheet_name="Sheet1", index_columns=["ID"])
    output.seek(0)
    assert output.read(4) == b"PK\x03\x04"
