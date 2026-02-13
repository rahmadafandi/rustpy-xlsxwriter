import io

import pandas as pd

from rustpy_xlsxwriter import FastExcel

XLSX_MAGIC = b"PK\x03\x04"


def test_dataframe_single_sheet():
    buf = io.BytesIO()
    df = pd.DataFrame(
        {"Name": ["Alice", "Bob"], "Age": [30, 25], "Score": [95.5, 88.0]}
    )
    FastExcel(buf).sheet("Sheet1", df).save()
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC


def test_dataframe_multi_sheet():
    buf = io.BytesIO()
    df1 = pd.DataFrame({"A": [1, 2]})
    df2 = pd.DataFrame({"B": [3, 4]})
    FastExcel(buf).sheet("Sheet1", df1).sheet("Sheet2", df2).save()
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC
