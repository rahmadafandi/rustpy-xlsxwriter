import io

import pandas as pd

from rustpy_xlsxwriter import FastExcel

XLSX_MAGIC = b"PK\x03\x04"


def test_float_format():
    buf = io.BytesIO()
    df = pd.DataFrame({"Value": [123.45678]})
    FastExcel(buf).format(float_format="0.00").sheet("Sheet1", df).save()
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC


def test_index_columns():
    buf = io.BytesIO()
    df = pd.DataFrame({"ID": [1, 2], "Name": ["A", "B"]})
    FastExcel(buf).format(index_columns=["ID"]).sheet("Sheet1", df).save()
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC


def test_combined_styling():
    buf = io.BytesIO()
    df = pd.DataFrame({"ID": [1], "Score": [99.123]})
    (
        FastExcel(buf)
        .format(float_format="0.00", index_columns=["ID"])
        .freeze(row=1)
        .sheet("Data", df)
        .save()
    )
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC
