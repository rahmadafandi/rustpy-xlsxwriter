import io

from rustpy_xlsxwriter import FastExcel

XLSX_MAGIC = b"PK\x03\x04"


def test_single_sheet_bytesio():
    buf = io.BytesIO()
    FastExcel(buf).sheet("Sheet1", [{"Name": "Alice", "Age": 30}]).save()
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC


def test_multi_sheet_bytesio():
    buf = io.BytesIO()
    (
        FastExcel(buf)
        .sheet("Sheet1", [{"Name": "Alice"}])
        .sheet("Sheet2", [{"Item": "Laptop"}])
        .save()
    )
    buf.seek(0)
    assert buf.read(4) == XLSX_MAGIC
