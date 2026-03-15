import io

import openpyxl
import pandas as pd
import pytest

from rustpy_xlsxwriter import FastExcel

from conftest import XLSX_MAGIC


class TestFloatFormat:
    def test_number_format_applied(self, tmp_path):
        records = [{"Value": 123.45678}]
        path = str(tmp_path / "float_fmt.xlsx")
        FastExcel(path).format(float_format="0.00").sheet("Sheet1", records).save()

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        assert ws.cell(2, 1).value == pytest.approx(123.45678)
        assert ws.cell(2, 1).number_format == "0.00"
        wb.close()

    def test_with_bytesio(self):
        buf = io.BytesIO()
        df = pd.DataFrame({"Value": [123.45678]})
        FastExcel(buf).format(float_format="0.00").sheet("Sheet1", df).save()
        buf.seek(0)
        assert buf.read(4) == XLSX_MAGIC


class TestIndexColumns:
    def test_bold_applied_to_index_column(self, tmp_path):
        records = [{"ID": 1, "Name": "Alice"}, {"ID": 2, "Name": "Bob"}]
        path = str(tmp_path / "index_cols.xlsx")
        FastExcel(path).format(index_columns=["ID"]).sheet("Sheet1", records).save()

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        assert ws.cell(2, 1).font.bold is True
        assert ws.cell(2, 2).font.bold is not True
        wb.close()


class TestCombinedStyling:
    def test_float_format_index_columns_and_freeze(self, tmp_path):
        records = [{"ID": 1, "Score": 99.123}]
        path = str(tmp_path / "combined.xlsx")
        (
            FastExcel(path)
            .format(float_format="0.00", index_columns=["ID"])
            .freeze(row=1)
            .sheet("Data", records)
            .save()
        )

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        assert ws.freeze_panes == "A2"
        assert ws.cell(2, 1).font.bold is True
        assert ws.cell(2, 2).number_format == "0.00"
        wb.close()
