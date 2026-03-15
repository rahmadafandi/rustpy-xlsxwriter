import io
from datetime import date, datetime

import pandas as pd
import polars as pl
import pytest

from rustpy_xlsxwriter import FastExcel, write_csv


class TestCSVAutoDetect:
    def test_records_csv(self, tmp_path):
        path = str(tmp_path / "out.csv")
        records = [{"Name": "Alice", "Age": 30}, {"Name": "Bob", "Age": 25}]
        FastExcel(path).sheet("S", records).save()

        content = open(path).read()
        lines = content.strip().split("\n")
        assert lines[0] == "Name,Age"
        assert lines[1] == "Alice,30"
        assert lines[2] == "Bob,25"

    def test_records_tsv(self, tmp_path):
        path = str(tmp_path / "out.tsv")
        records = [{"A": 1, "B": 2}]
        FastExcel(path).sheet("S", records).save()

        content = open(path).read()
        assert "A\tB" in content
        assert "1\t2" in content

    def test_pandas_csv(self, tmp_path):
        path = str(tmp_path / "out.csv")
        df = pd.DataFrame({"X": [1, 2], "Y": [3.5, 4.5]})
        FastExcel(path).sheet("S", df).save()

        content = open(path).read()
        lines = content.strip().split("\n")
        assert lines[0] == "X,Y"
        assert "3.5" in lines[1]

    def test_polars_csv(self, tmp_path):
        path = str(tmp_path / "out.csv")
        df = pl.DataFrame({"A": ["hello", "world"], "B": [True, False]})
        FastExcel(path).sheet("S", df).save()

        content = open(path).read()
        lines = content.strip().split("\n")
        assert lines[0] == "A,B"
        assert lines[1] == "hello,true"
        assert lines[2] == "world,false"

    def test_xlsx_still_works(self, tmp_path):
        """Ensure .xlsx files still go through Excel path."""
        import openpyxl

        path = str(tmp_path / "out.xlsx")
        FastExcel(path).sheet("S", [{"A": 1}]).save()

        wb = openpyxl.load_workbook(path)
        assert wb.active.cell(2, 1).value == 1
        wb.close()


class TestCSVDataTypes:
    def test_none_values(self, tmp_path):
        path = str(tmp_path / "out.csv")
        records = [{"A": "hello", "B": None}, {"A": None, "B": "world"}]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[1] == "hello,"
        assert lines[2] == ",world"

    def test_boolean_values(self, tmp_path):
        path = str(tmp_path / "out.csv")
        records = [{"flag": True}, {"flag": False}]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[1] == "true"
        assert lines[2] == "false"

    def test_datetime_values(self, tmp_path):
        path = str(tmp_path / "out.csv")
        records = [{"ts": datetime(2024, 6, 15, 10, 30, 45)}]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[1] == "2024-06-15T10:30:45"

    def test_date_values(self, tmp_path):
        path = str(tmp_path / "out.csv")
        records = [{"d": date(2024, 3, 20)}]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[1] == "2024-03-20"

    def test_csv_escaping(self, tmp_path):
        path = str(tmp_path / "out.csv")
        records = [{"A": 'has,comma', "B": 'has"quote', "C": "has\nnewline"}]
        FastExcel(path).sheet("S", records).save()

        content = open(path).read()
        assert '"has,comma"' in content
        assert '"has""quote"' in content
        assert '"has\nnewline"' in content


class TestCSVGenerator:
    def test_generator_input(self, tmp_path):
        def gen():
            for i in range(100):
                yield {"id": i, "value": f"row_{i}"}

        path = str(tmp_path / "out.csv")
        FastExcel(path).sheet("S", gen()).save()

        lines = open(path).read().strip().split("\n")
        assert lines[0] == "id,value"
        assert lines[1] == "0,row_0"
        assert len(lines) == 101  # header + 100 rows


class TestWriteCSVFunction:
    def test_direct_call(self, tmp_path):
        path = str(tmp_path / "out.csv")
        write_csv([{"A": 1, "B": 2}], path)

        content = open(path).read()
        assert "A,B" in content
        assert "1,2" in content

    def test_bytesio(self):
        buf = io.BytesIO()
        write_csv([{"X": "hello"}], buf)
        content = buf.getvalue().decode()
        assert "X" in content
        assert "hello" in content

    def test_custom_delimiter(self, tmp_path):
        path = str(tmp_path / "out.csv")
        write_csv([{"A": 1, "B": 2}], path, delimiter=";")

        content = open(path).read()
        assert "A;B" in content
        assert "1;2" in content
