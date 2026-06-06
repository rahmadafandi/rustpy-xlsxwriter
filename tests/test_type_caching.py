"""Type-cache edge cases shared by the Excel and CSV Records paths.

The first-row type cache must not mis-handle a Python ``bool`` that lands in a
column whose first row was a plain ``int`` (bool is a subclass of int). Excel
and CSV must agree: a bool is always written as a boolean, never as a number.
"""

import openpyxl

from rustpy_xlsxwriter import FastExcel


class TestBoolInIntColumn:
    def test_excel_writes_bool_as_boolean(self, tmp_path):
        # Column starts int (caches ColType::Int), then a bool appears.
        records = [{"x": 1}, {"x": True}, {"x": 2}]
        path = str(tmp_path / "bool_int.xlsx")
        FastExcel(path).sheet("Sheet1", records).save()

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        # Row 3 (cell 3,1) is the bool. It must be a real boolean True,
        # not the number 1.
        cell = ws.cell(3, 1).value
        assert cell is True
        assert isinstance(cell, bool)
        wb.close()

    def test_csv_writes_bool_as_boolean(self, tmp_path):
        records = [{"x": 1}, {"x": True}, {"x": 2}]
        path = str(tmp_path / "bool_int.csv")
        FastExcel(path).sheet("Sheet1", records).save()

        content = open(path, encoding="utf-8").read()
        # Header + three rows; the bool row is "true", not "1".
        assert content == "x\n1\ntrue\n2\n"
