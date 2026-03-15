import openpyxl

from rustpy_xlsxwriter import FastExcel


class TestPassword:
    def test_protected_sheet(self, tmp_path, small_records):
        path = str(tmp_path / "protected.xlsx")
        FastExcel(path, password="secret").sheet("Sheet1", small_records).save()

        wb = openpyxl.load_workbook(path)
        assert wb.active.protection.sheet is True
        wb.close()

    def test_unprotected_sheet(self, tmp_path, small_records):
        path = str(tmp_path / "unprotected.xlsx")
        FastExcel(path).sheet("Sheet1", small_records).save()

        wb = openpyxl.load_workbook(path)
        assert wb.active.protection.sheet is False
        wb.close()
