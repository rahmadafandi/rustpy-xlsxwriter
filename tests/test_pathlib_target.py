"""os.PathLike (e.g. pathlib.Path) acceptance across every public entry
point. `_coerce_target` runs os.fspath() on PathLike and passes other
objects (file paths as str, file-like buffers) through unchanged."""

from pathlib import Path

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets


def _read_cell(path, row, col, sheet=None):
    wb = openpyxl.load_workbook(str(path))
    ws = wb[sheet] if sheet else wb.active
    val = ws.cell(row, col).value
    wb.close()
    return val


class TestPathLikeTarget:
    def test_fastexcel_path(self, tmp_path):
        p = Path(tmp_path) / "fe.xlsx"
        records = [{"Name": "Alice", "Age": 30}]
        FastExcel(p).sheet("Data", records).save()

        assert p.exists()
        assert _read_cell(p, 1, 1) == "Name"
        assert _read_cell(p, 2, 1) == "Alice"
        assert _read_cell(p, 2, 2) == 30

    def test_write_worksheet_path(self, tmp_path):
        p = Path(tmp_path) / "ws.xlsx"
        records = [{"A": 1, "B": "x"}, {"A": 2, "B": "y"}]
        write_worksheet(records, p, sheet_name="Sheet1")

        assert p.exists()
        assert _read_cell(p, 1, 1, "Sheet1") == "A"
        assert _read_cell(p, 2, 1, "Sheet1") == 1
        assert _read_cell(p, 3, 2, "Sheet1") == "y"

    def test_write_worksheets_path(self, tmp_path):
        p = Path(tmp_path) / "wss.xlsx"
        write_worksheets(
            [("S1", [{"A": 1}]), ("S2", [{"B": 2}])],
            p,
        )

        assert p.exists()
        assert _read_cell(p, 1, 1, "S1") == "A"
        assert _read_cell(p, 2, 1, "S1") == 1
        assert _read_cell(p, 2, 1, "S2") == 2

    def test_str_still_works(self, tmp_path):
        """str path must keep working (pass-through, not coerced)."""
        p = tmp_path / "str.xlsx"
        FastExcel(str(p)).sheet("Data", [{"A": 9}]).save()
        assert _read_cell(p, 2, 1) == 9
