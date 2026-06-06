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


class TestCSVRecordsTypeCache:
    """Exercise the Records first-row type cache in csv_writer: the cached
    column type is set from row 1, hit on stable columns, and must fall back
    to the full cascade when a later row's value type differs."""

    def test_stable_types_cache_hit(self, tmp_path):
        path = str(tmp_path / "stable.csv")
        records = [
            {"i": 1, "f": 1.5, "s": "a", "b": True},
            {"i": 2, "f": 2.5, "s": "b", "b": False},
            {"i": 3, "f": 3.5, "s": "c", "b": True},
        ]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[0] == "i,f,s,b"
        assert lines[1] == "1,1.5,a,true"
        assert lines[2] == "2,2.5,b,false"
        assert lines[3] == "3,3.5,c,true"

    def test_type_switch_falls_back(self, tmp_path):
        """Column ``v``: int in row 1 (caches Int), then string / float /
        bool in later rows — cache misses must re-run the cascade, not coerce
        to the cached type."""
        path = str(tmp_path / "switch.csv")
        records = [
            {"v": 1},
            {"v": "two"},
            {"v": 3.5},
            {"v": True},
        ]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[0] == "v"
        assert lines[1] == "1"
        assert lines[2] == "two"
        assert lines[3] == "3.5"
        assert lines[4] == "true"

    def test_none_in_first_row_then_typed(self, tmp_path):
        """A ``None`` in row 1 leaves the column type Unknown; the first
        non-null value sets the cache, and escaping still applies on miss."""
        path = str(tmp_path / "none_first.csv")
        records = [
            {"x": None},
            {"x": "has,comma"},
            {"x": "plain"},
        ]
        FastExcel(path).sheet("S", records).save()

        lines = open(path).read().strip().split("\n")
        assert lines[0] == "x"
        assert lines[1] == ""
        assert lines[2] == '"has,comma"'
        assert lines[3] == "plain"


class TestCSVArrowPath:
    """Verify DataFrame → CSV goes through the Arrow C Data Interface
    and produces the expected output for mixed types + nulls."""

    def test_pandas_arrow_types(self, tmp_path):
        path = str(tmp_path / "arrow.csv")
        df = pd.DataFrame(
            {
                "i": [1, 2, 3],
                "f": [1.5, 2.5, 3.5],
                "s": ["a", "b", "c"],
                "b": [True, False, True],
                "ts": pd.to_datetime(
                    ["2024-01-02 03:04:05", "2025-06-07 08:09:10", pd.NaT]
                ),
            }
        )
        write_csv(df, path)

        lines = open(path).read().strip().split("\n")
        assert lines[0] == "i,f,s,b,ts"
        assert lines[1] == "1,1.5,a,true,2024-01-02T03:04:05"
        assert lines[2] == "2,2.5,b,false,2025-06-07T08:09:10"
        # Null datetime → empty cell
        assert lines[3].split(",")[4] == ""

    def test_polars_arrow_types(self, tmp_path):
        path = str(tmp_path / "arrow_pl.csv")
        df = pl.DataFrame(
            {
                "i": [10, 20, 30],
                "s": ["x", "y", None],
                "b": [True, None, False],
            }
        )
        write_csv(df, path)

        lines = open(path).read().strip().split("\n")
        assert lines[0] == "i,s,b"
        assert lines[1] == "10,x,true"
        # Row 1: b is None → trailing empty
        assert lines[2] == "20,y,"
        # Row 2: s is None → middle empty
        assert lines[3] == "30,,false"

    def test_pandas_escape_in_arrow_path(self, tmp_path):
        path = str(tmp_path / "arrow_esc.csv")
        df = pd.DataFrame({"A": ['has,comma', 'has"quote']})
        write_csv(df, path)

        content = open(path).read()
        assert '"has,comma"' in content
        assert '"has""quote"' in content

    def test_pathlib_path_target(self, tmp_path):
        """write_csv accepts a pathlib.Path, not only str."""
        from pathlib import Path

        p = Path(tmp_path) / "pathlib.csv"
        df = pd.DataFrame({"A": [1, 2]})
        write_csv(df, p)
        assert p.read_text().startswith("A\n")


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


class TestCSVFormulaSanitization:
    """Opt-in CSV-injection guard (`sanitize_formulas`)."""

    def test_off_by_default(self, tmp_path):
        path = str(tmp_path / "raw.csv")
        records = [{"a": "=1+2", "b": "@cmd"}, {"a": "-5", "b": "safe"}]
        FastExcel(path).sheet("S", records).save()
        content = open(path, encoding="utf-8").read()
        # No guard quote prefixed.
        assert content == "a,b\n=1+2,@cmd\n-5,safe\n"

    def test_records_sanitized(self, tmp_path):
        path = str(tmp_path / "safe.csv")
        records = [{"a": "=1+2", "b": "@cmd"}, {"a": "-5", "b": "safe"}]
        FastExcel(path, sanitize_formulas=True).sheet("S", records).save()
        content = open(path, encoding="utf-8").read()
        assert content == "a,b\n'=1+2,'@cmd\n'-5,safe\n"

    def test_functional_api_kwarg(self, tmp_path):
        path = str(tmp_path / "fn.csv")
        write_csv([{"a": "=BAD"}], path, sanitize_formulas=True)
        assert open(path, encoding="utf-8").read() == "a\n'=BAD\n"

    def test_polars_sanitized(self, tmp_path):
        path = str(tmp_path / "pl.csv")
        df = pl.DataFrame({"a": ["=1+2", "ok"]})
        FastExcel(path, sanitize_formulas=True).sheet("S", df).save()
        content = open(path, encoding="utf-8").read()
        assert content == "a\n'=1+2\nok\n"

    def test_sanitized_header(self, tmp_path):
        path = str(tmp_path / "hdr.csv")
        FastExcel(path, sanitize_formulas=True).sheet("S", [{"=col": 1}]).save()
        content = open(path, encoding="utf-8").read()
        assert content == "'=col\n1\n"
