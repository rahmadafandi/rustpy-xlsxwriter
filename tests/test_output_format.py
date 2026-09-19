"""Explicit ``output_format`` on the builder.

Extension sniffing covers paths, but a buffer has no extension — so before
this, ``FastExcel(BytesIO())`` could only ever produce xlsx, and CSV into a
buffer meant dropping down to ``write_csv``.
"""

import io
import zipfile

import pytest

from rustpy_xlsxwriter import FastExcel

ROWS = [{"name": "Alice", "age": 30}, {"name": "Bob", "age": 25}]


def test_csv_into_buffer():
    buf = io.BytesIO()
    FastExcel(buf, output_format="csv").sheet("S", ROWS).save()
    assert buf.getvalue().decode() == "name,age\nAlice,30\nBob,25\n"


def test_tsv_into_buffer():
    buf = io.BytesIO()
    FastExcel(buf, output_format="tsv").sheet("S", ROWS).save()
    assert buf.getvalue().decode() == "name\tage\nAlice\t30\nBob\t25\n"


def test_buffer_still_defaults_to_xlsx():
    buf = io.BytesIO()
    FastExcel(buf).sheet("S", ROWS).save()
    assert zipfile.is_zipfile(io.BytesIO(buf.getvalue()))


def test_explicit_format_overrides_the_extension(tmp_path):
    """A .txt path asked to hold CSV writes CSV, not Excel."""
    path = tmp_path / "data.txt"
    FastExcel(str(path), output_format="csv").sheet("S", ROWS).save()
    assert path.read_text() == "name,age\nAlice,30\nBob,25\n"


def test_xlsx_forced_over_a_csv_extension(tmp_path):
    path = tmp_path / "confusing.csv"
    FastExcel(str(path), output_format="xlsx").sheet("S", ROWS).save()
    assert zipfile.is_zipfile(path)


def test_extension_still_wins_when_unset(tmp_path):
    path = tmp_path / "data.csv"
    FastExcel(str(path)).sheet("S", ROWS).save()
    assert path.read_text() == "name,age\nAlice,30\nBob,25\n"


def test_pathlib_extension_is_detected(tmp_path):
    """Path targets reach the extension check without being coerced first."""
    path = tmp_path / "data.tsv"
    FastExcel(path).sheet("S", ROWS).save()
    assert path.read_text() == "name\tage\nAlice\t30\nBob\t25\n"


def test_unknown_format_is_rejected():
    with pytest.raises(ValueError, match="output_format must be one of"):
        FastExcel(io.BytesIO(), output_format="parquet")


def test_csv_into_buffer_still_warns_about_excel_options():
    buf = io.BytesIO()
    writer = FastExcel(buf, output_format="csv").format(bold_headers=True)
    with pytest.warns(UserWarning, match="ignores Excel-only options"):
        writer.sheet("S", ROWS).save()


def test_multi_sheet_csv_is_rejected():
    buf = io.BytesIO()
    writer = FastExcel(buf, output_format="csv").sheet("A", ROWS).sheet("B", ROWS)
    with pytest.raises(ValueError, match="single sheet"):
        writer.save()
