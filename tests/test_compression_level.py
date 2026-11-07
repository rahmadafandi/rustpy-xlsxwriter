"""The zip compression level is the one knob on what saving costs.

`save()` deflates every worksheet's XML, and on a multi-sheet workbook that is
most of the write — measured at 59% of the total, and it is serial, which is
why threading the row loops was measured and rejected. Lowering the level
moves that cost directly: on four sheets of 200k rows, level 3 cut the whole
write by a third and level 1 by nearly half.

`rust_xlsxwriter` prints to stderr and carries on with the default when the
level is out of range, so the wrapper checks it first — a number nobody meant
should not quietly do nothing.
"""

import zipfile

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [{"id": i, "name": f"name-{i}", "note": "x" * 50} for i in range(2_000)]


def _read_back(path):
    sheet = openpyxl.load_workbook(path).active
    return [[cell.value for cell in row] for row in sheet.iter_rows(max_row=2)]


@pytest.mark.parametrize("level", [1, 3, 6, 9])
def test_every_level_writes_a_readable_file(tmp_path, level):
    path = tmp_path / f"level{level}.xlsx"
    write_worksheet(ROWS, str(path), compression_level=level)

    assert zipfile.ZipFile(path).testzip() is None
    assert _read_back(path) == [
        ["id", "name", "note"],
        [0, "name-0", "x" * 50],
    ]


def test_a_lower_level_gives_a_larger_file(tmp_path):
    """The trade the parameter exists for, stated as a fact about the output."""
    small, large = tmp_path / "l9.xlsx", tmp_path / "l1.xlsx"
    write_worksheet(ROWS, str(small), compression_level=9)
    write_worksheet(ROWS, str(large), compression_level=1)
    assert large.stat().st_size > small.stat().st_size


def test_the_default_is_unchanged(tmp_path):
    """Omitting it must write exactly what earlier versions wrote."""
    without, explicit = tmp_path / "default.xlsx", tmp_path / "six.xlsx"
    write_worksheet(ROWS, str(without))
    write_worksheet(ROWS, str(explicit), compression_level=6)
    assert without.stat().st_size == explicit.stat().st_size


@pytest.mark.parametrize("level", [0, -1, 10, 255, 1000])
def test_an_out_of_range_level_raises(tmp_path, level):
    """The crate would only print to stderr and use the default."""
    with pytest.raises(ValueError, match=r"compression_level must be between 1 and 9"):
        write_worksheet(ROWS, str(tmp_path / "bad.xlsx"), compression_level=level)


def test_the_message_names_the_level(tmp_path):
    with pytest.raises(ValueError, match=r"not 12"):
        write_worksheet(ROWS, str(tmp_path / "bad.xlsx"), compression_level=12)


def test_multi_sheet_takes_it(tmp_path):
    path = tmp_path / "multi.xlsx"
    write_worksheets(
        [("A", ROWS), ("B", ROWS)],
        str(path),
        compression_level=1,
    )
    assert openpyxl.load_workbook(path).sheetnames == ["A", "B"]


def test_fast_excel_takes_it(tmp_path):
    path = tmp_path / "fast.xlsx"
    FastExcel(str(path), compression_level=1).sheet("S", ROWS).save()
    assert zipfile.ZipFile(path).testzip() is None
    assert _read_back(path)[0] == ["id", "name", "note"]


def test_fast_excel_rejects_an_out_of_range_level(tmp_path):
    with pytest.raises(ValueError, match=r"compression_level must be between 1 and 9"):
        FastExcel(str(tmp_path / "fast.xlsx"), compression_level=0).sheet("S", ROWS).save()


def test_csv_output_warns_that_it_is_ignored(tmp_path):
    """A zip level means nothing to a CSV, and silently dropping an option is
    what the warning for every other Excel-only option exists to prevent."""
    with pytest.warns(UserWarning, match=r"compression_level"):
        FastExcel(str(tmp_path / "out.csv"), compression_level=1).sheet("S", ROWS).save()
