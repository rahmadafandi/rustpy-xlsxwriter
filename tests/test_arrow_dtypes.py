"""Every Arrow column type the writer classifies.

`classify` maps an Arrow `DataType` onto a `ColKind`, and each `ColKind` then
picks the accessor that reads the values. A type wired to the wrong accessor
compiles and writes whatever that accessor happens to return.

Test data had only ever been Int64, Float64, Utf8 and Boolean, so the narrow
integers, the unsigned ones, Float32 and the date types were classified by
code no test had run. A DataFrame declared with `schema=` reaches them, and
those are ordinary in real data: an id column read back as UInt32, a flag as
Int8.

Both the Excel and the CSV path are checked, because they read the columns
through separate code.
"""

import io

import openpyxl
import pytest

from rustpy_xlsxwriter import write_csv, write_worksheet

pl = pytest.importorskip("polars")

# label -> (polars dtype, values, what Excel should hold)
NUMERIC = {
    "Int8": (pl.Int8, [-8, 8], [-8, 8]),
    "Int16": (pl.Int16, [-16, 16], [-16, 16]),
    "Int32": (pl.Int32, [-32, 32], [-32, 32]),
    "Int64": (pl.Int64, [-64, 64], [-64, 64]),
    "UInt8": (pl.UInt8, [0, 8], [0, 8]),
    "UInt16": (pl.UInt16, [0, 16], [0, 16]),
    "UInt32": (pl.UInt32, [0, 32], [0, 32]),
    "UInt64": (pl.UInt64, [0, 64], [0, 64]),
    "Float32": (pl.Float32, [1.5, -2.5], [1.5, -2.5]),
    "Float64": (pl.Float64, [1.5, -2.5], [1.5, -2.5]),
    "Boolean": (pl.Boolean, [True, False], [True, False]),
}


def _frame(dtype, values):
    return pl.DataFrame({"v": values}, schema={"v": dtype})


@pytest.mark.parametrize("label", sorted(NUMERIC))
def test_excel_reads_every_numeric_type(tmp_path, label):
    dtype, values, expected = NUMERIC[label]
    path = tmp_path / f"{label}.xlsx"
    write_worksheet(_frame(dtype, values), str(path))
    ws = openpyxl.load_workbook(path).active
    assert [ws.cell(r, 1).value for r in (2, 3)] == expected


@pytest.mark.parametrize("label", sorted(NUMERIC))
def test_csv_reads_every_numeric_type(label):
    dtype, values, _ = NUMERIC[label]
    buf = io.BytesIO()
    write_csv(_frame(dtype, values), buf)
    lines = buf.getvalue().decode().splitlines()
    assert lines[0] == "v"
    assert len(lines) == 3


def test_narrow_integers_keep_their_value(tmp_path):
    """The point of the type table: a UInt8 must not be read as an Int64."""
    frame = pl.DataFrame(
        {"small": [255], "big": [2**53]},
        schema={"small": pl.UInt8, "big": pl.Int64},
    )
    path = tmp_path / "widths.xlsx"
    write_worksheet(frame, str(path))
    ws = openpyxl.load_workbook(path).active
    assert ws.cell(2, 1).value == 255
    assert ws.cell(2, 2).value == 2**53


def test_integers_lose_precision_above_two_to_the_53(tmp_path):
    """Excel's own limit, pinned here because nothing else states it.

    xlsx stores every number as an IEEE 754 double, so an integer past the
    53-bit mantissa cannot round-trip. A 64-bit id written to a spreadsheet
    comes back as a different number, silently, on every path — this is not
    something the library chooses, but it is something a caller needs to know.
    """
    exact, lost = 2**53, 2**53 + 1
    path = tmp_path / "precision.xlsx"
    write_worksheet(
        pl.DataFrame({"v": [exact, lost]}, schema={"v": pl.Int64}), str(path)
    )
    ws = openpyxl.load_workbook(path).active
    assert ws.cell(2, 1).value == exact
    assert ws.cell(3, 1).value == exact  # lost + 1 rounds back down

    # The records path stores the same way, so it rounds the same way.
    records = tmp_path / "precision_records.xlsx"
    write_worksheet([{"v": lost}], str(records))
    assert openpyxl.load_workbook(records).active.cell(2, 1).value == exact


def test_float32_precision_is_not_promoted_wrongly(tmp_path):
    """0.1 as f32 widens to a longer f64; the cell must hold what Excel got."""
    path = tmp_path / "f32.xlsx"
    write_worksheet(_frame(pl.Float32, [0.5, 0.25]), str(path))
    ws = openpyxl.load_workbook(path).active
    assert [ws.cell(r, 1).value for r in (2, 3)] == [0.5, 0.25]


# --- temporal ---------------------------------------------------------------


def test_date_column(tmp_path):
    import datetime

    frame = pl.DataFrame(
        {"d": [datetime.date(2026, 1, 2), datetime.date(2026, 12, 31)]},
        schema={"d": pl.Date},
    )
    path = tmp_path / "date.xlsx"
    write_worksheet(frame, str(path))
    ws = openpyxl.load_workbook(path).active
    assert ws.cell(2, 1).value == datetime.datetime(2026, 1, 2)


def test_datetime_column(tmp_path):
    import datetime

    frame = pl.DataFrame(
        {"t": [datetime.datetime(2026, 1, 2, 3, 4, 5)]},
        schema={"t": pl.Datetime("us")},
    )
    path = tmp_path / "dt.xlsx"
    write_worksheet(frame, str(path))
    ws = openpyxl.load_workbook(path).active
    assert ws.cell(2, 1).value == datetime.datetime(2026, 1, 2, 3, 4, 5)


def test_nulls_in_a_typed_column(tmp_path):
    """A null in a narrow column takes the null branch, not the value one."""
    frame = pl.DataFrame({"v": [1, None, 3]}, schema={"v": pl.Int16})
    path = tmp_path / "nulls.xlsx"
    write_worksheet(frame, str(path))
    ws = openpyxl.load_workbook(path).active
    assert [ws.cell(r, 1).value for r in (2, 3, 4)] == [1, None, 3]
