"""`Decimal` in CSV keeps its exact digits.

Excel has no decimal type, so xlsx gets the nearest float. CSV is text, and
it used to get that same float: `Decimal("1.10")` came out as `1.1`, and
`Decimal("0.1")` as a float that is really 0.1000000000000000055…. Writing
the decimal's own text round-trips exactly. The Arrow path does the same
from the dtype's scale, with no float in between.
"""

import io
from decimal import Decimal

import openpyxl
import pytest

from rustpy_xlsxwriter import write_csv, write_worksheet


def _csv(data, **kwargs):
    buf = io.BytesIO()
    write_csv(data, buf, **kwargs)
    return buf.getvalue().decode().splitlines()


@pytest.mark.parametrize(
    "value, text",
    [
        (Decimal("1.10"), "1.10"),
        (Decimal("0.1"), "0.1"),
        (Decimal("-2.250"), "-2.250"),
        (Decimal("12345678901234567890.12"), "12345678901234567890.12"),
        (Decimal("1E+2"), "1E+2"),
        (Decimal("-0.000"), "-0.000"),
    ],
)
def test_records_write_the_exact_text(value, text):
    (line,) = _csv([{"d": value}], header=False)
    assert line == text
    assert Decimal(line) == value
    assert str(Decimal(line)) == str(value)


def test_nan_and_infinity_take_the_float_representations():
    rows = [{"d": Decimal(v)} for v in ("NaN", "sNaN", "Infinity", "-Infinity")]
    assert _csv(rows, header=False) == ["", "", "", ""]
    assert _csv(rows, header=False, na_rep="NA", inf_value="INF") == ["NA", "NA", "INF", "-INF"]


def test_cached_decimal_column_still_catches_other_types():
    rows = [{"d": Decimal("1.10")}, {"d": None}, {"d": 2.5}, {"d": "x"}, {"d": Decimal("3.00")}]
    assert _csv(rows, header=False) == ["1.10", "", "2.5", "x", "3.00"]


def test_decimal_next_to_floats_and_ints():
    assert _csv([{"f": 1.10, "i": 3, "d": Decimal("1.10")}]) == ["f,i,d", "1.1,3,1.10"]


@pytest.mark.parametrize("lib", ["polars", "pandas"])
def test_dataframe_decimals_keep_their_scale(lib):
    mod = pytest.importorskip(lib)
    pytest.importorskip("pyarrow")
    df = mod.DataFrame({"d": [Decimal("1.10"), Decimal("-2.25"), None]})
    assert _csv(df) == ["d", "1.10", "-2.25", ""]


def test_polars_non_arrow_path_writes_exact_text():
    """A polars Series list of Decimal objects reaches the object cascade."""
    pl = pytest.importorskip("polars")
    values = pl.Series([Decimal("1.10"), Decimal("0.1")]).to_list()
    assert _csv([{"d": v} for v in values], header=False) == ["1.10", "0.10"]


def test_xlsx_still_writes_the_nearest_float(tmp_path):
    path = tmp_path / "out.xlsx"
    write_worksheet([{"d": Decimal("1.10")}, {"d": Decimal("sNaN")}], str(path), na_rep="NA")
    ws = openpyxl.load_workbook(path).active
    assert [ws["A2"].value, ws["A3"].value] == [1.1, "NA"]
