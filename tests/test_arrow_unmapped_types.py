"""Arrow column types beyond the plain scalars.

The Arrow path used to write any type it had no accessor for as empty cells,
and report success. A pandas `category` column, a polars `Enum`, a time, a
decimal, a list: the file came back with the column blank and nothing said so.

Now each type either arrives with its value or raises a `ValueError` naming
the column. Every case runs through both writers and, where the library can
produce the type, through both polars and pandas, since they hand over
different Arrow types for the same data (string views vs large strings,
different dictionary key widths, `time64[ns]` vs `time64[us]`).
"""

import datetime as dt
import decimal
import io

import openpyxl
import pytest

from rustpy_xlsxwriter import write_csv, write_worksheet

pl = pytest.importorskip("polars")
pd = pytest.importorskip("pandas")
pytest.importorskip("pyarrow")

T1 = dt.time(13, 30)
T2 = dt.time(1, 2, 3, 500000)
D1 = decimal.Decimal("1.10")
D2 = decimal.Decimal("-2.25")


def _csv(df, **kwargs):
    buf = io.BytesIO()
    write_csv(df, buf, **kwargs)
    return buf.getvalue().decode().splitlines()


def _cells(tmp_path, df, **kwargs):
    path = tmp_path / "out.xlsx"
    write_worksheet(df, str(path), **kwargs)
    ws = openpyxl.load_workbook(path).active
    return [[c for c in row] for row in ws.iter_rows(min_row=2)]


def _values(tmp_path, df, **kwargs):
    return [[c.value for c in row] for row in _cells(tmp_path, df, **kwargs)]


# label -> (frame factory, values Excel should hold, CSV lines after the header)
SUPPORTED = {
    "polars-categorical": (
        lambda: pl.DataFrame({"v": pl.Series(["open", "closed"], dtype=pl.Categorical)}),
        ["open", "closed"],
        ["open", "closed"],
    ),
    "polars-enum": (
        lambda: pl.DataFrame({"v": pl.Series(["b", "a"], dtype=pl.Enum(["a", "b"]))}),
        ["b", "a"],
        ["b", "a"],
    ),
    "pandas-category": (
        lambda: pd.DataFrame({"v": pd.Categorical(["open", "closed"])}),
        ["open", "closed"],
        ["open", "closed"],
    ),
    "pandas-category-int": (
        lambda: pd.DataFrame({"v": pd.Categorical([10, 20])}),
        [10, 20],
        ["10", "20"],
    ),
    "polars-time": (
        lambda: pl.DataFrame({"v": [T1, T2]}),
        [T1, T2],
        ["13:30:00", "01:02:03.500000"],
    ),
    "pandas-time": (
        lambda: pd.DataFrame({"v": [T1, T2]}),
        [T1, T2],
        ["13:30:00", "01:02:03.500000"],
    ),
    "polars-decimal": (
        lambda: pl.DataFrame({"v": [D1, D2]}),
        [1.1, -2.25],
        ["1.1", "-2.25"],
    ),
    "pandas-decimal": (
        lambda: pd.DataFrame({"v": [D1, D2]}),
        [1.1, -2.25],
        ["1.1", "-2.25"],
    ),
}


@pytest.mark.parametrize("label", sorted(SUPPORTED))
def test_excel_writes_the_value(tmp_path, label):
    make, expected, _ = SUPPORTED[label]
    assert [row[0] for row in _values(tmp_path, make())] == expected


@pytest.mark.parametrize("label", sorted(SUPPORTED))
def test_csv_writes_the_value(label):
    make, _, expected = SUPPORTED[label]
    assert _csv(make()) == ["v", *expected]


def _pa_frame(arrow_type, values):
    """A pandas frame backed by an exact Arrow type, for the types neither
    library produces from plain Python values."""
    import pyarrow as pa

    return pa.table({"v": pa.array(values, type=arrow_type)}).to_pandas(
        types_mapper=pd.ArrowDtype
    )


@pytest.mark.parametrize(
    "unit, values, csv",
    [
        ("s", [T1], "13:30:00"),
        ("ms", [dt.time(1, 2, 3, 250000)], "01:02:03.250000"),
        ("us", [T2], "01:02:03.500000"),
        ("ns", [T2], "01:02:03.500000"),
    ],
)
def test_every_time_unit(tmp_path, unit, values, csv):
    import pyarrow as pa

    arrow_type = pa.time32(unit) if unit in ("s", "ms") else pa.time64(unit)
    df = _pa_frame(arrow_type, values)
    assert _values(tmp_path, df) == [values]
    assert _csv(df) == ["v", csv]


def test_sub_microsecond_time_keeps_its_digits_in_csv():
    import pyarrow as pa

    table = pa.table({"v": pa.array([3_723_000_000_001], type=pa.time64("ns"))})
    assert _csv(table.to_pandas(types_mapper=pd.ArrowDtype)) == ["v", "01:02:03.000000001"]


@pytest.mark.parametrize("width", [32, 64, 128, 256])
def test_every_decimal_width(tmp_path, width):
    import pyarrow as pa

    make = getattr(pa, f"decimal{width}")
    df = _pa_frame(make(9, 2), [D1, D2])
    assert _values(tmp_path, df) == [[1.1], [-2.25]]
    assert _csv(df) == ["v", "1.1", "-2.25"]


def test_decimal_is_correctly_rounded_past_two_to_the_53(tmp_path):
    # The unscaled integer is past 2^53, so converting it to f64 and then
    # dividing by 10^4 rounds twice and lands one ulp off (…222.791).
    value = decimal.Decimal("14495820535222.7900")
    df = pl.DataFrame({"v": [value]})
    assert _values(tmp_path, df) == [[float(value)]]
    assert _csv(df) == ["v", repr(float(value))]


def test_time_cells_carry_a_time_format(tmp_path):
    cell = _cells(tmp_path, pl.DataFrame({"v": [T1]}))[0][0]
    assert cell.number_format == "hh:mm:ss"


def test_column_format_overrides_the_time_format(tmp_path):
    from rustpy_xlsxwriter import Format

    fmt = Format().set_num_format("hh:mm:ss.000")
    cell = _cells(tmp_path, pl.DataFrame({"v": [T2]}), column_formats={"v": fmt})[0][0]
    assert cell.number_format == "hh:mm:ss.000"
    assert cell.value == T2


def test_banded_time_keeps_its_format(tmp_path):
    rows = _cells(tmp_path, pl.DataFrame({"v": [T1, T2, T1]}), banded_rows="#EEEEEE")
    assert [r[0].number_format for r in rows] == ["hh:mm:ss"] * 3
    assert [r[0].value for r in rows] == [T1, T2, T1]


@pytest.mark.parametrize(
    "make",
    [
        lambda: pl.DataFrame({"v": pl.Series(["a", None], dtype=pl.Categorical)}),
        lambda: pd.DataFrame({"v": pd.Categorical(["a", None])}),
    ],
    ids=["polars", "pandas"],
)
def test_null_in_a_categorical_takes_na_rep(tmp_path, make):
    assert _values(tmp_path, make(), na_rep="-") == [["a"], ["-"]]
    assert _csv(make(), na_rep="-") == ["v", "a", "-"]


def test_all_null_categorical_with_an_empty_dictionary(tmp_path):
    df = pl.DataFrame({"v": pl.Series([None, None], dtype=pl.Categorical)})
    assert _values(tmp_path, df, na_rep="-") == [["-"], ["-"]]
    assert _csv(df, na_rep="-") == ["v", "-", "-"]


def test_categorical_supplies_url_display_text(tmp_path):
    df = pl.DataFrame(
        {
            "link": ["https://example.com"],
            "label": pl.Series(["Example"], dtype=pl.Categorical),
        }
    )
    path = tmp_path / "out.xlsx"
    write_worksheet(df, str(path), url_columns={"link": "label"})
    cell = openpyxl.load_workbook(path).active.cell(2, 1)
    assert cell.value == "Example"
    assert cell.hyperlink.target == "https://example.com"


# label -> (frame factory, a word the message must contain)
UNSUPPORTED = {
    "polars-duration": (lambda: pl.DataFrame({"v": [dt.timedelta(seconds=5)]}), "duration"),
    "pandas-duration": (lambda: pd.DataFrame({"v": pd.to_timedelta([5], unit="s")}), "duration"),
    "polars-binary": (lambda: pl.DataFrame({"v": [b"x"]}), "Binary"),
    "pandas-binary": (lambda: pd.DataFrame({"v": [b"x"]}), "Binary"),
    "polars-list": (lambda: pl.DataFrame({"v": [[1, 2]]}), "List"),
    "polars-struct": (lambda: pl.DataFrame({"v": [{"a": 1}]}), "Struct"),
}


@pytest.mark.parametrize("label", sorted(UNSUPPORTED))
def test_excel_refuses_and_names_the_column(tmp_path, label):
    make, word = UNSUPPORTED[label]
    with pytest.raises(ValueError, match=f"column 'v'.*{word}"):
        write_worksheet(make(), str(tmp_path / "out.xlsx"))


@pytest.mark.parametrize("label", sorted(UNSUPPORTED))
def test_csv_refuses_and_names_the_column(label):
    make, word = UNSUPPORTED[label]
    with pytest.raises(ValueError, match=f"column 'v'.*{word}"):
        write_csv(make(), io.BytesIO())


def test_duration_message_says_what_to_do():
    with pytest.raises(ValueError, match="string, or to a number of seconds"):
        write_csv(pl.DataFrame({"v": [dt.timedelta(seconds=5)]}), io.BytesIO())


def test_message_says_earlier_files_were_missing_the_column():
    with pytest.raises(ValueError, match="empty cells without an error"):
        write_csv(pl.DataFrame({"v": [[1]]}), io.BytesIO())


@pytest.mark.parametrize(
    "make",
    [
        lambda: pl.DataFrame({"id": [1, 2], "tags": [[1], [2]]}),
        lambda: pd.DataFrame({"id": [1, 2], "blob": [b"x", b"y"]}),
    ],
    ids=["polars", "pandas"],
)
def test_mixed_frame_raises_instead_of_writing_the_good_columns(tmp_path, make):
    """The realistic case: most columns are fine and one is not. The good
    columns used to survive with the bad one blank, which is the hardest
    version of the bug to notice."""
    path = tmp_path / "out.xlsx"
    with pytest.raises(ValueError, match="column '(tags|blob)'"):
        write_worksheet(make(), str(path))
    with pytest.raises(ValueError, match="column '(tags|blob)'"):
        write_csv(make(), str(tmp_path / "out.csv"))
    assert not (tmp_path / "out.csv").exists()


def test_empty_frame_with_an_unsupported_column_still_raises(tmp_path):
    df = pl.DataFrame({"v": []}, schema={"v": pl.List(pl.Int64)})
    with pytest.raises(ValueError, match="column 'v'"):
        write_worksheet(df, str(tmp_path / "out.xlsx"))
    with pytest.raises(ValueError, match="column 'v'"):
        write_csv(df, io.BytesIO())


def test_csv_columns_can_leave_an_unsupported_column_out():
    df = pl.DataFrame({"id": [1], "tags": [[1, 2]]})
    assert _csv(df, columns=["id"]) == ["id", "1"]


def test_null_column_is_still_blank(tmp_path):
    df = pl.DataFrame({"id": [1, 2], "nothing": [None, None]})
    assert _values(tmp_path, df) == [[1, None], [2, None]]
    assert _csv(df) == ["id,nothing", "1,", "2,"]
