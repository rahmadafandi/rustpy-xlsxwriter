"""Timezone-aware datetimes in xlsx: always the local wall clock.

Excel has no timezone, so something has to be dropped. The rule is one
sentence: an aware datetime is written as its own local wall clock. 12:00 UTC
converted to Asia/Jakarta is written as 19:00, which is what the frame shows
when printed, and what polars' ``replace_time_zone(None)`` and pandas'
``tz_localize(None)`` give.

A Python ``datetime`` always worked that way. The Arrow path used to write
the UTC time instead, so the same data gave different cells depending on
whether it arrived as dicts or as a DataFrame. The conversion now happens in
Python, with the DataFrame library's own methods, before the Arrow export.
CSV keeps the instant with its offset, and is unchanged.
"""

import datetime as dt
import io

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_csv, write_worksheet, write_worksheets
from rustpy_xlsxwriter._timezones import local_wall_clock

pl = pytest.importorskip("polars")
pd = pytest.importorskip("pandas")
pa = pytest.importorskip("pyarrow")

UTC_NOON = dt.datetime(2026, 1, 1, 12, tzinfo=dt.timezone.utc)
PLUS_7 = dt.timezone(dt.timedelta(hours=7))
LOCAL = dt.datetime(2026, 1, 1, 19)


def _values(tmp_path, data, **kwargs):
    path = tmp_path / "out.xlsx"
    write_worksheet(data, str(path), **kwargs)
    return [[c.value for c in row] for row in openpyxl.load_workbook(path).active.iter_rows(min_row=2)]


def _frames():
    return {
        "polars": pl.DataFrame({"id": [1], "t": [UTC_NOON]}).with_columns(
            pl.col("t").dt.convert_time_zone("Asia/Jakarta")
        ),
        "pandas-tz-dtype": pd.DataFrame(
            {"id": [1], "t": pd.Series([UTC_NOON]).dt.tz_convert("Asia/Jakarta")}
        ),
        "pandas-arrow-dtype": pd.DataFrame(
            {
                "id": [1],
                "t": pd.array(
                    [UTC_NOON], dtype=pd.ArrowDtype(pa.timestamp("us", tz="Asia/Jakarta"))
                ),
            }
        ),
        "pandas-object": pd.DataFrame(
            {"id": [1], "t": pd.Series([dt.datetime(2026, 1, 1, 19, tzinfo=PLUS_7)], dtype=object)}
        ),
        "records": [{"id": 1, "t": UTC_NOON.astimezone(PLUS_7)}],
    }


SOURCES = ["polars", "pandas-tz-dtype", "pandas-arrow-dtype", "pandas-object", "records"]


@pytest.mark.parametrize("source", SOURCES)
def test_every_input_writes_the_local_wall_clock(tmp_path, source):
    assert _values(tmp_path, _frames()[source]) == [[1, LOCAL]]


def test_matches_the_libraries_own_zone_drop():
    frame = _frames()["polars"]
    assert frame.with_columns(pl.col("t").dt.replace_time_zone(None))["t"][0] == LOCAL
    series = _frames()["pandas-tz-dtype"]["t"]
    assert series.dt.tz_localize(None)[0] == pd.Timestamp(LOCAL)


def test_object_column_with_mixed_offsets_keeps_each_wall_clock(tmp_path):
    """pyarrow would coerce such a column to its first value's zone; each
    value keeps its own wall clock instead, as a list of dicts does."""
    values = [dt.datetime(2026, 1, 1, 12, tzinfo=PLUS_7), dt.datetime(2026, 1, 1, 12, tzinfo=dt.timezone.utc)]
    df = pd.DataFrame({"t": pd.Series(values, dtype=object)})
    assert _values(tmp_path, df) == [[dt.datetime(2026, 1, 1, 12)]] * 2


def test_nulls_in_an_aware_column(tmp_path):
    df = pl.DataFrame({"t": [UTC_NOON, None]}).with_columns(
        pl.col("t").dt.convert_time_zone("Asia/Jakarta")
    )
    assert _values(tmp_path, df, na_rep="-") == [[LOCAL], ["-"]]


def test_caller_frame_is_not_modified(tmp_path):
    frames = _frames()
    _values(tmp_path, frames["polars"])
    _values(tmp_path, frames["pandas-tz-dtype"])
    assert frames["polars"].schema["t"].time_zone == "Asia/Jakarta"
    assert str(frames["pandas-tz-dtype"]["t"].dtype.tz) == "Asia/Jakarta"


def test_through_write_worksheets_and_fastexcel(tmp_path):
    frame = _frames()["polars"]
    path = tmp_path / "multi.xlsx"
    write_worksheets([("A", frame), ("B", frame)], str(path))
    assert openpyxl.load_workbook(path)["B"]["B2"].value == LOCAL
    path = tmp_path / "fast.xlsx"
    FastExcel(str(path)).sheet("S", frame).save()
    assert openpyxl.load_workbook(path).active["B2"].value == LOCAL


def test_raw_arrow_stream_with_an_aware_column_raises(tmp_path):
    table = pa.table({"id": [1], "t": pa.array([UTC_NOON], type=pa.timestamp("us", tz="Asia/Jakarta"))})
    with pytest.raises(ValueError, match=r"column 't' is timezone-aware \(Asia/Jakarta\).*DataFrame"):
        write_worksheet(table, str(tmp_path / "out.xlsx"))


def test_raw_arrow_stream_fixed_with_local_timestamp(tmp_path):
    import pyarrow.compute as pc

    column = pa.array([UTC_NOON], type=pa.timestamp("us", tz="Asia/Jakarta"))
    assert _values(tmp_path, pa.table({"t": pc.local_timestamp(column)})) == [[LOCAL]]


def test_raw_arrow_stream_aware_column_left_out_by_columns(tmp_path):
    table = pa.table({"id": [1], "t": pa.array([UTC_NOON], type=pa.timestamp("us", tz="UTC"))})
    assert _values(tmp_path, table, columns=["id"]) == [[1]]


def test_raw_arrow_stream_naive_is_fine(tmp_path):
    table = pa.table({"t": pa.array([dt.datetime(2026, 1, 1, 19)], type=pa.timestamp("us"))})
    assert _values(tmp_path, table) == [[LOCAL]]


def test_csv_is_unchanged():
    buf = io.BytesIO()
    write_csv(_frames()["polars"], buf)
    assert buf.getvalue() == b"id,t\n1,2026-01-01T12:00:00+00:00\n"


# --- the helper itself ----------------------------------------------------


@pytest.mark.parametrize(
    "frame",
    [
        pl.DataFrame({"a": [1], "t": [dt.datetime(2026, 1, 1)]}),
        pd.DataFrame({"a": [1], "t": [dt.datetime(2026, 1, 1)]}),
    ],
    ids=["polars", "pandas"],
)
def test_naive_frame_comes_back_as_the_same_object(frame):
    assert local_wall_clock(frame) is frame


def test_only_aware_columns_are_rewritten():
    df = pd.DataFrame(
        {"keep": [dt.datetime(2026, 1, 1)], "t": pd.Series([UTC_NOON]).dt.tz_convert("Asia/Jakarta")}
    )
    out = local_wall_clock(df)
    assert out is not df
    assert out["keep"].dtype == df["keep"].dtype
    assert out["t"].dt.tz is None and out["t"][0] == pd.Timestamp(LOCAL)


def test_other_objects_pass_through_or_report_none():
    table = pa.table({"t": [1]})
    assert local_wall_clock(table) is table
    assert local_wall_clock(table, ["t"]) is None
