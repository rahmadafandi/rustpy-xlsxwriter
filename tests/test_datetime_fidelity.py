"""What survives of a datetime, a time and a timezone on the way out.

Three losses lived in the same two functions. Both writers read only year to
second, so microseconds vanished. A `datetime.time` fell through to `str()`
and landed in Excel as text, which cannot be formatted, summed or compared.
And CSV dropped the UTC offset of an aware datetime, so a UTC export read
back as naive local time.

The `time` column in `conftest.py` comes from `Faker.time()`, which returns a
string, so before these tests no real `datetime.time` had reached the writer.
"""

import datetime as dt
import io
import zipfile
from datetime import timedelta, timezone

import openpyxl
import pytest

from rustpy_xlsxwriter import Format, write_csv, write_worksheet

UTC_NOON = dt.datetime(2026, 1, 1, 12, tzinfo=timezone.utc)
JKT_NOON = dt.datetime(2026, 1, 1, 12, tzinfo=timezone(timedelta(hours=7)))
HALF_SECOND = dt.datetime(2026, 1, 1, 12, 0, 0, 500000)
BEFORE_EPOCH = dt.datetime(1969, 12, 31, 23, 59, 59, 250000)
T1 = dt.time(13, 30)
T2 = dt.time(1, 2, 3, 500000)


def _csv(data, **kwargs):
    buf = io.BytesIO()
    write_csv(data, buf, **kwargs)
    return buf.getvalue().decode().splitlines()


def _cells(tmp_path, data, **kwargs):
    path = tmp_path / "out.xlsx"
    write_worksheet(data, str(path), **kwargs)
    ws = openpyxl.load_workbook(path).active
    return [list(row) for row in ws.iter_rows(min_row=2)]


def _sheet_xml(tmp_path, data, **kwargs):
    path = tmp_path / "raw.xlsx"
    write_worksheet(data, str(path), **kwargs)
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode()


# --- B1: timezone offsets ------------------------------------------------


@pytest.mark.parametrize(
    "value, text",
    [
        (UTC_NOON, "2026-01-01T12:00:00+00:00"),
        (JKT_NOON, "2026-01-01T12:00:00+07:00"),
        (
            dt.datetime(2026, 1, 1, 12, tzinfo=timezone(-timedelta(hours=5, minutes=30))),
            "2026-01-01T12:00:00-05:30",
        ),
        (
            dt.datetime(2026, 1, 1, 12, tzinfo=timezone(timedelta(hours=1, seconds=15))),
            "2026-01-01T12:00:00+01:00:15",
        ),
        (dt.datetime(2026, 1, 1, 12), "2026-01-01T12:00:00"),
    ],
    ids=["utc", "plus-seven", "minus-half-hour", "offset-seconds", "naive"],
)
def test_csv_writes_the_offset_of_an_aware_datetime(value, text):
    assert _csv([{"t": value}]) == ["t", text]


def test_csv_offset_matches_isoformat_and_round_trips():
    for value in (UTC_NOON, JKT_NOON):
        (line,) = _csv([{"t": value}], header=False)
        assert line == value.isoformat()
        assert dt.datetime.fromisoformat(line) == value


def test_csv_keeps_two_instants_apart():
    """The reported case: seven hours apart, collapsed into one value."""
    lines = _csv([{"t": UTC_NOON}, {"t": JKT_NOON}], header=False)
    assert len(set(lines)) == 2
    assert [dt.datetime.fromisoformat(x) for x in lines] == [UTC_NOON, JKT_NOON]


def test_csv_aware_time_carries_its_offset():
    t = dt.time(9, 15, tzinfo=timezone(timedelta(hours=2)))
    assert _csv([{"t": t}]) == ["t", "09:15:00+02:00"]


def test_excel_keeps_the_wall_clock_of_an_aware_datetime(tmp_path):
    """Excel has no timezone, so an aware datetime is written as its own
    local wall clock and the offset is dropped."""
    rows = _cells(tmp_path, [{"t": UTC_NOON}, {"t": JKT_NOON}])
    assert [r[0].value for r in rows] == [dt.datetime(2026, 1, 1, 12)] * 2


def _aware_frames():
    pl = pytest.importorskip("polars")
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    polars = pl.DataFrame({"t": [dt.datetime(2026, 1, 1, 12)]}).with_columns(
        pl.col("t").dt.replace_time_zone("UTC").dt.convert_time_zone("Asia/Jakarta")
    )
    pandas = pd.DataFrame(
        {"t": pd.to_datetime(["2026-01-01 12:00"]).tz_localize("UTC").tz_convert("Asia/Jakarta")}
    )
    return {"polars": polars, "pandas": pandas}


@pytest.mark.parametrize("lib", ["polars", "pandas"])
def test_csv_writes_an_aware_arrow_column_as_a_utc_instant(lib):
    """Arrow stores an aware column as UTC instants plus a zone name. Without
    a timezone database the zone's local time is not known, so the column is
    written in UTC and labelled as such, which names the same instant."""
    (line,) = _csv(_aware_frames()[lib], header=False)
    assert line == "2026-01-01T12:00:00+00:00"
    assert dt.datetime.fromisoformat(line) == UTC_NOON


@pytest.mark.parametrize("lib", ["polars", "pandas"])
def test_excel_writes_an_aware_arrow_column_as_local_wall_clock(tmp_path, lib):
    """12:00 UTC shown in Asia/Jakarta is 19:00, as the frame prints it; see
    test_timezones_xlsx.py for the rest."""
    rows = _cells(tmp_path, _aware_frames()[lib])
    assert rows[0][0].value == dt.datetime(2026, 1, 1, 19)


# --- B2: fractions of a second -------------------------------------------


def test_csv_keeps_microseconds():
    assert _csv([{"t": HALF_SECOND}, {"t": dt.datetime(2026, 1, 1, 0, 0, 0, 7)}]) == [
        "t",
        "2026-01-01T12:00:00.500000",
        "2026-01-01T00:00:00.000007",
    ]


def test_excel_keeps_microseconds(tmp_path):
    rows = _cells(tmp_path, [{"t": HALF_SECOND}])
    assert rows[0][0].value == HALF_SECOND


def test_excel_serial_holds_the_fraction(tmp_path):
    xml = _sheet_xml(tmp_path, [{"t": dt.datetime(2026, 1, 1, 0, 0, 0, 1000)}])
    # 1 ms is 1/86_400_000 of a day: 46023 + 1.157e-8.
    assert "<v>46023.0000000115" in xml


def test_excel_accepts_seconds_past_the_millisecond_limit(tmp_path):
    """`ExcelDateTime::and_hms` refuses anything above 59.999 s; the fraction
    is added to the serial instead, so 59.9995 s still writes."""
    value = dt.datetime(2026, 1, 1, 12, 0, 59, 999500)
    (row,) = _cells(tmp_path, [{"t": value}])
    assert abs(row[0].value - value) < timedelta(milliseconds=1)


@pytest.mark.parametrize("unit", ["ms", "us", "ns"])
def test_arrow_timestamp_keeps_its_fraction(tmp_path, unit):
    pl = pytest.importorskip("polars")
    df = pl.DataFrame({"t": [HALF_SECOND]}, schema={"t": pl.Datetime(unit)})
    assert _csv(df) == ["t", "2026-01-01T12:00:00.500000"]
    assert _cells(tmp_path, df)[0][0].value == HALF_SECOND


def test_arrow_nanoseconds_reach_csv():
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    df = pd.DataFrame({"t": [pd.Timestamp("2026-01-01 12:00:00.000000001")]})
    assert _csv(df) == ["t", "2026-01-01T12:00:00.000000001"]


def test_arrow_timestamp_before_1970(tmp_path):
    """Truncating division put a moment just before the epoch on the wrong
    side of it: 1969-12-31T23:59:59.25 came out as 1970-01-01T00:00:00."""
    pl = pytest.importorskip("polars")
    df = pl.DataFrame({"t": [BEFORE_EPOCH, dt.datetime(1950, 6, 1, 8, 30)]})
    assert _csv(df) == ["t", "1969-12-31T23:59:59.250000", "1950-06-01T08:30:00"]
    assert [r[0].value for r in _cells(tmp_path, df)] == [
        BEFORE_EPOCH,
        dt.datetime(1950, 6, 1, 8, 30),
    ]


def test_date64_before_1970():
    pa = pytest.importorskip("pyarrow")
    pd = pytest.importorskip("pandas")
    table = pa.table({"d": pa.array([-1], type=pa.date64())})
    df = table.to_pandas(types_mapper=pd.ArrowDtype)
    assert _csv(df) == ["d", "1969-12-31T23:59:59.999000"]


# --- B3: datetime.time ---------------------------------------------------


def test_excel_writes_time_as_a_number_with_a_time_format(tmp_path):
    rows = _cells(tmp_path, [{"t": T1}, {"t": T2}])
    assert [r[0].value for r in rows] == [T1, T2]
    assert [r[0].number_format for r in rows] == ["hh:mm:ss"] * 2


def test_excel_time_is_not_inline_text(tmp_path):
    """The raw cell the bug report showed: `t="inlineStr"` with 13:30:00."""
    xml = _sheet_xml(tmp_path, [{"t": T1}])
    assert '<c r="A2" s="1"><v>0.5625</v></c>' in xml


def test_csv_writes_time_as_iso():
    assert _csv([{"t": T1}, {"t": T2}]) == ["t", "13:30:00", "01:02:03.500000"]


def test_time_column_with_a_leading_none(tmp_path):
    """The first row decides the cached type; a None there must not stop the
    times after it from being recognised."""
    rows = _cells(tmp_path, [{"t": None}, {"t": T1}, {"t": T2}])
    assert [r[0].value for r in rows] == [None, T1, T2]


def test_cached_time_column_still_catches_other_types(tmp_path):
    rows = _cells(tmp_path, [{"t": T1}, {"t": "late"}, {"t": T2}])
    assert [r[0].value for r in rows] == [T1, "late", T2]
    assert _csv([{"t": T1}, {"t": "late"}]) == ["t", "13:30:00", "late"]


def test_time_next_to_datetime_and_date(tmp_path):
    record = {"dt": dt.datetime(2026, 1, 1, 13, 30), "d": dt.date(2026, 1, 1), "t": T1}
    (row,) = _cells(tmp_path, [record])
    assert [c.value for c in row] == [dt.datetime(2026, 1, 1, 13, 30), dt.datetime(2026, 1, 1), T1]
    assert row[2].number_format == "hh:mm:ss"


def test_column_format_overrides_the_time_format(tmp_path):
    fmt = Format().set_num_format("hh:mm:ss.000")
    (row,) = _cells(tmp_path, [{"t": T2}], column_formats={"t": fmt})
    assert row[0].number_format == "hh:mm:ss.000"


def test_banded_rows_keep_the_time_format(tmp_path):
    rows = _cells(tmp_path, [{"t": T1}, {"t": T2}, {"t": T1}], banded_rows="#EEEEEE")
    assert [r[0].number_format for r in rows] == ["hh:mm:ss"] * 3
    assert rows[1][0].fill.fgColor.rgb.endswith("EEEEEE")


# --- dates Excel's calendar cannot hold ----------------------------------


def _early_frames():
    pl = pytest.importorskip("polars")
    pd = pytest.importorskip("pandas")
    pa = pytest.importorskip("pyarrow")
    return {
        "polars-datetime": pl.DataFrame(
            {"ok": [1, 2], "born": [dt.datetime(2000, 1, 1), dt.datetime(1850, 1, 1, 12)]}
        ),
        "polars-date": pl.DataFrame({"ok": [1], "born": [dt.date(1850, 3, 4)]}),
        "pandas-date32": pd.DataFrame(
            {"ok": [1], "born": pd.array([dt.date(1850, 3, 4)], dtype=pd.ArrowDtype(pa.date32()))}
        ),
        "pandas-datetime": pd.DataFrame({"ok": [1], "born": [pd.Timestamp("1850-01-01 12:00")]}),
    }


@pytest.mark.parametrize(
    "label", ["polars-datetime", "polars-date", "pandas-date32", "pandas-datetime"]
)
def test_dataframe_date_before_1900_raises_naming_column_and_value(tmp_path, label):
    """A DataFrame used to get an empty cell here while a list of dicts
    raised. Excel's serial has no room below its epoch, so both raise."""
    with pytest.raises(ValueError, match=r"column 'born': 1850-0\d-0\d.* outside Excel's calendar"):
        write_worksheet(_early_frames()[label], str(tmp_path / "out.xlsx"))


def test_dataframe_date_past_9999_raises(tmp_path):
    pl = pytest.importorskip("polars")
    df = pl.Series("far", [10_500 * 366], dtype=pl.Int32).cast(pl.Date).to_frame()
    with pytest.raises(ValueError, match=r"column 'far': \+12491-10-11 is outside Excel's calendar"):
        write_worksheet(df, str(tmp_path / "out.xlsx"))


def test_records_date_before_1900_still_raises(tmp_path):
    with pytest.raises(ValueError, match="1850"):
        write_worksheet([{"born": dt.datetime(1850, 1, 1)}], str(tmp_path / "out.xlsx"))


def test_csv_still_writes_dates_before_1900():
    pl = pytest.importorskip("polars")
    assert _csv(pl.DataFrame({"born": [dt.date(1850, 3, 4)]})) == ["born", "1850-03-04"]
