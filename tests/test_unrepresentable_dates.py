"""Dates a writer cannot store say so, and say where they came from.

CSV has no calendar floor, so 1850 is a perfectly good CSV date. What it
cannot write is a year outside 0..=9999, which a plain ISO-8601 date has no
digits for. polars can hold such a date, and so can a pandas column with a
non-nanosecond unit. The Arrow CSV path used to write an empty field for it.

The object paths had a worse version of the same hole: a pandas `Timestamp`
past 9999 keeps a placeholder year in its C-level datetime fields, which is
what the writers read, so `Timestamp("10500-03-04")` came out as 1970-03-04.
Its `year` property is right, and subclasses are now checked against it.
"""

import datetime as dt
import io

import pytest

from rustpy_xlsxwriter import write_csv

pl = pytest.importorskip("polars")
pd = pytest.importorskip("pandas")
np = pytest.importorskip("numpy")
pytest.importorskip("pyarrow")


def _csv(data, **kwargs):
    buf = io.BytesIO()
    write_csv(data, buf, **kwargs)
    return buf.getvalue().decode().splitlines()


def _far_timestamp():
    """A pandas Timestamp in the year 10500."""
    return pd.Series(np.array(["10500-03-04T05:06:07"], dtype="datetime64[s]"))[0]


class _NoArrowFrame:
    """Pandas-shaped with no usable Arrow stream: the `.values` path runs."""

    def __init__(self, df):
        self.columns = list(df.columns)
        self.values = df.values
        self.dtypes = df.dtypes

    def __arrow_c_stream__(self, requested_schema=None):
        raise ImportError("pyarrow is not installed")


ISO_RANGE = "is outside the years 0 to 9999 that an ISO-8601 date can show"


# --- Arrow CSV ------------------------------------------------------------


def test_arrow_date_past_9999_raises_with_the_expanded_date():
    df = pl.Series("far", [10_500 * 366], dtype=pl.Int32).cast(pl.Date).to_frame()
    with pytest.raises(ValueError, match=rf"column 'far': \+12491-10-11 {ISO_RANGE}"):
        write_csv(df, io.BytesIO())


def test_arrow_timestamp_past_9999_raises():
    df = pd.DataFrame({"t": pd.Series(np.array(["10500-03-04T05:06:07"], dtype="datetime64[s]"))})
    with pytest.raises(ValueError, match=rf"column 't': \+10500-03-04T05:06:07 {ISO_RANGE}"):
        write_csv(df, io.BytesIO())


def test_arrow_date_before_year_0_raises():
    df = pl.Series("old", [-800_000], dtype=pl.Int32).cast(pl.Date).to_frame()
    with pytest.raises(ValueError, match=rf"column 'old': -\d{{4}}-\d\d-\d\d {ISO_RANGE}"):
        write_csv(df, io.BytesIO())


def test_arrow_error_says_earlier_versions_wrote_an_empty_field():
    df = pl.Series("far", [10_500 * 366], dtype=pl.Int32).cast(pl.Date).to_frame()
    with pytest.raises(ValueError, match="empty field without an error"):
        write_csv(df, io.BytesIO())


def test_arrow_far_date_left_out_by_columns():
    df = pl.DataFrame({"id": [1]}).with_columns(
        pl.Series("far", [10_500 * 366], dtype=pl.Int32).cast(pl.Date)
    )
    assert _csv(df, columns=["id"]) == ["id", "1"]


# --- object paths ---------------------------------------------------------


def test_records_far_timestamp_raises_naming_record_and_key():
    with pytest.raises(ValueError, match=rf"record 0, key 'born': 10500-03-04T05:06:07 {ISO_RANGE}"):
        write_csv([{"id": 1, "born": _far_timestamp()}], io.BytesIO())


def test_records_far_timestamp_on_the_by_name_path():
    rows = [{"id": 1, "born": pd.Timestamp("2026-01-01")}, {"born": _far_timestamp(), "id": 2}]
    with pytest.raises(ValueError, match="record 1, key 'born'"):
        write_csv(rows, io.BytesIO())


@pytest.mark.parametrize("columns", [None, ["born"]])
def test_values_path_far_timestamp_raises_naming_the_column(columns):
    frame = _NoArrowFrame(pd.DataFrame({"x": ["a"], "born": [_far_timestamp()]}))
    with pytest.raises(ValueError, match=rf"column 'born': 10500-03-04T05:06:07 {ISO_RANGE}"):
        write_csv(frame, io.BytesIO(), columns=columns)


def test_object_error_says_earlier_versions_wrote_a_wrong_year():
    with pytest.raises(ValueError, match="wrong year without an error"):
        write_csv([{"born": _far_timestamp()}], io.BytesIO())


def test_polars_to_list_cannot_produce_a_far_year():
    """The polars non-Arrow path goes through `to_list`, which refuses such
    a date itself, so it cannot reach the writer silently."""
    far = pl.Series("far", [10_500 * 366], dtype=pl.Int32).cast(pl.Date)
    with pytest.raises(ValueError, match="out of range"):
        far.to_list()


# --- what must keep working -----------------------------------------------


def test_pre_1900_is_still_fine_in_csv():
    old = dt.datetime(1850, 1, 1, 12)
    assert _csv([{"t": old}]) == ["t", "1850-01-01T12:00:00"]
    assert _csv(pl.DataFrame({"t": [old]})) == ["t", "1850-01-01T12:00:00"]
    assert _csv(pl.DataFrame({"d": [dt.date(1850, 3, 4)]})) == ["d", "1850-03-04"]


def test_year_1_is_fine_in_csv():
    assert _csv([{"d": dt.date(1, 1, 1)}]) == ["d", "0001-01-01"]
    assert _csv(pl.DataFrame({"d": [dt.date(1, 1, 1)]})) == ["d", "0001-01-01"]


def test_in_range_timestamps_are_unchanged():
    ts = pd.Timestamp("2026-01-01 12:00:00.5")
    assert _csv([{"t": ts}]) == ["t", "2026-01-01T12:00:00.500000"]
    frame = _NoArrowFrame(pd.DataFrame({"t": [ts]}))
    assert _csv(frame) == ["t", "2026-01-01T12:00:00.500000"]
