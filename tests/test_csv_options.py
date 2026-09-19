"""``bom``, ``columns`` and ``header`` on the CSV writer.

CSV has four input paths — Arrow, Polars, Pandas and Records — and an option
that only works on the one the test happens to hit is worse than no option, so
each is exercised separately. Real DataFrames take the Arrow path; the two
stubs below fail their Arrow stream the way pandas without pyarrow does, which
is the only way to reach the other two branches.
"""

import io

import pytest

from rustpy_xlsxwriter import FastExcel, write_csv

ROWS = [
    {"name": "a", "score": 1.5, "count": 1},
    {"name": "b", "score": 2.5, "count": 2},
]
BOM = b"\xef\xbb\xbf"


def _csv(data, **kwargs):
    buf = io.BytesIO()
    write_csv(data, buf, **kwargs)
    return buf.getvalue()


class _Series:
    def __init__(self, values):
        self._values = list(values)

    def to_list(self):
        return self._values


class _PandasLike:
    """Reaches the ``.values`` branch: Arrow advertised, Arrow unavailable."""

    def __init__(self, data):
        self._data = data
        self.columns = list(data)

    def __arrow_c_stream__(self, requested_schema=None):
        raise ImportError("`Import pyarrow` failed.")

    @property
    def values(self):
        return [list(row) for row in zip(*self._data.values())]


class _PolarsLike(_PandasLike):
    """Same, but with the Polars accessor so the other branch is exercised."""

    def get_column(self, name):
        return _Series(self._data[name])

    def __len__(self):
        return len(next(iter(self._data.values())))

    @property
    def values(self):  # pragma: no cover - must not be reached
        raise AssertionError("polars frames must go through get_column")


DATA = {"name": ["a", "b"], "score": [1.5, 2.5], "count": [1, 2]}


# --- BOM --------------------------------------------------------------------


def test_bom_is_off_by_default():
    assert not _csv(ROWS).startswith(BOM)


def test_bom_prefixes_the_file_once():
    out = _csv(ROWS, bom=True)
    assert out.startswith(BOM)
    assert out.count(BOM) == 1
    # Everything after the marker is the file it would have been.
    assert out[len(BOM) :] == _csv(ROWS)


def test_bom_survives_the_arrow_fallback():
    """The fallback resets the buffer; it must not drop or double the marker."""
    out = _csv(_PandasLike(DATA), bom=True)
    assert out.startswith(BOM)
    assert out.count(BOM) == 1


# --- header -----------------------------------------------------------------


def test_header_can_be_suppressed():
    assert _csv(ROWS, header=False) == b"a,1.5,1\nb,2.5,2\n"


@pytest.mark.parametrize("frame", [_PandasLike, _PolarsLike], ids=["pandas", "polars"])
def test_header_suppressed_on_fallback_paths(frame):
    assert _csv(frame(DATA), header=False) == b"a,1.5,1\nb,2.5,2\n"


# --- columns ----------------------------------------------------------------


def test_columns_selects_and_reorders():
    assert _csv(ROWS, columns=["count", "name"]) == b"count,name\n1,a\n2,b\n"


def test_columns_can_select_one():
    assert _csv(ROWS, columns=["score"]) == b"score\n1.5\n2.5\n"


def test_columns_with_header_suppressed():
    assert _csv(ROWS, columns=["count", "name"], header=False) == b"1,a\n2,b\n"


@pytest.mark.parametrize("frame", [_PandasLike, _PolarsLike], ids=["pandas", "polars"])
def test_columns_on_fallback_paths(frame):
    assert _csv(frame(DATA), columns=["count", "name"]) == b"count,name\n1,a\n2,b\n"


def test_columns_on_the_arrow_path():
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    df = pd.DataFrame(DATA)
    assert _csv(df, columns=["count", "name"]) == b"count,name\n1,a\n2,b\n"


def test_columns_on_polars():
    pl = pytest.importorskip("polars")
    assert _csv(pl.DataFrame(DATA), columns=["count", "name"]) == b"count,name\n1,a\n2,b\n"


def test_a_row_missing_a_selected_key_leaves_the_field_empty():
    """Shifting the remaining columns left would corrupt every later field."""
    rows = [{"a": 1, "b": 2}, {"a": 3}]
    assert _csv(rows, columns=["a", "b"]) == b"a,b\n1,2\n3,\n"


# --- columns: unknown names -------------------------------------------------


def test_unknown_column_raises():
    with pytest.raises(ValueError, match="columns: 'nope' is not in the data"):
        _csv(ROWS, columns=["nope"])


def test_unknown_column_names_what_is_available():
    with pytest.raises(ValueError, match="available: name, score, count"):
        _csv(ROWS, columns=["nope"])


@pytest.mark.parametrize("frame", [_PandasLike, _PolarsLike], ids=["pandas", "polars"])
def test_unknown_column_raises_on_fallback_paths(frame):
    with pytest.raises(ValueError, match="columns: 'nope'"):
        _csv(frame(DATA), columns=["nope"])


def test_unknown_column_raises_on_the_arrow_path():
    """Must not fall through to a slower path and fail there instead."""
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    with pytest.raises(ValueError, match="columns: 'nope'"):
        _csv(pd.DataFrame(DATA), columns=["nope"])


# --- builder ----------------------------------------------------------------


def test_builder_passes_the_options_through():
    buf = io.BytesIO()
    (
        FastExcel(buf, output_format="csv", bom=True, columns=["count", "name"])
        .sheet("S", ROWS)
        .save()
    )
    assert buf.getvalue() == BOM + b"count,name\n1,a\n2,b\n"


def test_builder_tsv_with_options(tmp_path):
    path = tmp_path / "out.tsv"
    FastExcel(path, columns=["name"], header=False).sheet("S", ROWS).save()
    assert path.read_bytes() == b"a\nb\n"


def test_builder_rejects_unknown_column():
    with pytest.raises(ValueError, match="columns: 'nope'"):
        FastExcel(io.BytesIO(), output_format="csv", columns=["nope"]).sheet(
            "S", ROWS
        ).save()
