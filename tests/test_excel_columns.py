"""`columns=` for the Excel writers.

`write_csv` could select and order its output columns while `write_worksheet`,
`write_worksheets` and `FastExcel.sheet` could not, so choosing columns for
xlsx meant reshaping the records in Python first. The Excel writers now take
the same option with the same rules: an unknown or repeated name raises, and
a key a record carries outside the list is dropped on purpose.
"""

import io

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, Format, write_csv, write_worksheet, write_worksheets

RECORDS = [
    {"id": 1, "name": "a", "score": 1.5},
    {"id": 2, "name": "b", "score": 2.5},
]


def _rows(path, sheet=None):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet] if sheet else wb.active
    return [[c.value for c in row] for row in ws.iter_rows()]


class _Series:
    def __init__(self, values):
        self._values = values

    def tolist(self):
        return list(self._values)


class _Dtype:
    def __init__(self, kind):
        self.kind = kind


class _NoArrowFrame:
    """Pandas-shaped with no Arrow stream, so the per-column writer runs."""

    def __init__(self, data, kinds):
        self._data = data
        self.columns = list(data)
        self.dtypes = [_Dtype(k) for k in kinds]

    def __len__(self):
        return len(next(iter(self._data.values())))

    def __getitem__(self, key):
        return _Series(self._data[key])


def _sources():
    pl = pytest.importorskip("polars")
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    data = {k: [r[k] for r in RECORDS] for k in RECORDS[0]}
    return {
        "records": RECORDS,
        "generator": (dict(r) for r in RECORDS),
        "polars": pl.DataFrame(data),
        "pandas": pd.DataFrame(data),
        "no-arrow": _NoArrowFrame(data, ["i", "O", "f"]),
    }


SOURCES = ["records", "generator", "polars", "pandas", "no-arrow"]


@pytest.mark.parametrize("source", SOURCES)
def test_selects_and_orders_columns(tmp_path, source):
    path = tmp_path / "out.xlsx"
    write_worksheet(_sources()[source], str(path), columns=["score", "id"])
    assert _rows(path) == [["score", "id"], [1.5, 1], [2.5, 2]]


@pytest.mark.parametrize("source", SOURCES)
def test_unknown_name_raises(tmp_path, source):
    with pytest.raises(ValueError, match="columns: 'nope' is not in the data"):
        write_worksheet(_sources()[source], str(tmp_path / "out.xlsx"), columns=["id", "nope"])


@pytest.mark.parametrize("source", SOURCES)
def test_repeated_name_raises(tmp_path, source):
    with pytest.raises(ValueError, match="columns: 'id' is listed more than once"):
        write_worksheet(_sources()[source], str(tmp_path / "out.xlsx"), columns=["id", "id"])


def test_csv_also_refuses_a_repeated_name():
    with pytest.raises(ValueError, match="listed more than once"):
        write_csv(RECORDS, io.BytesIO(), columns=["id", "id"])


def test_keys_outside_the_list_are_dropped(tmp_path):
    """With `columns` the header is a chosen subset, so a key outside it is
    not an error, unlike a record with an unknown key and no `columns`."""
    records = [{"id": 1, "name": "a"}, {"id": 2, "name": "b", "extra": "x"}]
    path = tmp_path / "out.xlsx"
    write_worksheet(records, str(path), columns=["id"])
    assert _rows(path) == [["id"], [1], [2]]


def test_without_columns_an_unknown_key_still_raises(tmp_path):
    with pytest.raises(ValueError, match="no column 'extra'"):
        write_worksheet([{"id": 1}, {"id": 2, "extra": "x"}], str(tmp_path / "out.xlsx"))


def test_reordered_and_missing_keys_with_columns(tmp_path):
    records = [
        {"id": 1, "name": "a", "score": 1.5},
        {"score": 2.5, "id": 2},
        {"name": "c", "id": 3},
    ]
    path = tmp_path / "out.xlsx"
    write_worksheet(records, str(path), columns=["name", "score"])
    assert _rows(path) == [["name", "score"], ["a", 1.5], [None, 2.5], ["c", None]]


def test_shared_key_objects_take_the_fast_path_correctly(tmp_path):
    """Records built in one comprehension share their key objects, so every
    row after the first is placed by pointer comparison. The selection has to
    map those positions, not just pass them through."""
    records = [{"a": i, "b": i * 10, "c": i * 100} for i in range(50)]
    path = tmp_path / "out.xlsx"
    write_worksheet(records, str(path), columns=["c", "a"])
    rows = _rows(path)
    assert rows[0] == ["c", "a"]
    assert rows[1:] == [[i * 100, i] for i in range(50)]


def test_unsupported_arrow_column_can_be_left_out(tmp_path):
    pl = pytest.importorskip("polars")
    df = pl.DataFrame({"id": [1], "tags": [[1, 2]]})
    path = tmp_path / "out.xlsx"
    write_worksheet(df, str(path), columns=["id"])
    assert _rows(path) == [["id"], [1]]
    with pytest.raises(ValueError, match="column 'tags'"):
        write_worksheet(df, str(path), columns=["tags"])


def test_options_keyed_by_name_follow_the_selection(tmp_path):
    path = tmp_path / "out.xlsx"
    write_worksheet(
        RECORDS,
        str(path),
        columns=["score", "id"],
        column_formats={"score": Format().set_num_format("0.000")},
        formula_columns={"double": "=A{row}*2"},
    )
    ws = openpyxl.load_workbook(path).active
    assert [c.value for c in ws[1]] == ["score", "id", "double"]
    assert ws["A2"].number_format == "0.000"
    assert ws["C2"].value == "=A2*2"


def test_write_worksheets_per_sheet_and_general(tmp_path):
    path = tmp_path / "out.xlsx"
    write_worksheets(
        [("One", RECORDS), ("Two", RECORDS)],
        str(path),
        columns={"general": ["id"], "Two": ["name", "id"]},
    )
    assert _rows(path, "One") == [["id"], [1], [2]]
    assert _rows(path, "Two") == [["name", "id"], ["a", 1], ["b", 2]]


def test_fastexcel_sheet_columns(tmp_path):
    path = tmp_path / "out.xlsx"
    FastExcel(str(path)).sheet("S", RECORDS, columns=["name"]).save()
    assert _rows(path) == [["name"], ["a"], ["b"]]


def test_fastexcel_constructor_columns_now_reach_xlsx(tmp_path):
    """The constructor's `columns` used to be CSV-only and was silently
    ignored for xlsx. It is now the default for every sheet, and a sheet's
    own `columns` wins."""
    path = tmp_path / "out.xlsx"
    FastExcel(str(path), columns=["id"]).sheet("A", RECORDS).sheet(
        "B", RECORDS, columns=["score"]
    ).save()
    assert _rows(path, "A") == [["id"], [1], [2]]
    assert _rows(path, "B") == [["score"], [1.5], [2.5]]


def test_fastexcel_sheet_columns_reach_csv_without_a_warning(tmp_path, recwarn):
    path = tmp_path / "out.csv"
    FastExcel(str(path)).sheet("S", RECORDS, columns=["name", "id"]).save()
    assert path.read_text() == "name,id\na,1\nb,2\n"
    assert not [w for w in recwarn if "ignores Excel-only options" in str(w.message)]


def test_fastexcel_sheet_columns_override_the_constructor_for_csv(tmp_path):
    path = tmp_path / "out.csv"
    FastExcel(str(path), columns=["id"]).sheet("S", RECORDS, columns=["name"]).save()
    assert path.read_text() == "name\na\nb\n"
