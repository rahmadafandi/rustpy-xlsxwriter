"""Errors that name what went wrong and where.

A formula column could reuse a data column's header, which left every option
keyed by column name pointing at two columns, and no complaint. And running
out of rows produced the crate's bare "Row or column exceeds Excel's allowed
limits", with no sheet and no row, which in a multi-sheet workbook leaves the
caller guessing.

The row-limit tests move the header near the bottom of the sheet rather than
writing a million rows: `header_row=1_048_573` leaves room for two records.
"""

import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

NEAR_END = 1_048_573  # header on row 1,048,574; data rows 1,048,575-1,048,576


class _Series:
    def __init__(self, values):
        self._values = values

    def tolist(self):
        return list(self._values)


class _Dtype:
    kind = "i"


class _NoArrowFrame:
    """Pandas-shaped with no Arrow stream, so the per-column writer runs."""

    def __init__(self, data):
        self._data = data
        self.columns = list(data)
        self.dtypes = [_Dtype() for _ in data]

    def __len__(self):
        return len(next(iter(self._data.values())))

    def __getitem__(self, key):
        return _Series(self._data[key])


def _frames(data):
    """The same data as records, polars, pandas and the no-Arrow writer."""
    pl = pytest.importorskip("polars")
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    n = len(next(iter(data.values())))
    return {
        "records": [{k: v[i] for k, v in data.items()} for i in range(n)],
        "polars": pl.DataFrame(data),
        "pandas": pd.DataFrame(data),
        "no-arrow": _NoArrowFrame(data),
    }


SOURCES = ["records", "polars", "pandas", "no-arrow"]


# --- formula column vs data column ---------------------------------------


@pytest.mark.parametrize("source", SOURCES)
def test_formula_column_may_not_reuse_a_data_header(tmp_path, source):
    data = _frames({"a": [1], "b": [2]})[source]
    with pytest.raises(ValueError, match="formula_columns: 'a' is already a data column"):
        write_worksheet(data, str(tmp_path / "out.xlsx"), formula_columns={"a": "=1+1"})


def test_formula_collision_through_the_builder(tmp_path):
    with pytest.raises(ValueError, match="'b' is already a data column"):
        FastExcel(str(tmp_path / "out.xlsx")).sheet(
            "S", [{"a": 1, "b": 2}], formula_columns={"total": "=A{row}", "b": "=1"}
        ).save()


def test_distinct_formula_header_still_works(tmp_path):
    import openpyxl

    path = tmp_path / "out.xlsx"
    write_worksheet([{"a": 1, "b": 2}], str(path), formula_columns={"sum": "=A{row}+B{row}"})
    ws = openpyxl.load_workbook(path).active
    assert [c.value for c in ws[1]] == ["a", "b", "sum"]
