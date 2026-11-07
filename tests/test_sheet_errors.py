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


# --- row limit -----------------------------------------------------------


@pytest.mark.parametrize("source", SOURCES)
def test_row_limit_names_the_sheet_and_the_record(tmp_path, source):
    data = _frames({"a": [1, 2, 3]})[source]
    with pytest.raises(RuntimeError) as exc:
        write_worksheet(data, str(tmp_path / "out.xlsx"), sheet_name="Data", header_row=NEAR_END)
    message = str(exc.value)
    assert "sheet 'Data'" in message
    assert "record 3" in message
    assert "row 1048577" in message
    assert "room for 2 records" in message


@pytest.mark.parametrize("source", SOURCES)
def test_data_that_exactly_fits_is_written(tmp_path, source):
    data = _frames({"a": [1, 2]})[source]
    write_worksheet(data, str(tmp_path / "out.xlsx"), header_row=NEAR_END)


def test_row_limit_names_the_failing_sheet_of_several(tmp_path):
    sheets = [
        ("Fine", [{"a": 1}, {"a": 2}, {"a": 3}]),
        ("Overflow", [{"a": 1}, {"a": 2}, {"a": 3}]),
    ]
    with pytest.raises(RuntimeError, match="sheet 'Overflow': record 3"):
        write_worksheets(
            sheets, str(tmp_path / "out.xlsx"), header_row={"Overflow": NEAR_END}
        )


def test_column_limit_names_the_sheet(tmp_path):
    record = {f"c{i}": i for i in range(16_385)}
    with pytest.raises(RuntimeError, match="sheet 'Wide' has 16385 columns, past Excel's limit of 16384"):
        write_worksheet([record], str(tmp_path / "out.xlsx"), sheet_name="Wide")


def test_formula_columns_count_towards_the_column_limit(tmp_path):
    record = {f"c{i}": i for i in range(16_384)}
    with pytest.raises(RuntimeError, match="16385 columns"):
        write_worksheet([record], str(tmp_path / "out.xlsx"), formula_columns={"f": "=1"})


# --- an Arrow stream that breaks after rows are written ------------------


class _StreamBreaksLater:
    """Writes some batches, then fails, and also offers the fallback
    accessors. Falling back at that point would write every row again over
    rows already flushed, so the error has to propagate."""

    def __init__(self, data):
        self._data = data
        self.columns = list(data)
        self.dtypes = [_Dtype() for _ in data]

    def __len__(self):
        return len(next(iter(self._data.values())))

    def __getitem__(self, key):
        return _Series(self._data[key])

    def __arrow_c_stream__(self, requested_schema=None):
        import pyarrow as pa

        table = pa.table(self._data)
        batches = table.to_batches(max_chunksize=10)

        def gen():
            yield batches[0]
            raise RuntimeError("stream broke")

        return pa.RecordBatchReader.from_batches(table.schema, gen()).__arrow_c_stream__(
            requested_schema
        )


def test_arrow_failure_after_rows_were_written_propagates(tmp_path):
    pytest.importorskip("pyarrow")
    frame = _StreamBreaksLater({"a": list(range(100))})
    with pytest.raises(RuntimeError, match="stream broke"):
        write_worksheet(frame, str(tmp_path / "out.xlsx"))
