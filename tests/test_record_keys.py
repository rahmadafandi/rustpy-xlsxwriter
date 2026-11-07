"""Records are matched to the header by key name, not by position.

The header comes from the first record, since a generator gives no way to look
ahead. Records after it used to be written in their own dict's iteration order,
which is the same order for uniform records but puts the values under the wrong
headers when the orders differ.

Both writers keep a fast path for records carrying the first one's key objects
in the same order, so these tests build dicts in different orders to force the
slow path.
"""

import openpyxl
import pytest

from rustpy_xlsxwriter import write_csv, write_worksheet


def _cells(path):
    ws = openpyxl.load_workbook(path).active
    return [[c.value for c in row] for row in ws.iter_rows()]


# --- Excel ------------------------------------------------------------------


def test_reordered_keys_land_under_their_own_header(tmp_path):
    path = str(tmp_path / "reordered.xlsx")
    write_worksheet([{"a": 1, "b": 2}, {"b": 20, "a": 10}], path)
    assert _cells(path) == [["a", "b"], [1, 2], [10, 20]]


def test_a_missing_key_leaves_its_cell_empty(tmp_path):
    """Not shifted: every column after it would otherwise be off by one."""
    path = str(tmp_path / "missing.xlsx")
    write_worksheet([{"a": 1, "b": 2}, {"b": 20}], path)
    assert _cells(path) == [["a", "b"], [1, 2], [None, 20]]


def test_an_unknown_key_raises_and_names_the_column(tmp_path):
    with pytest.raises(ValueError, match=r"record 1: no column 'c'"):
        write_worksheet(
            [{"a": 1, "b": 2}, {"a": 10, "b": 20, "c": 30}],
            str(tmp_path / "extra.xlsx"),
        )


def test_the_error_names_the_header_it_was_matched_against(tmp_path):
    with pytest.raises(ValueError, match=r"first record \(a, b\)"):
        write_worksheet([{"a": 1, "b": 2}, {"c": 3}], str(tmp_path / "e.xlsx"))


def test_reordering_survives_a_column_format(tmp_path):
    """The format is attached to the column, so a misplaced value would show."""
    path = str(tmp_path / "fmt.xlsx")
    write_worksheet(
        [{"n": 1.5, "s": "x"}, {"s": "y", "n": 2.5}],
        path,
        float_format="0.000",
    )
    ws = openpyxl.load_workbook(path).active
    assert [c.value for c in ws[3]] == [2.5, "y"]
    assert ws.cell(3, 1).number_format == "0.000"


# --- CSV --------------------------------------------------------------------


def test_csv_reordered_keys_keep_the_header_order(tmp_path):
    path = tmp_path / "reordered.csv"
    write_csv([{"a": 1, "b": 2}, {"b": 20, "a": 10}], str(path))
    assert path.read_text().strip() == "a,b\n1,2\n10,20"


def test_csv_missing_key_writes_an_empty_field(tmp_path):
    path = tmp_path / "missing.csv"
    write_csv([{"a": 1, "b": 2}, {"b": 20}], str(path))
    assert path.read_text().strip() == "a,b\n1,2\n,20"


def test_csv_unknown_key_raises_rather_than_writing_a_ragged_line(tmp_path):
    with pytest.raises(ValueError, match=r"record 1: no column 'c'"):
        write_csv(
            [{"a": 1, "b": 2}, {"a": 10, "b": 20, "c": 30}],
            str(tmp_path / "extra.csv"),
        )


def test_csv_every_line_has_as_many_fields_as_the_header(tmp_path):
    """Every line carries as many fields as the header."""
    path = tmp_path / "shape.csv"
    write_csv([{"a": 1, "b": 2}, {"b": 20, "a": 10}, {"a": 100}], str(path))
    lines = path.read_text().strip().splitlines()
    assert {line.count(",") for line in lines} == {1}


def test_csv_selected_columns_still_drop_the_rest(tmp_path):
    """With `columns` the header is a chosen subset, so an unlisted key is
    dropped rather than rejected."""
    path = tmp_path / "selected.csv"
    write_csv(
        [{"a": 1, "b": 2, "c": 3}, {"c": 30, "a": 10, "b": 20}],
        str(path),
        columns=["a", "c"],
    )
    assert path.read_text().strip() == "a,c\n1,3\n10,30"


def test_csv_fast_and_slow_paths_agree(tmp_path):
    """Same data, one built to hit the pointer fast path and one to miss it."""
    fast = [{"a": 1, "b": "x"}, {"a": 2, "b": "y"}]
    slow = [{"a": 1, "b": "x"}, {"b": "y", "a": 2}]

    fast_path, slow_path = tmp_path / "fast.csv", tmp_path / "slow.csv"
    write_csv(fast, str(fast_path))
    write_csv(slow, str(slow_path))
    assert fast_path.read_bytes() == slow_path.read_bytes()
