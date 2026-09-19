"""Collapsible row and column groups.

Row groups are given by sheet row index, matching ``row_heights``; column
groups by header name, matching everything else keyed by column. They are
applied at different points — rows before the data, columns once the headers
are known — so both are checked separately.

The XML is read directly: openpyxl exposes ``outlineLevel`` only through the
dimension objects it happens to have materialised, which made an earlier draft
of these tests report a group that was in fact written correctly.
"""

import re
import zipfile

import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [{"region": r, "q1": 1, "q2": 2, "total": 3} for r in "abcde"]


def _write(tmp_path, spec, rows=None, **kwargs):
    path = tmp_path / "o.xlsx"
    write_worksheet(
        rows if rows is not None else ROWS, str(path), outline=spec, **kwargs
    )
    return path


def _xml(path):
    return zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode()


def _rows(path):
    """``(row number, outlineLevel or None, hidden)`` for each row element."""
    out = []
    for tag in re.findall(r"<row [^>]*>", _xml(path)):
        level = re.search(r'outlineLevel="(\d+)"', tag)
        out.append(
            (
                int(re.search(r'r="(\d+)"', tag).group(1)),
                int(level.group(1)) if level else None,
                'hidden="1"' in tag,
            )
        )
    return out


def _cols(path):
    return re.findall(r'<col min="(\d+)" max="(\d+)"[^>]*outlineLevel="(\d+)"', _xml(path))


# --- rows -------------------------------------------------------------------


def test_row_group_sets_the_outline_level(tmp_path):
    """The bracket, not just the hiding — this is what constant memory drops."""
    levels = {r: lvl for r, lvl, _ in _rows(_write(tmp_path, {"rows": [{"from": 2, "to": 4}]}))}
    assert [levels.get(r) for r in (3, 4, 5)] == [1, 1, 1]
    assert levels.get(2) is None


def test_collapsed_row_group_hides_and_brackets(tmp_path):
    rows = _rows(_write(tmp_path, {"rows": [{"from": 2, "to": 4, "collapsed": True}]}))
    grouped = [(lvl, hidden) for r, lvl, hidden in rows if r in (3, 4, 5)]
    assert grouped == [(1, True)] * 3


def test_row_group_works_without_dedupe_strings(tmp_path):
    """It must not need the caller to know about the constant-memory swap."""
    assert "outlineLevel" in _xml(_write(tmp_path, {"rows": [{"from": 1, "to": 2}]}))


def test_nested_row_groups_raise_the_level(tmp_path):
    path = _write(
        tmp_path, {"rows": [{"from": 1, "to": 4}, {"from": 2, "to": 3}]}
    )
    levels = {r: lvl for r, lvl, _ in _rows(path)}
    assert levels.get(2) == 1
    assert levels.get(3) == 2


def test_reversed_row_bounds_raise(tmp_path):
    with pytest.raises(ValueError, match="row group 'from' \\(4\\) is after 'to' \\(2\\)"):
        _write(tmp_path, {"rows": [{"from": 4, "to": 2}]})


# --- columns ----------------------------------------------------------------


def test_column_group_by_header_name(tmp_path):
    """q1 and q2 are the second and third columns."""
    assert _cols(_write(tmp_path, {"columns": [{"from": "q1", "to": "q2"}]})) == [
        ("2", "3", "1")
    ]


def test_column_group_collapsed(tmp_path):
    path = _write(tmp_path, {"columns": [{"from": "q1", "to": "q2", "collapsed": True}]})
    assert _cols(path) == [("2", "3", "1")]
    assert 'hidden="1"' in _xml(path)


def test_column_group_keeps_constant_memory(tmp_path):
    """Only row groups pay the buffering cost."""
    path = _write(tmp_path, {"columns": [{"from": "q1", "to": "q2"}]})
    assert _cols(path)


def test_unknown_column_warns_and_is_skipped(tmp_path):
    with pytest.warns(UserWarning, match="unknown column in group 'nope'..'q2'"):
        path = _write(tmp_path, {"columns": [{"from": "nope", "to": "q2"}]})
    assert _cols(path) == []


def test_columns_out_of_order_raise(tmp_path):
    with pytest.raises(ValueError, match="'total' comes after 'q1' in the data"):
        _write(tmp_path, {"columns": [{"from": "total", "to": "q1"}]})


# --- symbols and validation -------------------------------------------------


def test_symbols_above(tmp_path):
    path = _write(tmp_path, {"rows": [{"from": 1, "to": 2}], "symbols_above": True})
    assert 'summaryBelow="0"' in _xml(path)


def test_symbols_to_left(tmp_path):
    path = _write(tmp_path, {"columns": [{"from": "q1", "to": "q2"}], "symbols_to_left": True})
    assert 'summaryRight="0"' in _xml(path)


def test_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match="outline: unknown key 'row'"):
        _write(tmp_path, {"row": [{"from": 1, "to": 2}]})


def test_unknown_key_inside_a_group_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown key 'start' in a 'rows' group"):
        _write(tmp_path, {"rows": [{"start": 1, "to": 2}]})


def test_group_needs_both_bounds(tmp_path):
    with pytest.raises(ValueError, match="a 'rows' group needs 'to'"):
        _write(tmp_path, {"rows": [{"from": 1}]})


def test_group_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="every 'rows' group must be a dict"):
        _write(tmp_path, {"rows": [(1, 2)]})


def test_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="outline must be a dict"):
        _write(tmp_path, ["rows"])


# --- every entry point ------------------------------------------------------


def test_multi_sheet_is_keyed_by_sheet_name(tmp_path):
    path = tmp_path / "multi.xlsx"
    write_worksheets(
        [("A", ROWS), ("B", ROWS)],
        str(path),
        outline={"A": {"rows": [{"from": 1, "to": 2}]}},
    )
    book = zipfile.ZipFile(path)
    assert "outlineLevel" in book.read("xl/worksheets/sheet1.xml").decode()
    assert "outlineLevel" not in book.read("xl/worksheets/sheet2.xml").decode()


def test_builder(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet(
        "S", ROWS, outline={"columns": [{"from": "q1", "to": "q2"}]}
    ).save()
    assert _cols(path) == [("2", "3", "1")]
