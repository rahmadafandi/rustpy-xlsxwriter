"""Per-row trend charts.

Sparklines live in the x14 extension block, which openpyxl drops on read, so
these read the sheet XML. The part worth checking is the split: one chart per
data row, each reading that row's span, in the column that was left empty.
"""

import re
import zipfile

import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [
    {"name": "a", "q1": 1, "q2": 5, "q3": 3, "q4": 8, "trend": None},
    {"name": "b", "q1": 2, "q2": 1, "q3": 7, "q4": 4, "trend": None},
]


def _write(tmp_path, spec, rows=None, **kwargs):
    path = tmp_path / "s.xlsx"
    write_worksheet(
        rows if rows is not None else ROWS, str(path), sparklines=spec, **kwargs
    )
    return path


def _xml(path):
    return zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode()


def _lines(path):
    """``(source range, target cell)`` for each sparkline."""
    return re.findall(r"<xm:f>(.*?)</xm:f><xm:sqref>(.*?)</xm:sqref>", _xml(path))


def _group(path):
    # The plural <sparklineGroups> container wraps it, so anchor on the
    # attribute boundary rather than the tag prefix.
    m = re.search(r"<x14:sparklineGroup[ >][^>]*>", _xml(path))
    return m.group(0) if m else ""


# --- placement --------------------------------------------------------------


def test_one_chart_per_data_row(tmp_path):
    assert _lines(_write(tmp_path, {"trend": {"from": "q1", "to": "q4"}})) == [
        ("Sheet1!B2:E2", "F2"),
        ("Sheet1!B3:E3", "F3"),
    ]


def test_follows_header_row(tmp_path):
    lines = _lines(_write(tmp_path, {"trend": {"from": "q1", "to": "q4"}}, header_row=2))
    assert lines == [("Sheet1!B4:E4", "F4"), ("Sheet1!B5:E5", "F5")]


def test_span_can_be_narrower(tmp_path):
    assert _lines(_write(tmp_path, {"trend": {"from": "q2", "to": "q3"}}))[0] == (
        "Sheet1!C2:D2",
        "F2",
    )


def test_no_data_rows_writes_nothing(tmp_path):
    assert _lines(_write(tmp_path, {"trend": {"from": "q1", "to": "q4"}}, rows=[])) == []


def test_multi_sheet_range_names_its_own_sheet(tmp_path):
    path = tmp_path / "m.xlsx"
    write_worksheets(
        [("First", ROWS), ("Second", ROWS)],
        str(path),
        sparklines={"Second": {"trend": {"from": "q1", "to": "q4"}}},
    )
    xml = zipfile.ZipFile(path).read("xl/worksheets/sheet2.xml").decode()
    assert "Second!B2:E2" in xml


# --- options ----------------------------------------------------------------


def test_type_column(tmp_path):
    assert 'type="column"' in _group(
        _write(tmp_path, {"trend": {"from": "q1", "to": "q4", "type": "column"}})
    )


def test_type_win_lose(tmp_path):
    assert 'type="stacked"' in _group(
        _write(tmp_path, {"trend": {"from": "q1", "to": "q4", "type": "win_lose"}})
    )


def test_line_is_the_default(tmp_path):
    """Excel omits the attribute for line, its own default."""
    assert "type=" not in _group(_write(tmp_path, {"trend": {"from": "q1", "to": "q4"}}))


def test_point_toggles(tmp_path):
    group = _group(
        _write(
            tmp_path,
            {"trend": {"from": "q1", "to": "q4", "high_point": True,
                       "low_point": True, "markers": True}},
        )
    )
    assert 'high="1"' in group
    assert 'low="1"' in group
    assert 'markers="1"' in group


def test_color(tmp_path):
    path = _write(tmp_path, {"trend": {"from": "q1", "to": "q4", "color": "#FF0000"}})
    assert 'colorSeries rgb="FFFF0000"' in _xml(path)


# --- validation -------------------------------------------------------------


def test_unknown_target_column_warns(tmp_path):
    with pytest.warns(UserWarning, match="sparklines: unknown column in 'nope'"):
        path = _write(tmp_path, {"nope": {"from": "q1", "to": "q4"}})
    assert _lines(path) == []


def test_unknown_span_column_warns(tmp_path):
    with pytest.warns(UserWarning, match=r"\('q1'\.\.'q9'\)"):
        path = _write(tmp_path, {"trend": {"from": "q1", "to": "q9"}})
    assert _lines(path) == []


def test_reversed_span_raises(tmp_path):
    with pytest.raises(ValueError, match="'q4' comes after 'q1' in the data"):
        _write(tmp_path, {"trend": {"from": "q4", "to": "q1"}})


def test_needs_from_and_to(tmp_path):
    with pytest.raises(ValueError, match="'trend' needs 'to'"):
        _write(tmp_path, {"trend": {"from": "q1"}})


def test_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown key 'kind' for 'trend'"):
        _write(tmp_path, {"trend": {"from": "q1", "to": "q4", "kind": "line"}})


def test_unknown_type_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown type 'area'"):
        _write(tmp_path, {"trend": {"from": "q1", "to": "q4", "type": "area"}})


def test_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="must be a dict keyed by column name"):
        _write(tmp_path, ["trend"])


def test_rule_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="rule for 'trend' must be a dict"):
        _write(tmp_path, {"trend": ["q1", "q4"]})


# --- builder ----------------------------------------------------------------


def test_builder(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet(
        "S", ROWS, sparklines={"trend": {"from": "q1", "to": "q4"}}
    ).save()
    assert len(_lines(path)) == 2
