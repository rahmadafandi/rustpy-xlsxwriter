"""Charts anchored to a cell.

The ranges are what can silently come out wrong: a series must cover the data
rows of the column it names, never the header, and the categories must line up
with them. Those are read straight from the chart XML.
"""

import re
import zipfile

import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [
    {"region": "north", "q1": 1, "q2": 4},
    {"region": "south", "q1": 2, "q2": 5},
    {"region": "east", "q1": 3, "q2": 6},
]


def _write(tmp_path, charts, rows=None, **kwargs):
    path = tmp_path / "c.xlsx"
    write_worksheet(
        rows if rows is not None else ROWS, str(path), charts=charts, **kwargs
    )
    return path


def _chart(path, n=1):
    return zipfile.ZipFile(path).read(f"xl/charts/chart{n}.xml").decode()


def _refs(path, n=1):
    return re.findall(r"<c:f>(.*?)</c:f>", _chart(path, n))


def _has_chart(path):
    return any("charts/chart" in name for name in zipfile.ZipFile(path).namelist())


COLUMN = {"type": "column", "series": ["q1"]}


# --- ranges -----------------------------------------------------------------


def test_series_covers_the_data_rows(tmp_path):
    """Never the header row — that is the series name, not a value."""
    assert _refs(_write(tmp_path, [COLUMN])) == ["Sheet1!$B$1", "Sheet1!$B$2:$B$4"]


def test_series_name_links_to_the_header_cell(tmp_path):
    """So the legend follows the header if it is ever edited."""
    assert _refs(_write(tmp_path, [COLUMN]))[0] == "Sheet1!$B$1"


def test_explicit_series_name_replaces_the_link(tmp_path):
    path = _write(
        tmp_path, [{"type": "column", "series": [{"values": "q1", "name": "Quarter 1"}]}]
    )
    assert _refs(path) == ["Sheet1!$B$2:$B$4"]
    assert "Quarter 1" in _chart(path)


def test_categories(tmp_path):
    refs = _refs(
        _write(tmp_path, [{"type": "column", "series": ["q1"], "categories": "region"}])
    )
    assert refs == ["Sheet1!$B$1", "Sheet1!$A$2:$A$4", "Sheet1!$B$2:$B$4"]


def test_several_series(tmp_path):
    refs = _refs(_write(tmp_path, [{"type": "line", "series": ["q1", "q2"]}]))
    assert refs == [
        "Sheet1!$B$1",
        "Sheet1!$B$2:$B$4",
        "Sheet1!$C$1",
        "Sheet1!$C$2:$C$4",
    ]


def test_ranges_follow_header_row(tmp_path):
    refs = _refs(_write(tmp_path, [COLUMN], header_row=2))
    assert refs == ["Sheet1!$B$3", "Sheet1!$B$4:$B$6"]


def test_no_data_rows_writes_no_chart(tmp_path):
    assert not _has_chart(_write(tmp_path, [COLUMN], rows=[]))


def test_multi_sheet_ranges_name_their_own_sheet(tmp_path):
    path = tmp_path / "m.xlsx"
    write_worksheets(
        [("First", ROWS), ("Second", ROWS)],
        str(path),
        charts={"Second": [COLUMN]},
    )
    assert "Second!$B$2:$B$4" in zipfile.ZipFile(path).read("xl/charts/chart1.xml").decode()


# --- placement and presentation ---------------------------------------------


def test_default_placement_clears_the_data(tmp_path):
    """Three data columns, so the chart starts at E — one column clear."""
    drawing = zipfile.ZipFile(_write(tmp_path, [COLUMN])).read(
        "xl/drawings/drawing1.xml"
    ).decode()
    assert "<xdr:col>4</xdr:col>" in drawing


def test_explicit_placement(tmp_path):
    drawing = zipfile.ZipFile(_write(tmp_path, [{**COLUMN, "row": 10, "col": 1}])).read(
        "xl/drawings/drawing1.xml"
    ).decode()
    assert "<xdr:col>1</xdr:col>" in drawing
    assert "<xdr:row>10</xdr:row>" in drawing


def test_title_and_axis_names(tmp_path):
    text = re.findall(
        r"<a:t>(.*?)</a:t>",
        _chart(_write(tmp_path, [{**COLUMN, "title": "Sales", "x_axis": "Region",
                                  "y_axis": "Revenue"}])),
    )
    assert {"Sales", "Region", "Revenue"} <= set(text)


def test_legend_can_be_hidden(tmp_path):
    assert "<c:legend>" not in _chart(_write(tmp_path, [{**COLUMN, "legend": False}]))


def test_legend_is_there_by_default(tmp_path):
    assert "<c:legend>" in _chart(_write(tmp_path, [COLUMN]))


@pytest.mark.parametrize(
    "kind,marker",
    [
        ("column", "<c:barChart>"),
        ("bar", "<c:barChart>"),
        ("line", "<c:lineChart>"),
        ("pie", "<c:pieChart>"),
        ("doughnut", "<c:doughnutChart>"),
        ("area", "<c:areaChart>"),
        ("radar", "<c:radarChart>"),
    ],
)
def test_chart_types(tmp_path, kind, marker):
    assert marker in _chart(_write(tmp_path, [{"type": kind, "series": ["q1"]}]))


def test_scatter_needs_categories(tmp_path):
    """Its categories are the x values, so a scatter without them is not a chart.

    The crate rejects this too, but only after every row has been written;
    catching it up front costs the caller nothing.
    """
    with pytest.raises(ValueError, match="a scatter chart needs 'categories'"):
        _write(tmp_path, [{"type": "scatter", "series": ["q1"]}])


def test_scatter_with_categories(tmp_path):
    path = _write(
        tmp_path, [{"type": "scatter", "series": ["q1"], "categories": "q2"}]
    )
    assert "<c:scatterChart>" in _chart(path)


def test_stacked_variant(tmp_path):
    chart = _chart(_write(tmp_path, [{"type": "column_stacked", "series": ["q1", "q2"]}]))
    assert 'val="stacked"' in chart


def test_two_charts(tmp_path):
    path = _write(
        tmp_path,
        [COLUMN, {"type": "line", "series": ["q2"], "row": 20, "col": 1}],
    )
    assert _refs(path, 1) == ["Sheet1!$B$1", "Sheet1!$B$2:$B$4"]
    assert _refs(path, 2) == ["Sheet1!$C$1", "Sheet1!$C$2:$C$4"]


# --- validation -------------------------------------------------------------


def test_unknown_series_column_warns_and_skips_the_chart(tmp_path):
    """A chart missing a series would draw a misleading picture."""
    with pytest.warns(UserWarning, match=r"charts\[0\]: unknown column 'nope'"):
        path = _write(tmp_path, [{"type": "column", "series": ["nope"]}])
    assert not _has_chart(path)


def test_unknown_categories_column_warns(tmp_path):
    with pytest.warns(UserWarning, match="unknown column 'nope'"):
        path = _write(tmp_path, [{**COLUMN, "categories": "nope"}])
    assert not _has_chart(path)


def test_unknown_type_raises(tmp_path):
    with pytest.raises(ValueError, match="charts: unknown type 'bubble'"):
        _write(tmp_path, [{"type": "bubble", "series": ["q1"]}])


def test_needs_type(tmp_path):
    with pytest.raises(ValueError, match=r"charts\[0\]: needs 'type'"):
        _write(tmp_path, [{"series": ["q1"]}])


def test_needs_series(tmp_path):
    with pytest.raises(ValueError, match=r"charts\[0\]: needs 'series'"):
        _write(tmp_path, [{"type": "column"}])


def test_empty_series_raises(tmp_path):
    with pytest.raises(ValueError, match="'series' is empty"):
        _write(tmp_path, [{"type": "column", "series": []}])


def test_series_dict_needs_values(tmp_path):
    with pytest.raises(ValueError, match="a series dict needs 'values'"):
        _write(tmp_path, [{"type": "column", "series": [{"name": "Q1"}]}])


def test_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match=r"charts\[0\]: unknown key 'colour'"):
        _write(tmp_path, [{**COLUMN, "colour": "red"}])


def test_index_is_named_in_the_error(tmp_path):
    with pytest.raises(ValueError, match=r"charts\[1\]: needs 'type'"):
        _write(tmp_path, [COLUMN, {"series": ["q2"]}])


def test_must_be_a_list(tmp_path):
    with pytest.raises(ValueError, match="each chart must be a dict"):
        _write(tmp_path, ["column"])


# --- builder ----------------------------------------------------------------


def test_builder(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet("S", ROWS, charts=[COLUMN]).save()
    assert _has_chart(path)
