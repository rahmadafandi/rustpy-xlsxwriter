"""Per-column conditional formatting.

openpyxl exposes conditional formatting well enough to check both the rule and
the range it covers, and the range is the part worth checking: it is computed
from the rows actually written, so an off-by-one would shade the header or
miss the last row.

openpyxl warns that it is dropping the x14 extension when it reads a data bar
— rust_xlsxwriter writes the modern form, which openpyxl does not model. The
rule itself still reads back, so the warning is openpyxl's limit, not a flaw
in the file.
"""

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, Format, write_worksheet, write_worksheets

ROWS = [{"score": 10, "name": "alpha"}, {"score": 20, "name": "beta"}]


def _sheet(tmp_path, formats, rows=None, **kwargs):
    path = tmp_path / "cf.xlsx"
    write_worksheet(
        rows if rows is not None else ROWS,
        str(path),
        conditional_formats=formats,
        **kwargs,
    )
    return openpyxl.load_workbook(path).active


def _rules(ws):
    """Every ``(range, rule)`` pair on the sheet."""
    return [(str(rng.sqref), rule) for rng in ws.conditional_formatting for rule in rng.rules]


# --- range ------------------------------------------------------------------


def test_range_covers_the_data_rows_only(tmp_path):
    """Never the header, and exactly as many rows as were written."""
    ws = _sheet(tmp_path, {"score": {"type": "data_bar"}})
    assert [rng for rng, _ in _rules(ws)] == ["A2:A3"]


def test_range_follows_header_row(tmp_path):
    ws = _sheet(tmp_path, {"score": {"type": "data_bar"}}, header_row=2)
    assert [rng for rng, _ in _rules(ws)] == ["A4:A5"]


def test_range_picks_the_named_column(tmp_path):
    ws = _sheet(tmp_path, {"name": {"type": "duplicate"}})
    assert [rng for rng, _ in _rules(ws)] == ["B2:B3"]


def test_no_data_rows_writes_no_rule(tmp_path):
    """An empty range is not a range Excel accepts."""
    ws = _sheet(tmp_path, {"score": {"type": "data_bar"}}, rows=[])
    assert _rules(ws) == []


# --- rule types -------------------------------------------------------------


def test_data_bar(tmp_path):
    ws = _sheet(tmp_path, {"score": {"type": "data_bar", "color": "#638EC6"}})
    (_, rule), = _rules(ws)
    assert rule.type == "dataBar"
    assert rule.dataBar.color.rgb == "FF638EC6"


def test_three_color_scale(tmp_path):
    ws = _sheet(tmp_path, {"score": {"type": "3_color_scale"}})
    (_, rule), = _rules(ws)
    assert rule.type == "colorScale"
    assert len(rule.colorScale.cfvo) == 3


def test_two_color_scale(tmp_path):
    ws = _sheet(tmp_path, {"score": {"type": "2_color_scale"}})
    (_, rule), = _rules(ws)
    assert len(rule.colorScale.cfvo) == 2


def test_cell_greater_than(tmp_path):
    ws = _sheet(
        tmp_path,
        {"score": {"type": "cell", "criteria": ">", "value": 15,
                   "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.operator == "greaterThan"
    assert rule.formula == ["15"]


def test_cell_between(tmp_path):
    ws = _sheet(
        tmp_path,
        {"score": {"type": "cell", "criteria": "between", "min": 5, "max": 15,
                   "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.operator == "between"
    assert rule.formula == ["5", "15"]


def test_text_contains(tmp_path):
    ws = _sheet(
        tmp_path,
        {"name": {"type": "text", "criteria": "contains", "value": "alp",
                  "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.type == "containsText"
    assert rule.text == "alp"


def test_top(tmp_path):
    ws = _sheet(tmp_path, {"score": {"type": "top", "value": 1,
                                     "format": Format().set_bold()}})
    (_, rule), = _rules(ws)
    assert rule.type == "top10"
    assert rule.rank == 1


def test_average(tmp_path):
    ws = _sheet(tmp_path, {"score": {"type": "average", "criteria": "above",
                                     "format": Format().set_bold()}})
    (_, rule), = _rules(ws)
    assert rule.type == "aboveAverage"


def test_duplicate_and_unique_differ(tmp_path):
    dup = _sheet(tmp_path, {"name": {"type": "duplicate"}})
    uniq = _sheet(tmp_path, {"name": {"type": "unique"}})
    assert _rules(dup)[0][1].type == "duplicateValues"
    assert _rules(uniq)[0][1].type == "uniqueValues"


# --- several rules ----------------------------------------------------------


def test_a_column_can_carry_several_rules(tmp_path):
    ws = _sheet(
        tmp_path,
        {"score": [
            {"type": "data_bar"},
            {"type": "cell", "criteria": ">", "value": 15,
             "format": Format().set_bold()},
        ]},
    )
    assert sorted(rule.type for _, rule in _rules(ws)) == ["cellIs", "dataBar"]


def test_several_columns(tmp_path):
    ws = _sheet(
        tmp_path,
        {"score": {"type": "data_bar"}, "name": {"type": "duplicate"}},
    )
    assert sorted(rng for rng, _ in _rules(ws)) == ["A2:A3", "B2:B3"]


# --- validation -------------------------------------------------------------


def test_unknown_column_warns_and_is_skipped(tmp_path):
    """A rule that cannot be placed costs shading, not the export."""
    with pytest.warns(UserWarning, match="unknown column 'nope'"):
        ws = _sheet(tmp_path, {"nope": {"type": "data_bar"}})
    assert _rules(ws) == []


def test_unknown_type_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown type 'sparkline'"):
        _sheet(tmp_path, {"score": {"type": "sparkline"}})


def test_unknown_criteria_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown criteria '=>'"):
        _sheet(tmp_path, {"score": {"type": "cell", "criteria": "=>", "value": 1}})


def test_missing_value_raises(tmp_path):
    with pytest.raises(ValueError, match="a 'cell' rule needs 'value'"):
        _sheet(tmp_path, {"score": {"type": "cell", "criteria": ">"}})


def test_between_needs_min_and_max(tmp_path):
    with pytest.raises(ValueError, match="a 'cell' rule needs 'max'"):
        _sheet(tmp_path, {"score": {"type": "cell", "criteria": "between", "min": 1}})


def test_missing_type_raises(tmp_path):
    with pytest.raises(ValueError, match="needs 'type'"):
        _sheet(tmp_path, {"score": {"criteria": ">"}})


def test_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="must be a dict keyed by column name"):
        _sheet(tmp_path, ["score"])


# --- every entry point ------------------------------------------------------


def test_multi_sheet_is_keyed_by_sheet_name(tmp_path):
    path = tmp_path / "multi.xlsx"
    write_worksheets(
        [("A", ROWS), ("B", ROWS)],
        str(path),
        conditional_formats={"A": {"score": {"type": "data_bar"}}},
    )
    book = openpyxl.load_workbook(path)
    assert len(_rules(book["A"])) == 1
    assert _rules(book["B"]) == []


def test_builder(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet(
        "S", ROWS, conditional_formats={"score": {"type": "data_bar"}}
    ).save()
    assert len(_rules(openpyxl.load_workbook(path).active)) == 1


# --- every criteria, not just the one per type that was exercised -----------
#
# These map a string onto an enum variant and nothing else. A variant wired to
# the wrong string compiles and writes a file that is quietly wrong, so each
# needs to be seen once. Before this, four of the eight cell criteria and three
# of the four text criteria had never been executed.


@pytest.mark.parametrize(
    "criteria,operator",
    [
        ("==", "equal"),
        ("equal_to", "equal"),
        ("!=", "notEqual"),
        ("not_equal_to", "notEqual"),
        (">", "greaterThan"),
        ("greater_than", "greaterThan"),
        (">=", "greaterThanOrEqual"),
        ("<", "lessThan"),
        ("less_than", "lessThan"),
        ("<=", "lessThanOrEqual"),
    ],
)
def test_every_cell_criteria(tmp_path, criteria, operator):
    ws = _sheet(
        tmp_path,
        {"score": {"type": "cell", "criteria": criteria, "value": 15,
                   "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.operator == operator


def test_cell_not_between(tmp_path):
    ws = _sheet(
        tmp_path,
        {"score": {"type": "cell", "criteria": "not_between", "min": 5, "max": 15,
                   "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.operator == "notBetween"
    assert rule.formula == ["5", "15"]


@pytest.mark.parametrize(
    "criteria,kind",
    [
        ("contains", "containsText"),
        ("does_not_contain", "notContainsText"),
        ("begins_with", "beginsWith"),
        ("ends_with", "endsWith"),
    ],
)
def test_every_text_criteria(tmp_path, criteria, kind):
    ws = _sheet(
        tmp_path,
        {"name": {"type": "text", "criteria": criteria, "value": "alp",
                  "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.type == kind


# Excel omits a flag that is at its default, so `bottom` absent means top and
# `aboveAverage` absent means above. The expectations below are the attribute
# values as written, which is what makes each case tell the others apart.
@pytest.mark.parametrize(
    "criteria,bottom,percent",
    [
        ("top", None, None),
        ("bottom", True, None),
        ("top_percent", None, True),
        ("bottom_percent", True, True),
    ],
)
def test_every_top_criteria(tmp_path, criteria, bottom, percent):
    ws = _sheet(
        tmp_path,
        {"score": {"type": "top", "criteria": criteria, "value": 1,
                   "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.type == "top10"
    assert rule.rank == 1
    assert (rule.bottom, rule.percent) == (bottom, percent)


@pytest.mark.parametrize(
    "criteria,above,equal",
    [
        ("above", None, None),
        ("below", False, None),
        ("equal_or_above", None, True),
        ("equal_or_below", False, True),
    ],
)
def test_every_average_criteria(tmp_path, criteria, above, equal):
    """Excel models all four as aboveAverage with two flags, not four types."""
    ws = _sheet(
        tmp_path,
        {"score": {"type": "average", "criteria": criteria,
                   "format": Format().set_bold()}},
    )
    (_, rule), = _rules(ws)
    assert rule.type == "aboveAverage"
    assert (rule.aboveAverage, rule.equalAverage) == (above, equal)
