"""Screen presentation (``sheet_view``) and error indicators (``ignore_errors``).

Both are per sheet. ``sheet_view`` is applied before the data with the rest of
the layout; ``ignore_errors`` after it, since its range depends on how many
rows there turned out to be.
"""

import re
import zipfile

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [{"sku": "0012", "qty": 1}, {"sku": "0034", "qty": 2}]


def _write(tmp_path, **kwargs):
    path = tmp_path / "v.xlsx"
    write_worksheet(ROWS, str(path), **kwargs)
    return path


def _xml(path):
    return zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode()


# --- sheet_view -------------------------------------------------------------


def test_tab_color(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, sheet_view={"tab_color": "#FF0000"})).active
    assert ws.sheet_properties.tabColor.rgb == "FFFF0000"


def test_tab_color_by_name(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, sheet_view={"tab_color": "red"})).active
    assert ws.sheet_properties.tabColor.rgb == "FFFF0000"


def test_gridlines_off(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, sheet_view={"gridlines": False})).active
    assert ws.sheet_view.showGridLines is False


def test_gridlines_stay_on_by_default(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path)).active
    assert ws.sheet_view.showGridLines is not False


def test_zoom(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, sheet_view={"zoom": 120})).active
    assert ws.sheet_view.zoomScale == 120


def test_right_to_left(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, sheet_view={"right_to_left": True})).active
    assert ws.sheet_view.rightToLeft


def test_hidden(tmp_path):
    path = tmp_path / "h.xlsx"
    write_worksheets(
        [("A", ROWS), ("B", ROWS)],
        str(path),
        sheet_view={"B": {"hidden": True}},
    )
    book = openpyxl.load_workbook(path)
    assert book["B"].sheet_state == "hidden"
    assert book["A"].sheet_state == "visible"


def test_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match="sheet_view: unknown key 'tabcolor'"):
        _write(tmp_path, sheet_view={"tabcolor": "red"})


def test_hidden_and_selected_together_raise(tmp_path):
    """Excel rejects a workbook whose active sheet is hidden."""
    with pytest.raises(ValueError, match="cannot be both 'hidden' and 'selected'"):
        _write(tmp_path, sheet_view={"hidden": True, "selected": True})


def test_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="sheet_view must be a dict"):
        _write(tmp_path, sheet_view=["gridlines"])


def test_bad_colour_is_rejected(tmp_path):
    with pytest.raises(ValueError, match="invalid color"):
        _write(tmp_path, sheet_view={"tab_color": "not-a-colour"})


# --- ignore_errors ----------------------------------------------------------


def _ignored(path):
    """``(sqref, attribute)`` for every ignoredError entry."""
    return re.findall(r'<ignoredError sqref="([^"]+)" (\w+)="1"', _xml(path))


def test_list_form_means_number_stored_as_text(tmp_path):
    """The case this exists for: an ID column of digits kept as text."""
    assert _ignored(_write(tmp_path, ignore_errors=["sku"])) == [
        ("A2:A3", "numberStoredAsText")
    ]


def test_range_covers_the_data_rows_only(tmp_path):
    """Never the header — its text is not a mistyped number."""
    (sqref, _), = _ignored(_write(tmp_path, ignore_errors=["sku"], header_row=2))
    assert sqref == "A4:A5"


def test_dict_form_names_the_error(tmp_path):
    """``formula_error`` is Excel's ``evalError`` attribute."""
    assert _ignored(_write(tmp_path, ignore_errors={"qty": "formula_error"})) == [
        ("B2:B3", "evalError")
    ]


def test_a_column_cannot_carry_two_errors(tmp_path):
    """Excel allows one ignore rule per cell, so the API must not offer a list."""
    with pytest.raises(ValueError, match="must be a single error name"):
        _write(tmp_path, ignore_errors={"sku": ["number_stored_as_text"]})


def test_several_columns_share_one_entry(tmp_path):
    """Same error on two columns collapses into a single sqref."""
    assert _ignored(_write(tmp_path, ignore_errors=["sku", "qty"])) == [
        ("A2:A3 B2:B3", "numberStoredAsText")
    ]


def test_no_data_rows_writes_nothing(tmp_path):
    path = tmp_path / "e.xlsx"
    write_worksheet([], str(path), ignore_errors=["sku"])
    assert "ignoredError" not in _xml(path)


def test_unknown_column_warns_and_is_skipped(tmp_path):
    with pytest.warns(UserWarning, match="ignore_errors: unknown column 'nope'"):
        path = _write(tmp_path, ignore_errors=["nope"])
    assert "ignoredError" not in _xml(path)


def test_unknown_error_name_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown error 'green_triangle'"):
        _write(tmp_path, ignore_errors={"sku": "green_triangle"})


def test_bad_shape_is_rejected(tmp_path):
    with pytest.raises(ValueError, match="must be a list of column names or a dict"):
        _write(tmp_path, ignore_errors=42)


# --- builder ----------------------------------------------------------------


def test_builder_passes_both(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet(
        "S", ROWS, sheet_view={"zoom": 150}, ignore_errors=["sku"]
    ).save()
    assert openpyxl.load_workbook(path).active.sheet_view.zoomScale == 150
    assert _ignored(path) == [("A2:A3", "numberStoredAsText")]
