"""Page and print setup, given as one ``page_setup`` mapping.

Excel has about twenty of these settings. A keyword each would have taken
``write_worksheet`` from 27 parameters to nearly 50, so they arrive together
and are validated in one place — which also means a typo has to be caught
here, since a mapping key that does nothing gives no other signal.

Read back with openpyxl where it exposes the setting, and from the sheet XML
where it does not.
"""

import re
import zipfile

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [{"a": 1, "b": 2, "c": 3}]


def _sheet(tmp_path, **setup):
    path = tmp_path / "page.xlsx"
    write_worksheet(ROWS, str(path), page_setup=setup)
    return path, openpyxl.load_workbook(path).active


def _page_setup_xml(path):
    xml = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode()
    return re.search(r"<pageSetup[^>]*>", xml).group(0)


# --- orientation, paper, scaling --------------------------------------------


def test_landscape(tmp_path):
    _, ws = _sheet(tmp_path, landscape=True)
    assert ws.page_setup.orientation == "landscape"


def test_landscape_false_is_portrait(tmp_path):
    """The key is a bool, not a flag, so False must mean portrait."""
    _, ws = _sheet(tmp_path, landscape=False)
    assert ws.page_setup.orientation == "portrait"


def test_paper_size(tmp_path):
    _, ws = _sheet(tmp_path, paper_size=9)  # A4
    assert ws.page_setup.paperSize == 9


def test_scale(tmp_path):
    path, _ = _sheet(tmp_path, scale=80)
    assert 'scale="80"' in _page_setup_xml(path)


def test_fit_to_pages(tmp_path):
    """Height 0 lets the sheet run to as many pages as it needs."""
    _, ws = _sheet(tmp_path, fit_to_pages=(1, 0))
    assert ws.page_setup.fitToHeight == 0


def test_scale_and_fit_to_pages_together_raise(tmp_path):
    """Excel honours only one; failing beats opening a file that ignores half."""
    with pytest.raises(ValueError, match="mutually exclusive"):
        _sheet(tmp_path, scale=80, fit_to_pages=(1, 0))


def test_first_page_number(tmp_path):
    path, _ = _sheet(tmp_path, first_page_number=3)
    assert 'useFirstPageNumber="3"' in _page_setup_xml(path)


# --- repeated titles and print area -----------------------------------------


def test_repeat_rows_from_an_index(tmp_path):
    """The reason this feature exists: the header on every printed page."""
    _, ws = _sheet(tmp_path, repeat_rows=0)
    assert ws.print_title_rows == "$1:$1"


def test_repeat_rows_from_a_pair(tmp_path):
    _, ws = _sheet(tmp_path, repeat_rows=(0, 1))
    assert ws.print_title_rows == "$1:$2"


def test_repeat_columns(tmp_path):
    _, ws = _sheet(tmp_path, repeat_columns=0)
    assert ws.print_title_cols == "$A:$A"


def test_print_area(tmp_path):
    _, ws = _sheet(tmp_path, print_area=(0, 0, 9, 2))
    assert ws.print_area == "'Sheet1'!$A$1:$C$10"


def test_repeat_rows_rejects_a_bad_shape(tmp_path):
    with pytest.raises(ValueError, match="must be an index or a \\(first, last\\) pair"):
        _sheet(tmp_path, repeat_rows="one")


def test_print_area_rejects_a_bad_shape(tmp_path):
    with pytest.raises(ValueError, match="first_row, first_col, last_row, last_col"):
        _sheet(tmp_path, print_area=(0, 0))


# --- margins ----------------------------------------------------------------


def test_margins_keep_excel_defaults_for_omitted_sides(tmp_path):
    _, ws = _sheet(tmp_path, margins={"left": 0.5})
    assert ws.page_margins.left == 0.5
    assert ws.page_margins.right == 0.7
    assert ws.page_margins.top == 0.75


def test_all_margins(tmp_path):
    _, ws = _sheet(
        tmp_path,
        margins={
            "left": 0.1,
            "right": 0.2,
            "top": 0.3,
            "bottom": 0.4,
            "header": 0.5,
            "footer": 0.6,
        },
    )
    m = ws.page_margins
    assert (m.left, m.right, m.top, m.bottom, m.header, m.footer) == (
        0.1,
        0.2,
        0.3,
        0.4,
        0.5,
        0.6,
    )


def test_unknown_margin_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown margin 'middle'"):
        _sheet(tmp_path, margins={"middle": 1.0})


def test_margins_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="'margins' must be a dict"):
        _sheet(tmp_path, margins=(0.5, 0.5, 0.5, 0.5, 0.5, 0.5))


# --- headers, footers, print options ----------------------------------------


def test_header_and_footer(tmp_path):
    _, ws = _sheet(tmp_path, header="&CReport", footer="&RPage &P of &N")
    assert ws.oddHeader.center.text == "Report"
    # openpyxl hands back the &-codes as written, not expanded.
    assert ws.oddFooter.right.text == "Page &P of &N"


def test_print_gridlines_and_headings(tmp_path):
    _, ws = _sheet(tmp_path, print_gridlines=True, print_headings=True)
    assert ws.print_options.gridLines
    assert ws.print_options.headings


def test_centering(tmp_path):
    _, ws = _sheet(tmp_path, center_horizontally=True, center_vertically=True)
    assert ws.print_options.horizontalCentered
    assert ws.print_options.verticalCentered


# --- validation -------------------------------------------------------------


def test_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match="unknown key 'landscap'"):
        _sheet(tmp_path, landscap=True)


def test_unknown_key_names_the_accepted_ones(tmp_path):
    with pytest.raises(ValueError, match="repeat_rows"):
        _sheet(tmp_path, nope=1)


def test_must_be_a_dict(tmp_path):
    path = tmp_path / "x.xlsx"
    with pytest.raises(ValueError, match="page_setup must be a dict"):
        write_worksheet(ROWS, str(path), page_setup=["landscape"])


# --- every entry point ------------------------------------------------------


def test_multi_sheet_is_keyed_by_sheet_name(tmp_path):
    path = tmp_path / "multi.xlsx"
    write_worksheets(
        [("A", ROWS), ("B", ROWS)],
        str(path),
        page_setup={"A": {"landscape": True}},
    )
    book = openpyxl.load_workbook(path)
    assert book["A"].page_setup.orientation == "landscape"
    assert book["B"].page_setup.orientation != "landscape"


def test_builder(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet("S", ROWS, page_setup={"landscape": True}).save()
    assert openpyxl.load_workbook(path).active.page_setup.orientation == "landscape"


def test_csv_warns_that_it_is_dropped(tmp_path):
    with pytest.warns(UserWarning, match="page_setup"):
        FastExcel(str(tmp_path / "o.csv")).sheet(
            "S", ROWS, page_setup={"landscape": True}
        ).save()
