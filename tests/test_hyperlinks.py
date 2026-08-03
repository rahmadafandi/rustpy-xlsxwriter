"""Hyperlinks via ``url_columns``.

``write_url`` rejects anything it cannot classify — plain text, an empty
string, a URL past 2083 characters — so the writer falls back to plain text
rather than aborting an export on one bad cell. These tests pin both halves:
real links become links, everything else survives as readable text.
"""

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, Format, write_worksheet, write_worksheets


def _cell(path, row, col, sheet=None):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet] if sheet else wb.active
    return ws.cell(row=row, column=col)


def test_off_by_default(tmp_path):
    path = tmp_path / "off.xlsx"
    write_worksheet([{"link": "https://example.com"}], str(path))
    c = _cell(path, 2, 1)
    assert c.value == "https://example.com"
    assert c.hyperlink is None


def test_url_column_becomes_a_link(tmp_path):
    path = tmp_path / "link.xlsx"
    write_worksheet(
        [{"name": "Example", "link": "https://example.com"}],
        str(path),
        url_columns=["link"],
    )
    c = _cell(path, 2, 2)
    assert c.hyperlink is not None
    assert c.hyperlink.target == "https://example.com"
    assert c.value == "https://example.com"
    # The non-URL column is untouched.
    assert _cell(path, 2, 1).hyperlink is None


def test_mailto_and_internal_links(tmp_path):
    path = tmp_path / "kinds.xlsx"
    write_worksheet(
        [
            {"link": "mailto:a@b.com"},
            {"link": "internal:Sheet1!A1"},
        ],
        str(path),
        sheet_name="Sheet1",
        url_columns=["link"],
    )
    assert _cell(path, 2, 1, "Sheet1").hyperlink.target == "mailto:a@b.com"
    # Internal links are stored as a location, not a target.
    internal = _cell(path, 3, 1, "Sheet1").hyperlink
    assert internal.location == "Sheet1!A1"


@pytest.mark.parametrize(
    "value",
    [
        pytest.param("just some text", id="plain-text"),
        pytest.param("", id="empty"),
        pytest.param("https://e.com/" + "x" * 2085, id="over-2083-chars"),
    ],
)
def test_non_links_fall_back_to_text(tmp_path, value):
    """A bad value must not abort the export."""
    path = tmp_path / "fallback.xlsx"
    write_worksheet(
        [{"link": value}, {"link": "https://example.com"}],
        str(path),
        url_columns=["link"],
    )
    bad = _cell(path, 2, 1)
    assert bad.hyperlink is None
    assert (bad.value or "") == value
    # The following good row still linked.
    assert _cell(path, 3, 1).hyperlink.target == "https://example.com"


def test_none_stays_blank(tmp_path):
    path = tmp_path / "none.xlsx"
    write_worksheet(
        [{"link": None}, {"link": "https://example.com"}],
        str(path),
        url_columns=["link"],
    )
    assert _cell(path, 2, 1).hyperlink is None
    assert _cell(path, 3, 1).hyperlink is not None


def test_non_string_column_is_ignored(tmp_path):
    """Numbers in a url column stay numbers."""
    path = tmp_path / "numeric.xlsx"
    write_worksheet([{"link": 42}], str(path), url_columns=["link"])
    c = _cell(path, 2, 1)
    assert c.value == 42
    assert c.hyperlink is None


def test_unknown_column_warns(tmp_path):
    with pytest.warns(UserWarning, match="unknown column 'nope'"):
        write_worksheet(
            [{"link": "https://example.com"}],
            str(tmp_path / "warn.xlsx"),
            url_columns=["nope"],
        )


def test_links_keep_banding_and_column_format(tmp_path):
    path = tmp_path / "styled.xlsx"
    write_worksheet(
        [{"link": f"https://example.com/{i}"} for i in range(4)],
        str(path),
        url_columns=["link"],
        banded_rows="#F2F2F2",
    )
    banded = _cell(path, 3, 1)
    assert banded.hyperlink is not None
    assert banded.fill.fgColor.rgb == "FFF2F2F2"


@pytest.mark.parametrize("frame", ["pandas", "polars"])
def test_dataframe_paths(tmp_path, frame):
    mod = pytest.importorskip(frame)
    df = mod.DataFrame(
        {"name": ["a", "b"], "link": ["https://example.com", "not a link"]}
    )
    path = tmp_path / f"{frame}.xlsx"
    write_worksheet(df, str(path), url_columns=["link"])

    assert _cell(path, 2, 2).hyperlink.target == "https://example.com"
    assert _cell(path, 3, 2).hyperlink is None
    assert _cell(path, 3, 2).value == "not a link"


def test_dataframe_fallback_path(tmp_path):
    from tests.test_row_layout import _FakeFrame

    df = _FakeFrame({"link": ["https://example.com", "nope"]}, kinds=["O"])
    path = tmp_path / "fallback-df.xlsx"
    write_worksheet(df, str(path), url_columns=["link"])
    assert _cell(path, 2, 1).hyperlink.target == "https://example.com"
    assert _cell(path, 3, 1).hyperlink is None


def test_multi_sheet_is_per_sheet(tmp_path):
    path = tmp_path / "multi.xlsx"
    rows = [{"link": "https://example.com"}]
    write_worksheets(
        [("Linked", rows), ("Plain", rows)],
        str(path),
        url_columns={"Linked": ["link"]},
    )
    assert _cell(path, 2, 1, "Linked").hyperlink is not None
    assert _cell(path, 2, 1, "Plain").hyperlink is None


def test_fastexcel_builder(tmp_path):
    path = tmp_path / "builder.xlsx"
    (
        FastExcel(str(path))
        .sheet(
            "S",
            [{"link": "https://example.com"}],
            url_columns=["link"],
            header_format=Format().set_bold(),
        )
        .save()
    )
    assert _cell(path, 2, 1, "S").hyperlink is not None


def test_csv_warns_that_url_columns_are_dropped(tmp_path):
    with pytest.warns(UserWarning, match="url_columns"):
        FastExcel(str(tmp_path / "o.csv")).sheet(
            "S", [{"link": "https://example.com"}], url_columns=["link"]
        ).save()
