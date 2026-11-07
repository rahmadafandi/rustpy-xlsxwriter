"""Data whose columns repeat a name is refused, not guessed at.

Every option keyed by column name — ``column_formats``, ``column_widths``,
``url_columns``, conditional formats — would match both columns, which is why
a formula column repeating a data header already raises. A pandas frame is
worse than ambiguous: ``df["a"]`` returns a DataFrame rather than a Series
when ``a`` appears twice, so the Excel writer used to fail from inside its
column loop with ``AttributeError: 'DataFrame' object has no attribute
'tolist'`` — an internal leak that said nothing about the real problem.
"""

import pandas as pd
import pytest

from rustpy_xlsxwriter import FastExcel, write_csv, write_worksheet

MESSAGE = r"two columns named 'a'"


def _dup():
    return pd.DataFrame([[1, 2, 3]], columns=["a", "a", "b"])


def test_excel_names_the_repeated_column(tmp_path):
    """Used to be an AttributeError about tolist, from inside the column loop."""
    with pytest.raises(ValueError, match=MESSAGE):
        write_worksheet(_dup(), str(tmp_path / "out.xlsx"))


def test_the_message_says_why_it_is_refused(tmp_path):
    with pytest.raises(ValueError, match=r"keyed by column name ambiguous"):
        write_worksheet(_dup(), str(tmp_path / "out.xlsx"))


def test_csv_names_the_repeated_column(tmp_path):
    with pytest.raises(ValueError, match=MESSAGE):
        write_csv(_dup(), str(tmp_path / "out.csv"))


def test_fast_excel_too(tmp_path):
    with pytest.raises(ValueError, match=MESSAGE):
        FastExcel(str(tmp_path / "out.xlsx")).sheet("S", _dup()).save()


def test_selecting_the_other_column_does_not_excuse_it(tmp_path):
    """The frame cannot be indexed by name at all, so a subset is no escape."""
    with pytest.raises(ValueError, match=MESSAGE):
        write_csv(_dup(), str(tmp_path / "out.csv"), columns=["b"])


def test_an_option_keyed_by_the_repeated_name(tmp_path):
    """The case the message describes, reached before any cell is written."""
    with pytest.raises(ValueError, match=MESSAGE):
        write_worksheet(
            _dup(),
            str(tmp_path / "out.xlsx"),
            column_formats={"a": {"bold": True}},
        )


def test_no_file_is_left_behind(tmp_path):
    """The check runs on the headers, before anything is opened."""
    path = tmp_path / "out.csv"
    with pytest.raises(ValueError):
        write_csv(_dup(), str(path))
    assert not path.exists()


def test_unique_columns_still_write(tmp_path):
    """The guard must not fire on an ordinary frame."""
    path = tmp_path / "ok.csv"
    write_csv(pd.DataFrame([[1, 2]], columns=["a", "b"]), str(path))
    assert path.read_text().strip() == "a,b\n1,2"


def test_a_repeat_in_columns_keeps_its_own_message(tmp_path):
    """Listing one name twice is the caller's mistake, not the data's."""
    with pytest.raises(ValueError, match=r"'a' is listed more than once"):
        write_csv(
            pd.DataFrame([[1, 2]], columns=["a", "b"]),
            str(tmp_path / "out.csv"),
            columns=["a", "a"],
        )
