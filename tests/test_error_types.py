"""The exception class says what kind of problem it was.

A filesystem refusal and Excel refusing the data need different handling.
Saving used to raise ``OSError`` for both, while the same kinds of failure
caught at write time raised ``RuntimeError``.
"""

import io

import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet

ROWS = [{"a": 1}]


@pytest.mark.parametrize(
    "configure, match",
    [
        pytest.param(
            lambda f: f.sheet("S", ROWS).sheet("S", ROWS),
            "already been used",
            id="sheet_name_reused",
        ),
        pytest.param(
            lambda f: f.sheet("S", [{"a": "x" * 40_000}]),
            "32,767",
            id="string_too_long",
        ),
    ],
)
def test_excel_refusing_the_data_is_a_runtime_error(tmp_path, configure, match):
    with pytest.raises(RuntimeError, match=match):
        configure(FastExcel(str(tmp_path / "out.xlsx"))).save()


def test_excel_errors_are_not_os_errors(tmp_path):
    """The class is what changed, not the message."""
    with pytest.raises(Exception) as caught:
        FastExcel(str(tmp_path / "dup.xlsx")).sheet("S", ROWS).sheet("S", ROWS).save()
    assert not isinstance(caught.value, OSError)


def test_the_filesystem_refusing_is_an_os_error(tmp_path):
    """A missing directory, so no permission setup is needed."""
    with pytest.raises(OSError):
        write_worksheet(ROWS, str(tmp_path / "no" / "such" / "dir" / "out.xlsx"))


def test_a_buffer_target_reports_the_same_way():
    """The target does not change what kind of problem a reused name is."""
    with pytest.raises(RuntimeError, match="already been used"):
        FastExcel(io.BytesIO()).sheet("S", ROWS).sheet("S", ROWS).save()
