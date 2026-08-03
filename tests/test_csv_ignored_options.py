"""CSV/TSV output must say what it is throwing away.

Every Excel-only option is silently discarded on the CSV path — the Rust
``write_csv`` only takes a delimiter and the formula guard. Swapping a target
from ``.xlsx`` to ``.csv`` therefore drops all formatting, so the builder warns
instead of quietly producing a plain file.
"""

import contextlib
import warnings

import pytest

from rustpy_xlsxwriter import FastExcel, Format

ROWS = [{"a": 1.23456, "b": "x"}]


@contextlib.contextmanager
def _no_warnings():
    """Assert nothing warns (``pytest.warns(None)`` is an error in pytest 8+)."""
    with warnings.catch_warnings(record=True) as log:
        warnings.simplefilter("always")
        yield
    assert not log, f"unexpected warnings: {[str(x.message) for x in log]}"


def test_no_warning_when_nothing_excel_only_is_set(tmp_path):
    path = tmp_path / "plain.csv"
    with _no_warnings():
        FastExcel(str(path)).sheet("S", ROWS).save()
    assert path.read_text().strip() == "a,b\n1.23456,x"


def test_autofit_alone_does_not_warn(tmp_path):
    """autofit defaults to True, so warning on it would fire on every CSV."""
    with _no_warnings():
        FastExcel(str(tmp_path / "a.csv"), autofit=True).sheet("S", ROWS).save()


def test_sanitize_formulas_does_not_warn(tmp_path):
    """It is a CSV-only option, not an ignored one."""
    with _no_warnings():
        FastExcel(str(tmp_path / "s.csv"), sanitize_formulas=True).sheet(
            "S", ROWS
        ).save()


def test_warns_and_names_the_ignored_options(tmp_path):
    path = tmp_path / "fmt.csv"
    with pytest.warns(UserWarning) as rec:
        (
            FastExcel(str(path), password="s3cret")
            .format(float_format="0.00", bold_headers=True)
            .freeze(row=1)
            .sheet("S", ROWS, column_widths={"a": 20}, banded_rows="#EEEEEE")
            .save()
        )

    message = str(rec[0].message)
    for name in (
        "password",
        "float_format",
        "bold_headers",
        "freeze",
        "column_widths",
        "banded_rows",
    ):
        assert name in message, f"{name} missing from warning"
    # The data is still written correctly, unformatted.
    assert path.read_text().strip() == "a,b\n1.23456,x"


@pytest.mark.parametrize(
    "configure",
    [
        pytest.param(lambda f, rows: f.sheet("S", rows, header_row=1), id="header_row"),
        pytest.param(
            lambda f, rows: f.sheet("S", rows, row_heights={0: 20}), id="row_heights"
        ),
        pytest.param(
            lambda f, rows: f.sheet("S", rows, row_formats={0: Format().set_bold()}),
            id="row_formats",
        ),
        pytest.param(
            lambda f, rows: f.sheet("S", rows, dedupe_strings=True), id="dedupe_strings"
        ),
        pytest.param(
            lambda f, rows: f.sheet("S", rows, header_format=Format().set_bold()),
            id="header_format",
        ),
        pytest.param(
            lambda f, rows: f.sheet("S", rows, column_formats={"a": Format()}),
            id="column_formats",
        ),
        pytest.param(
            lambda f, rows: f.sheet("S", rows, column_width=15.0), id="column_width"
        ),
        pytest.param(
            lambda f, rows: f.format(datetime_format="yyyy").sheet("S", rows),
            id="datetime_format",
        ),
        pytest.param(
            lambda f, rows: f.format(index_columns=["a"]).sheet("S", rows),
            id="index_columns",
        ),
    ],
)
def test_each_excel_only_option_warns(tmp_path, configure):
    """Guards against a new option being added without joining the warning."""
    f = FastExcel(str(tmp_path / "each.csv"))
    with pytest.warns(UserWarning, match="ignores Excel-only options"):
        configure(f, ROWS).save()


def test_tsv_warns_too(tmp_path):
    with pytest.warns(UserWarning, match="ignores Excel-only options"):
        FastExcel(str(tmp_path / "o.tsv")).format(float_format="0.00").sheet(
            "S", ROWS
        ).save()


def test_xlsx_never_warns(tmp_path):
    with _no_warnings():
        (
            FastExcel(str(tmp_path / "o.xlsx"), password="s3cret")
            .format(float_format="0.00")
            .sheet("S", ROWS, banded_rows="#EEEEEE")
            .save()
        )
