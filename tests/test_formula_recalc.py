"""End-to-end formula check: a real spreadsheet engine computes the results.

The unit tests in ``test_formulas.py`` prove the formula *text* survives into
the file. They cannot prove the file is one a spreadsheet actually understands
— a wrong ``_xlfn.`` prefix or missing dynamic-array metadata still produces
plausible-looking XML. Here LibreOffice opens each file headless, recalculates,
and we assert the computed values.

Skipped automatically when LibreOffice is not installed, so CI without it stays
green. Run just these with ``pytest -m recalc``.
"""

import shutil
import subprocess

import openpyxl
import pytest

from rustpy_xlsxwriter import write_worksheet

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

pytestmark = [
    pytest.mark.recalc,
    pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed"),
]


def _recalculate(path, tmp_path):
    """Round-trip through LibreOffice and return the recalculated sheet."""
    outdir = tmp_path / "recalc"
    outdir.mkdir(exist_ok=True)
    # A private profile dir keeps this from touching the user's LibreOffice
    # config and lets concurrent runs coexist.
    profile = tmp_path / "loprofile"
    result = subprocess.run(
        [
            SOFFICE,
            "--headless",
            "--norestore",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(outdir),
            str(path),
        ],
        capture_output=True,
        timeout=240,
    )
    converted = outdir / path.name
    if not converted.exists():
        raise AssertionError(
            f"LibreOffice did not produce output: {result.stdout!r} {result.stderr!r}"
        )
    return openpyxl.load_workbook(converted, data_only=True).active


ROWS = [{"qty": 1, "price": 2.0}, {"qty": 2, "price": 4.0}, {"qty": 3, "price": 6.0}]


@pytest.fixture(scope="module")
def computed(tmp_path_factory):
    """One conversion for the whole module — LibreOffice startup is slow."""
    tmp = tmp_path_factory.mktemp("recalc")
    path = tmp / "formulas.xlsx"
    write_worksheet(
        ROWS,
        str(path),
        formula_columns={
            "product": "=A{row}*B{row}",
            "running": "=SUM(B${first}:B{row})",
            "branch": '=IF(A{row}>2,"big",IF(A{row}>1,"mid","small"))',
            "rounded": "=ROUND(B{row}/A{row},2)",
            "joined": '=A{row}&"x"&B{row}',
            "lookup": "=INDEX(B:B,MATCH(A{row},A:A,0))",
            "conditional": '=SUMIFS(B$2:B$4,A$2:A$4,">"&A{row})',
            "modern_ifs": '=IFS(A{row}>2,"hi",A{row}>1,"mid",TRUE,"lo")',
            "modern_join": '=TEXTJOIN("-",TRUE,A${first}:A{row})',
            "modern_max": '=MAXIFS(B$2:B$4,A$2:A$4,"<="&A{row})',
        },
        totals_row={
            "qty": "sum",
            "price": "=ROUND(AVERAGE({col}{first}:{col}{last}),2)",
        },
        totals_label=None,
    )
    return _recalculate(path, tmp)


@pytest.mark.parametrize(
    "cell,expected",
    [
        # =A2*B2 -> 1*2
        ("C2", 2),
        ("C4", 18),
        # running sum of price
        ("D2", 2),
        ("D4", 12),
        # nested IF
        ("E2", "small"),
        ("E3", "mid"),
        ("E4", "big"),
        # ROUND(price/qty, 2)
        ("F2", 2),
        ("F4", 2),
        # string concatenation
        ("G2", "1x2"),
        # INDEX/MATCH
        ("H3", 4),
        # SUMIFS with a comparison built by concatenation
        ("I2", 10),
        ("I4", 0),
        # future functions — these are the ones needing the _xlfn. prefix
        ("J2", "lo"),
        ("J4", "hi"),
        ("K2", "1"),
        ("K4", "1-2-3"),
        ("L4", 6),
    ],
)
def test_computed_values(computed, cell, expected):
    assert computed[cell].value == expected


def test_totals_row_computes(computed):
    """SUM aggregate and a free-form ROUND(AVERAGE(...)) side by side."""
    assert computed["A5"].value == 6  # 1+2+3
    assert computed["B5"].value == 4  # (2+4+6)/3


def test_dynamic_arrays_recalculate_or_are_unsupported(tmp_path):
    """A dynamic array either computes to the right value, or the engine says so.

    LibreOffice gained SORT/UNIQUE/XLOOKUP somewhere between 24.2 (``#NAME?``)
    and 26.2 (computes), so the cutoff is not worth pinning to a version. What
    holds on every engine: if SORT runs at all it yields the first sorted value,
    never the ``#VALUE?``/``#REF!``/``#SPILL!`` that malformed dynamic-array
    metadata would produce. The structural side — ``_xlfn.`` prefix,
    ``t="array"``, ``cm="1"``, ``xl/metadata.xml`` — is covered in
    ``test_formulas.py``.
    """
    path = tmp_path / "dynamic.xlsx"
    write_worksheet(
        ROWS,
        str(path),
        formula_columns={"sorted": "=SORT(A$2:A$4)", "plain": "=SUM(B$2:B$4)"},
    )
    sheet = _recalculate(path, tmp_path)

    # ROWS qty is already 1,2,3, so a correct SORT spills 1 into the first cell.
    assert sheet["C2"].value in (1, "#NAME?")
    # An ordinary formula in the same file always computes, so a #NAME? above is
    # the engine's missing function, not a broken file.
    assert sheet["D2"].value == 12


def test_no_error_values_anywhere(tmp_path):
    """Guards against a formula landing as #NAME? because of a bad prefix."""
    path = tmp_path / "clean.xlsx"
    write_worksheet(
        ROWS,
        str(path),
        formula_columns={
            "a": "=CONCAT(A${first}:A{row})",
            "b": '=IFNA(VLOOKUP(A{row},A:B,2,FALSE),"missing")',
            "c": "=STDEV.P(B$2:B$4)",
        },
    )
    sheet = _recalculate(path, tmp_path)
    errors = []
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("#"):
                errors.append((cell.coordinate, cell.value))
    assert not errors, f"formula errors in output: {errors}"
