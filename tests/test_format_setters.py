"""Every `Format` setter, exercised once.

The setters are generated from four lists in a macro, so the mechanism is
proven by any one of them. What is not proven is the wiring: a setter named in
the wrong list, or pointed at the wrong `parse_*`, compiles and produces a
file that is quietly wrong. Seventeen of the forty-one were reached by other
tests; the rest were reached by nothing at all.

`test_every_setter_is_covered_here` is what keeps that from happening again —
a setter added to the macro without an entry below fails immediately.
"""

import re
import zipfile

import openpyxl
import pytest

from rustpy_xlsxwriter import Format, write_worksheet

# name -> (args, read a cell with openpyxl, expected)
# A reader of None means openpyxl does not model it; the call is still made,
# which is what catches a setter wired into the wrong macro list.
SETTERS = {
    # -- flags, no argument --------------------------------------------------
    "set_bold": ((), lambda c: c.font.bold, True),
    "set_italic": ((), lambda c: c.font.italic, True),
    "set_text_wrap": ((), lambda c: c.alignment.wrapText, True),
    "set_shrink": ((), lambda c: c.alignment.shrinkToFit, True),
    "set_font_strikethrough": ((), lambda c: c.font.strikethrough, True),
    # Locked is Excel's default and rust_xlsxwriter omits it, so the file
    # looks identical to one with no format at all. What it does mean is
    # tested by test_locked_undoes_unlocked.
    "set_locked": ((), None, None),
    "set_unlocked": ((), lambda c: c.protection.locked, False),
    "set_hidden": ((), lambda c: c.protection.hidden, True),
    "set_quote_prefix": ((), None, None),
    "set_checkbox": ((), None, None),
    "set_hyperlink": ((), lambda c: c.font.underline, "single"),
    # -- one primitive argument ---------------------------------------------
    "set_font_size": ((14.0,), lambda c: c.font.size, 14.0),
    "set_font_family": ((3,), lambda c: c.font.family, 3.0),
    "set_font_charset": ((1,), lambda c: c.font.charset, 1),
    "set_rotation": ((45,), lambda c: c.alignment.textRotation, 45),
    "set_indent": ((2,), lambda c: c.alignment.indent, 2.0),
    "set_reading_direction": ((2,), lambda c: c.alignment.readingOrder, 2.0),
    # Index 9 is Excel's built-in percentage format.
    "set_num_format_index": ((9,), lambda c: c.number_format, "0%"),
    # -- one string argument -------------------------------------------------
    "set_num_format": (("0.00",), lambda c: c.number_format, "0.00"),
    "set_font_name": (("Courier New",), lambda c: c.font.name, "Courier New"),
    "set_underline": (("double",), lambda c: c.font.underline, "double"),
    # -- colours -------------------------------------------------------------
    "set_font_color": (("#FF0000",), lambda c: c.font.color.rgb, "FFFF0000"),
    "set_background_color": (("#00FF00",), lambda c: c.fill.fgColor.rgb, "FF00FF00"),
    "set_foreground_color": (("#0000FF",), None, None),
    "set_border_color": (("#FF0000",), None, None),
    "set_border_top_color": (("#FF0000",), None, None),
    "set_border_bottom_color": (("#FF0000",), None, None),
    "set_border_left_color": (("#FF0000",), None, None),
    "set_border_right_color": (("#FF0000",), None, None),
    "set_border_diagonal_color": (("#FF0000",), None, None),
    # -- parsed vocabularies -------------------------------------------------
    "set_align": (("center",), lambda c: c.alignment.horizontal, "center"),
    "set_border": (("thin",), lambda c: c.border.top.style, "thin"),
    "set_border_top": (("thick",), lambda c: c.border.top.style, "thick"),
    "set_border_bottom": (("dashed",), lambda c: c.border.bottom.style, "dashed"),
    "set_border_left": (("dotted",), lambda c: c.border.left.style, "dotted"),
    "set_border_right": (("double",), lambda c: c.border.right.style, "double"),
    "set_border_diagonal": (("thin",), lambda c: c.border.diagonal.style, "thin"),
    "set_border_diagonal_type": (("border_up",), lambda c: c.border.diagonalUp, True),
    "set_pattern": (("solid",), lambda c: c.fill.patternType, "solid"),
    # "headings" is written as Excel's "major". Not "body", which maps to
    # "minor" — already the default font's scheme, so it would assert nothing.
    "set_font_scheme": (("headings",), lambda c: c.font.scheme, "major"),
    "set_font_script": (("superscript",), lambda c: c.font.vertAlign, "superscript"),
}

# parsed setter -> a value outside its vocabulary
BAD_VALUES = {
    "set_align": "sideways",
    "set_border": "wiggly",
    "set_border_diagonal_type": "border_sideways",
    "set_pattern": "polkadot",
    "set_font_scheme": "major",  # the Excel name, not the one this API takes
    "set_font_script": "sideways",
    "set_font_color": "not-a-colour",
    "set_underline": "wiggly",
}


def _cell(tmp_path, name, args):
    """Apply one setter to a column and read the data cell back."""
    fmt = getattr(Format(), name)(*args)
    path = tmp_path / f"{name}.xlsx"
    write_worksheet([{"a": "x"}], str(path), column_formats={"a": fmt})
    return path, openpyxl.load_workbook(path).active.cell(2, 1)


# --- the guard --------------------------------------------------------------


def test_every_setter_is_covered_here():
    """A setter added to the macro without an entry above fails here."""
    live = {n for n in dir(Format) if n.startswith("set_")}
    assert live == set(SETTERS), {
        "missing from this file": sorted(live - set(SETTERS)),
        "no longer on Format": sorted(set(SETTERS) - live),
    }


# --- each setter ------------------------------------------------------------


@pytest.mark.parametrize("name", sorted(SETTERS))
def test_setter_applies(tmp_path, name):
    args, read, expected = SETTERS[name]
    _, cell = _cell(tmp_path, name, args)
    if read is None:
        return  # openpyxl does not model it; reaching here proves the wiring
    assert read(cell) == expected


def test_quote_prefix_reaches_the_file(tmp_path):
    """openpyxl drops it, so this one is read from the styles part."""
    path, _ = _cell(tmp_path, "set_quote_prefix", ())
    styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
    assert 'quotePrefix="1"' in styles


def test_locked_undoes_unlocked(tmp_path):
    """`set_locked` writes nothing on its own; what it does is cancel this."""
    def protection(fmt, name):
        path = tmp_path / f"{name}.xlsx"
        write_worksheet([{"a": "x"}], str(path), column_formats={"a": fmt})
        styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
        return re.findall(r"<protection[^>]*/>", styles)

    assert protection(Format().set_unlocked(), "unlocked") == ['<protection locked="0"/>']
    assert protection(Format().set_unlocked().set_locked(), "relocked") == []


def test_setters_chain(tmp_path):
    """Each returns self, so a Format is built in one expression."""
    fmt = Format().set_bold().set_italic().set_font_size(12.0).set_align("center")
    path = tmp_path / "chained.xlsx"
    write_worksheet([{"a": "x"}], str(path), column_formats={"a": fmt})
    cell = openpyxl.load_workbook(path).active.cell(2, 1)
    assert (cell.font.bold, cell.font.italic, cell.font.size) == (True, True, 12.0)
    assert cell.alignment.horizontal == "center"


# --- the vocabularies reject what they do not know --------------------------


@pytest.mark.parametrize("name", sorted(BAD_VALUES))
def test_parsed_setter_rejects_an_unknown_value(name):
    """Each `parse_*` names its vocabulary, so a typo says what was allowed."""
    with pytest.raises(ValueError, match="(?i)invalid"):
        getattr(Format(), name)(BAD_VALUES[name])


def test_setters_are_positional_only():
    """pyo3 names every macro-generated argument `value`; see test_type_stubs."""
    with pytest.raises(TypeError, match="unexpected keyword argument"):
        Format().set_font_size(size=12.0)
