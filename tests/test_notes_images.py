"""Header-cell notes and cell-anchored images.

Both are drawings rather than cell values, which is why neither takes the
sheet out of constant-memory mode the way a row group has to — asserted
below, since that was worth confirming rather than assuming.
"""

import struct
import zipfile
import zlib

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [{"revenue": 1, "sku": "a"}, {"revenue": 2, "sku": "b"}]


def _png(width=8, height=8, colour=b"\xff\x00\x00"):
    """A minimal valid PNG, so the tests need no binary fixture on disk."""

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + colour * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _write(tmp_path, **kwargs):
    path = tmp_path / "n.xlsx"
    write_worksheet(ROWS, str(path), **kwargs)
    return path


def _parts(path, needle):
    return [n for n in zipfile.ZipFile(path).namelist() if needle in n]


# --- notes ------------------------------------------------------------------


def test_note_lands_on_the_header_cell(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, notes={"revenue": "Net of returns"})).active
    assert ws["A1"].comment is not None
    assert "Net of returns" in ws["A1"].comment.text


def test_note_picks_the_named_column(tmp_path):
    ws = openpyxl.load_workbook(_write(tmp_path, notes={"sku": "Internal code"})).active
    assert ws["A1"].comment is None
    assert "Internal code" in ws["B1"].comment.text


def test_note_follows_header_row(tmp_path):
    ws = openpyxl.load_workbook(
        _write(tmp_path, notes={"revenue": "hi"}, header_row=2)
    ).active
    assert ws["A3"].comment is not None


def test_note_dict_form_sets_the_author(tmp_path):
    ws = openpyxl.load_workbook(
        _write(tmp_path, notes={"revenue": {"text": "hi", "author": "ops"}})
    ).active
    assert ws["A1"].comment.author == "ops"


def test_note_survives_constant_memory(tmp_path):
    """It is stored beside the cell data, but the default writer keeps it."""
    assert _parts(_write(tmp_path, notes={"revenue": "hi"}), "comments")


def test_note_unknown_column_warns(tmp_path):
    with pytest.warns(UserWarning, match="notes: unknown column 'nope'"):
        path = _write(tmp_path, notes={"nope": "hi"})
    assert not _parts(path, "comments")


def test_note_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match="notes: unknown key 'colour'"):
        _write(tmp_path, notes={"revenue": {"text": "hi", "colour": "red"}})


def test_note_dict_needs_text(tmp_path):
    with pytest.raises(ValueError, match="'revenue' needs 'text'"):
        _write(tmp_path, notes={"revenue": {"author": "ops"}})


def test_notes_must_be_a_dict(tmp_path):
    with pytest.raises(ValueError, match="notes must be a dict keyed by column name"):
        _write(tmp_path, notes=["revenue"])


# --- images -----------------------------------------------------------------


def test_image_from_a_path(tmp_path):
    logo = tmp_path / "logo.png"
    logo.write_bytes(_png())
    path = _write(tmp_path, images=[{"path": str(logo), "row": 0, "col": 3}])
    assert _parts(path, "media") == ["xl/media/image1.png"]
    assert _parts(path, "drawings/drawing1.xml")


def test_image_from_bytes(tmp_path):
    """What a web handler has: a logo in memory with no file to point at."""
    path = _write(tmp_path, images=[{"data": _png(), "row": 0, "col": 0}])
    assert _parts(path, "media") == ["xl/media/image1.png"]


def test_identical_images_are_stored_once(tmp_path):
    same = _png()
    path = _write(
        tmp_path,
        images=[{"data": same, "row": 0, "col": 0}, {"data": same, "row": 5, "col": 0}],
    )
    assert len(_parts(path, "media")) == 1


def test_different_images_are_both_stored(tmp_path):
    path = _write(
        tmp_path,
        images=[
            {"data": _png(colour=b"\xff\x00\x00"), "row": 0, "col": 0},
            {"data": _png(colour=b"\x00\xff\x00"), "row": 5, "col": 0},
        ],
    )
    assert len(_parts(path, "media")) == 2


def test_image_survives_constant_memory(tmp_path):
    """A drawing anchored to a cell, not a cell write — nothing to revisit."""
    assert _parts(_write(tmp_path, images=[{"data": _png()}]), "media")


def test_image_scale_and_alt_text_are_accepted(tmp_path):
    path = _write(
        tmp_path,
        images=[{"data": _png(), "scale": 0.5, "alt_text": "Company logo"}],
    )
    drawing = zipfile.ZipFile(path).read("xl/drawings/drawing1.xml").decode()
    assert "Company logo" in drawing


def test_image_fit_to_cell(tmp_path):
    path = _write(
        tmp_path,
        images=[{"data": _png(), "fit_to_cell": True, "keep_aspect_ratio": False}],
    )
    assert _parts(path, "media")


def test_keep_aspect_ratio_without_fit_to_cell_raises(tmp_path):
    """It would be accepted and then do nothing, which is the worse outcome."""
    with pytest.raises(ValueError, match="only applies with 'fit_to_cell'"):
        _write(tmp_path, images=[{"data": _png(), "keep_aspect_ratio": True}])


def test_path_and_data_together_raise(tmp_path):
    logo = tmp_path / "logo.png"
    logo.write_bytes(_png())
    with pytest.raises(ValueError, match="give 'path' or 'data', not both"):
        _write(tmp_path, images=[{"path": str(logo), "data": _png()}])


def test_image_needs_a_source(tmp_path):
    with pytest.raises(ValueError, match="needs 'path' or 'data'"):
        _write(tmp_path, images=[{"row": 0}])


def test_image_unknown_key_raises(tmp_path):
    with pytest.raises(ValueError, match=r"images\[0\]: unknown key 'width'"):
        _write(tmp_path, images=[{"data": _png(), "width": 10}])


def test_image_index_is_named_in_the_error(tmp_path):
    with pytest.raises(ValueError, match=r"images\[1\]: needs 'path' or 'data'"):
        _write(tmp_path, images=[{"data": _png()}, {"row": 2}])


def test_missing_file_raises(tmp_path):
    with pytest.raises(Exception, match="(?i)no such file|not found|io error"):
        _write(tmp_path, images=[{"path": str(tmp_path / "nope.png")}])


def test_images_must_be_a_list(tmp_path):
    with pytest.raises(ValueError, match="each image must be a dict"):
        _write(tmp_path, images=["logo.png"])


# --- every entry point ------------------------------------------------------


def test_multi_sheet_is_keyed_by_sheet_name(tmp_path):
    path = tmp_path / "multi.xlsx"
    write_worksheets(
        [("A", ROWS), ("B", ROWS)],
        str(path),
        notes={"A": {"revenue": "only on A"}},
        images={"A": [{"data": _png()}]},
    )
    book = openpyxl.load_workbook(path)
    assert book["A"]["A1"].comment is not None
    assert book["B"]["A1"].comment is None


def test_builder(tmp_path):
    path = tmp_path / "b.xlsx"
    FastExcel(path).sheet(
        "S", ROWS, notes={"revenue": "hi"}, images=[{"data": _png()}]
    ).save()
    assert _parts(path, "media")
    assert openpyxl.load_workbook(path).active["A1"].comment is not None
