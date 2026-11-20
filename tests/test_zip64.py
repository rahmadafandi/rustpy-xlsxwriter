"""Every workbook is written as ZIP64, so a huge export does not die at 4GB.

A plain zip entry stops at 4GB uncompressed, which is roughly 150 million
cells. Past that the write failed after several minutes with ``Large file
option has not been set``, and then a second, more confusing error about
seeking, because the packager streams the zip and cannot rewind to abort the
entry it had already started.

It cannot be switched on once the limit is reached: in constant-memory mode
the final size is unknown until the sheet is written, and by then the zip is
half out of the door. So it is always on. The cost is a flat 432 bytes.

Checking the flag directly beats writing 4GB in a test: the ZIP64 extra field
(header id ``0x0001``) is present on every entry exactly when the option is
set.
"""

import io
import struct
import zipfile

import openpyxl
import pytest

from rustpy_xlsxwriter import FastExcel, write_worksheet, write_worksheets

ROWS = [{"a": 1, "b": "x"}, {"a": 2, "b": "y"}]


def _has_zip64_extra(info: zipfile.ZipInfo) -> bool:
    extra, offset = info.extra, 0
    while offset + 4 <= len(extra):
        header_id, size = struct.unpack("<HH", extra[offset : offset + 4])
        if header_id == 0x0001:
            return True
        offset += 4 + size
    return False


def _all_entries_are_zip64(path_or_buffer) -> bool:
    entries = zipfile.ZipFile(path_or_buffer).infolist()
    assert entries, "the workbook has no zip entries"
    return all(_has_zip64_extra(entry) for entry in entries)


def test_write_worksheet(tmp_path):
    path = tmp_path / "one.xlsx"
    write_worksheet(ROWS, str(path))
    assert _all_entries_are_zip64(path)


def test_write_worksheets(tmp_path):
    path = tmp_path / "many.xlsx"
    write_worksheets([("A", ROWS), ("B", ROWS)], str(path))
    assert _all_entries_are_zip64(path)


def test_fast_excel(tmp_path):
    path = tmp_path / "fast.xlsx"
    FastExcel(str(path)).sheet("S", ROWS).save()
    assert _all_entries_are_zip64(path)


def test_a_buffer_target_too(tmp_path):
    """Buffers go through `save_to_buffer`, a different call."""
    buffer = io.BytesIO()
    write_worksheet(ROWS, buffer)
    buffer.seek(0)
    assert _all_entries_are_zip64(buffer)


def test_the_sheet_that_opts_out_of_constant_memory(tmp_path):
    """`dedupe_strings=True` takes a different worksheet path to the packager."""
    path = tmp_path / "dedupe.xlsx"
    write_worksheet(ROWS, str(path), dedupe_strings=True)
    assert _all_entries_are_zip64(path)


@pytest.mark.parametrize("level", [1, 9])
def test_it_survives_a_compression_level(tmp_path, level):
    path = tmp_path / f"level{level}.xlsx"
    write_worksheet(ROWS, str(path), compression_level=level)
    assert _all_entries_are_zip64(path)


def test_the_file_still_reads_normally(tmp_path):
    """ZIP64 is from 2001, but assert it rather than assume it."""
    path = tmp_path / "readable.xlsx"
    write_worksheet(ROWS, str(path))

    assert zipfile.ZipFile(path).testzip() is None
    sheet = openpyxl.load_workbook(path).active
    assert [[c.value for c in row] for row in sheet.iter_rows()] == [
        ["a", "b"],
        [1, "x"],
        [2, "y"],
    ]
