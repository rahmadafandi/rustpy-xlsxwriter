"""`write_csv` to a path streams instead of holding the file in memory.

It used to build the whole output in one buffer and write it at the end,
so peak memory grew by exactly the size of the file. A path target now
receives the output in chunks as the rows are produced, and a file-like
target still gets one `bytes` object.

Streaming means part of the file can reach disk before an error is found,
and two things rewind output: a record that misses the pointer fast path,
and the Arrow path, which on failure starts over on the per-object path.
Each test that needs it writes well past the 64 KiB chunk, so the rewind
happens after earlier chunks are already on disk.
"""

import io
import subprocess
import sys
import textwrap

import pytest

from rustpy_xlsxwriter import write_csv

ROWS = 20_000  # about 600 KB of CSV: several chunks


def _records(n=ROWS):
    return [{"id": i, "name": f"customer {i}", "score": i * 0.5} for i in range(n)]


def _to_bytes(data, **kwargs):
    buf = io.BytesIO()
    write_csv(data, buf, **kwargs)
    return buf.getvalue()


@pytest.mark.parametrize("bom", [False, True])
def test_path_output_matches_the_buffer_output(tmp_path, bom):
    path = tmp_path / "out.csv"
    write_csv(_records(), str(path), bom=bom)
    data = path.read_bytes()
    assert len(data) > 4 * 64 * 1024
    assert data == _to_bytes(_records(), bom=bom)


def test_reordered_records_rewind_within_a_chunk(tmp_path):
    """A reordered record misses the pointer fast path and is rewound and
    written again by name. That must still work in the middle of a stream."""
    records = _records()
    for i in range(1, ROWS, 997):
        records[i] = {"score": records[i]["score"], "name": records[i]["name"], "id": i}
    path = tmp_path / "out.csv"
    write_csv(records, path)
    assert path.read_bytes() == _to_bytes(_records())


def test_polars_frame_streams_through_the_arrow_path(tmp_path):
    pl = pytest.importorskip("polars")
    df = pl.DataFrame(_records())
    path = tmp_path / "out.csv"
    write_csv(df, path)
    assert path.read_bytes() == _to_bytes(df)
    assert path.read_bytes() == _to_bytes(_records())


class _FailingArrowFrame:
    """A pandas frame whose Arrow stream breaks after some batches.

    The writer then starts over on the `.values` path. By then several chunks
    of Arrow output are on disk, and the restart has to clear them."""

    def __init__(self, df, good_batches):
        self._df = df
        self._good = good_batches
        self.columns = df.columns
        self.values = df.values

    def __arrow_c_stream__(self, requested_schema=None):
        import pyarrow as pa

        table = pa.Table.from_pandas(self._df, preserve_index=False)
        batches = table.to_batches(max_chunksize=1000)

        def gen():
            yield from batches[: self._good]
            raise RuntimeError("stream broke")

        reader = pa.RecordBatchReader.from_batches(table.schema, gen())
        return reader.__arrow_c_stream__(requested_schema)


@pytest.mark.parametrize("bom", [False, True])
def test_arrow_retry_clears_what_already_reached_the_file(tmp_path, bom):
    pd = pytest.importorskip("pandas")
    pytest.importorskip("pyarrow")
    df = pd.DataFrame(_records())
    path = tmp_path / "out.csv"
    write_csv(_FailingArrowFrame(df, good_batches=10), path, bom=bom)
    data = path.read_bytes()
    assert data.count(b"\xef\xbb\xbf") == (1 if bom else 0)
    assert data == _to_bytes(_FailingArrowFrame(df, good_batches=10), bom=bom)
    assert data.decode("utf-8-sig").splitlines()[:2] == ["id,name,score", "0,customer 0,0.0"]
    assert len(data.splitlines()) == ROWS + 1


def test_error_after_several_chunks_leaves_no_file(tmp_path):
    """A truncated CSV reads as a complete one with fewer rows, so a write
    that fails part-way removes what it wrote."""
    records = _records()
    records.append({"id": 0, "unexpected": 1})
    path = tmp_path / "out.csv"
    with pytest.raises(ValueError, match="no column 'unexpected'"):
        write_csv(records, path)
    assert not path.exists()


def test_error_before_the_first_chunk_leaves_an_existing_file_alone(tmp_path):
    """The file is opened at the first chunk, so a write that fails before
    then never touches what was at the path, as before streaming."""
    path = tmp_path / "out.csv"
    path.write_text("previous export\n")
    with pytest.raises(ValueError):
        write_csv([{"a": 1}, {"b": 2}], path)
    assert path.read_text() == "previous export\n"


def test_error_after_the_first_chunk_leaves_an_existing_file_alone(tmp_path):
    """The rows go to a sibling temporary file, so the old export survives
    however far the write got before failing.

    This used to assert the opposite: streaming wrote straight to the target,
    so a failure past the first chunk deleted a file the caller never asked to
    replace, while the same mistake one row earlier left it untouched.
    """
    path = tmp_path / "out.csv"
    path.write_text("previous export\n")
    records = _records() + [{"id": 0, "unexpected": 1}]
    with pytest.raises(ValueError):
        write_csv(records, path)
    assert path.read_text() == "previous export\n"


def test_no_temporary_file_is_left_behind(tmp_path):
    path = tmp_path / "out.csv"
    path.write_text("previous export\n")
    records = _records() + [{"id": 0, "unexpected": 1}]
    with pytest.raises(ValueError):
        write_csv(records, path)
    assert [p.name for p in tmp_path.iterdir()] == ["out.csv"]


def test_file_like_target_gets_one_write():
    class Recorder:
        def __init__(self):
            self.calls = []

        def write(self, data):
            self.calls.append(data)

    rec = Recorder()
    write_csv(_records(), rec)
    assert len(rec.calls) == 1
    assert rec.calls[0] == _to_bytes(_records())


def test_bad_target_is_refused_before_any_row_is_written():
    calls = []

    def records():
        calls.append(1)
        yield {"a": 1}

    with pytest.raises(TypeError, match="path .* or a file-like object"):
        write_csv(records(), 42)
    assert calls == []


def test_missing_directory_is_an_oserror(tmp_path):
    with pytest.raises(OSError, match="Failed to write file"):
        write_csv(_records(10), tmp_path / "missing" / "out.csv")


def test_peak_memory_does_not_grow_with_the_file(tmp_path):
    """Measured in a fresh process: before streaming, peak RSS grew by the
    size of the file (about 31 MB for this one)."""
    script = textwrap.dedent(
        """
        import os, resource, sys
        from rustpy_xlsxwriter import write_csv
        rows = [{"id": i, "name": f"customer name {i}", "score": i * 0.25,
                 "city": "Jakarta", "ok": True} for i in range(400_000)]
        before = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        write_csv(rows, sys.argv[1])
        after = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        print((after - before) * 1024, os.path.getsize(sys.argv[1]))
        """
    )
    path = tmp_path / "big.csv"
    out = subprocess.run(
        [sys.executable, "-c", script, str(path)], capture_output=True, text=True, check=True
    )
    growth, size = map(int, out.stdout.split())
    assert size > 15_000_000
    assert growth < size / 4, f"peak RSS grew {growth} bytes for a {size}-byte file"
