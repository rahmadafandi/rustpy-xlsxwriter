"""A CSV written to a path replaces it atomically.

The rows go to a sibling temporary file that is renamed onto the target when
the write finishes. So a failure part-way leaves whatever was already there
untouched, and a reader never sees half a file.

Some targets cannot be stood in for by a temporary — a symlink would be
replaced by a regular file instead of followed, and a FIFO or a device has no
rename to do — so those are written through directly. That fallback is what
these tests mostly guard, because it is the part that is easy to break.
"""

import os
import stat
import threading

import pytest

from rustpy_xlsxwriter import write_csv

ROWS = [{"a": 1, "b": 2}]
# Enough rows to pass the 64 KiB chunk, so the file is really open when the
# last record fails.
BIG = [{"a": "x" * 40} for _ in range(20_000)]


def test_the_old_file_survives_a_late_failure(tmp_path):
    path = tmp_path / "out.csv"
    path.write_text("IMPORTANT,DATA\n1,2\n")
    with pytest.raises(ValueError):
        write_csv(BIG + [{"a": 1, "UNKNOWN": 2}], str(path))
    assert path.read_text() == "IMPORTANT,DATA\n1,2\n"


def test_nothing_extra_is_left_in_the_directory(tmp_path):
    path = tmp_path / "out.csv"
    path.write_text("old\n")
    with pytest.raises(ValueError):
        write_csv(BIG + [{"a": 1, "UNKNOWN": 2}], str(path))
    assert [p.name for p in tmp_path.iterdir()] == ["out.csv"]


def test_a_successful_write_still_replaces_the_file(tmp_path):
    path = tmp_path / "out.csv"
    path.write_text("old\n")
    write_csv(ROWS, str(path))
    assert path.read_text() == "a,b\n1,2\n"


def test_the_target_keeps_its_permissions(tmp_path):
    """The rename brings a new inode, which would otherwise arrive with the
    default umask rather than the mode the caller had set."""
    path = tmp_path / "out.csv"
    path.write_text("old\n")
    path.chmod(0o640)
    write_csv(ROWS, str(path))
    assert stat.S_IMODE(path.stat().st_mode) == 0o640


def test_a_new_file_is_created(tmp_path):
    path = tmp_path / "fresh.csv"
    write_csv(ROWS, str(path))
    assert path.read_text() == "a,b\n1,2\n"


def test_a_large_write_still_streams(tmp_path):
    """The temporary file must not bring the whole output back into memory."""
    import resource

    path = tmp_path / "big.csv"
    # Built before the baseline is read: the list itself costs tens of MB, and
    # measuring across its construction would hide whatever the writer does.
    records = [{"a": i, "s": "y" * 60} for i in range(200_000)]
    before = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    write_csv(records, str(path))
    growth_mb = (resource.getrusage(resource.RUSAGE_SELF).ru_maxrss - before) / 1024
    assert path.stat().st_size > 10 * 1024 * 1024
    assert growth_mb < 5, f"peak RSS grew {growth_mb:.1f} MB"


# --- targets that cannot be renamed into place -----------------------------


def test_a_symlink_is_followed_not_replaced(tmp_path):
    real, link = tmp_path / "real.csv", tmp_path / "link.csv"
    real.write_text("old\n")
    link.symlink_to(real)

    write_csv(ROWS, str(link))

    assert link.is_symlink()
    assert real.read_text() == "a,b\n1,2\n"


def test_a_fifo_receives_the_rows(tmp_path):
    fifo = tmp_path / "pipe.csv"
    os.mkfifo(fifo)
    got = []

    reader = threading.Thread(target=lambda: got.append(fifo.read_text()))
    reader.start()
    write_csv([{"a": 1}, {"a": 2}], str(fifo))
    reader.join(timeout=10)

    assert got == ["a\n1\n2\n"]
    assert stat.S_ISFIFO(os.stat(fifo).st_mode)


def test_dev_stdout_is_written_through(capfd):
    write_csv([{"a": 1}, {"a": 2}], "/dev/stdout")
    assert capfd.readouterr().out.endswith("a\n1\n2\n")
