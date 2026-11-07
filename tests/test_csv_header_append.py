"""``header=False`` does not make the write append.

The stub used to say "Set ``False`` to append to an existing file", which is
the one thing it must not be used for: a path target is truncated, so a caller
following that sentence lost the rows already there. Appending works through a
file object the caller opened in append mode, and these tests pin both halves
so the documentation cannot drift back.
"""

from rustpy_xlsxwriter import write_csv

FIRST = [{"a": 1, "b": 2}]
SECOND = [{"a": 3, "b": 4}]


def test_a_path_is_overwritten_not_appended(tmp_path):
    path = tmp_path / "out.csv"
    write_csv(FIRST, str(path))
    write_csv(SECOND, str(path), header=False)
    assert path.read_text() == "3,4\n"


def test_a_file_opened_for_append_appends(tmp_path):
    path = tmp_path / "out.csv"
    write_csv(FIRST, str(path))
    with open(path, "ab") as handle:
        write_csv(SECOND, handle, header=False)
    assert path.read_text() == "a,b\n1,2\n3,4\n"


def test_the_appended_part_carries_no_header(tmp_path):
    """Without `header=False` the second write would repeat the names."""
    path = tmp_path / "out.csv"
    write_csv(FIRST, str(path))
    with open(path, "ab") as handle:
        write_csv(SECOND, handle)
    assert path.read_text() == "a,b\n1,2\na,b\n3,4\n"
