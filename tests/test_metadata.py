from importlib.metadata import metadata, version

from rustpy_xlsxwriter import (
    get_authors,
    get_description,
    get_homepage,
    get_license,
    get_name,
    get_repository,
    get_version,
)


def test_version():
    assert get_version() == version("rustpy-xlsxwriter")


def test_name():
    assert get_name() == "rustpy-xlsxwriter"


def test_authors():
    assert get_authors() == "Rahmad Afandi <rahmadafandiii@gmail.com>"


def test_description():
    assert get_description() == metadata("rustpy-xlsxwriter")["Summary"]


def test_repository():
    assert get_repository() == "https://github.com/rahmadafandi/rustpy-xlsxwriter"


def test_homepage():
    assert get_homepage() == "https://github.com/rahmadafandi/rustpy-xlsxwriter"


def test_license():
    assert get_license() == "MIT"
