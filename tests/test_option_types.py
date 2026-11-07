"""Keep the option ``TypedDict``s in step with the keys Rust accepts.

``_types`` exists so a checker can reject a misspelled key, which is only true
while it lists the same keys the parser does. Nothing executes a ``TypedDict``,
so without this the two drift apart the first time an option gains a key: Rust
would accept it and every editor would flag it as a mistake.

The Rust ``KEYS`` arrays are the source of truth, read straight from the source
rather than mirrored here, so adding a key in one place and not the other is
what fails.
"""

import pathlib
import re

import pytest

from rustpy_xlsxwriter import _types

SRC = pathlib.Path(__file__).resolve().parents[1] / "src"

pytestmark = pytest.mark.skipif(
    not SRC.is_dir(), reason="needs the Rust sources, not just the installed wheel"
)

# TypedDict -> (file, name of the const holding its keys)
MIRRORED = {
    "PageSetup": ("page_setup.rs", "KEYS"),
    "Margins": ("page_setup.rs", "MARGIN_KEYS"),
    "SheetView": ("sheet_view.rs", "KEYS"),
    "Outline": ("outline.rs", "KEYS"),
    "Note": ("notes.rs", "KEYS"),
    "Image": ("images.rs", "KEYS"),
    "Chart": ("charts.rs", "KEYS"),
    "Sparkline": ("sparklines.rs", "KEYS"),
    "DataValidation": ("data_validation.rs", "KEYS"),
    "ConditionalRule": ("conditional_format.rs", "KEYS"),
}


def _rust_keys(filename: str, const: str) -> set:
    source = (SRC / filename).read_text()
    match = re.search(
        rf"const {const}: \[&str; \d+\] = \[(.*?)\];", source, re.DOTALL
    )
    assert match, f"{const} not found in {filename}"
    return set(re.findall(r'"([^"]+)"', match.group(1)))


@pytest.mark.parametrize("name", sorted(MIRRORED))
def test_typed_dict_matches_the_rust_keys(name):
    filename, const = MIRRORED[name]
    assert set(getattr(_types, name).__annotations__) == _rust_keys(filename, const)


def test_outline_groups_match_their_inline_keys():
    """These keys are an inline array in ``group_parts``, not a named const."""
    source = (SRC / "outline.rs").read_text()
    match = re.search(r'if !\[(.*?)\]\.contains', source, re.DOTALL)
    assert match, "the group key list moved"
    keys = set(re.findall(r'"([^"]+)"', match.group(1)))

    assert set(_types.OutlineRowGroup.__annotations__) == keys
    assert set(_types.OutlineColumnGroup.__annotations__) == keys


def test_every_typed_dict_is_optional_per_key():
    """Which keys a rule needs depends on its ``type``, so none are required."""
    for name in MIRRORED:
        assert not getattr(_types, name).__required_keys__, name


def test_exported_from_the_package():
    """They are only useful if callers can import them to annotate their own dicts."""
    import rustpy_xlsxwriter

    for name in _types.__all__:
        assert name in rustpy_xlsxwriter.__all__, name
        assert getattr(rustpy_xlsxwriter, name) is getattr(_types, name)
