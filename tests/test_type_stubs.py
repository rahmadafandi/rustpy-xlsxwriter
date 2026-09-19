"""Keep ``rustpy_xlsxwriter.pyi`` in step with the extension it describes.

The stub is hand-written, so nothing but a test stops it drifting — and it had
drifted, silently, because a stub is never executed. Everything here compares
the stub against the compiled module's own introspection: pyo3 emits a real
``__text_signature__``, so ``inspect.signature`` is the source of truth.

The Python-level API (``FastExcel``, the metadata helpers) is annotated inline
in ``__init__.py`` and deliberately has no stub, so there is nothing to compare.
"""

import ast
import inspect
import pathlib

import pytest

from rustpy_xlsxwriter import rustpy_xlsxwriter as ext

# The extension filename carries an ABI tag, so build the stub path from the
# package directory rather than by swapping a suffix.
PKG_DIR = pathlib.Path(ext.__file__).parent
STUB_PATH = PKG_DIR / "rustpy_xlsxwriter.pyi"
STUB = ast.parse(STUB_PATH.read_text())


def _stub_params(node: ast.FunctionDef) -> list[str]:
    args = node.args
    return [a.arg for a in args.posonlyargs + args.args + args.kwonlyargs if a.arg != "self"]


def _stub_functions() -> dict[str, ast.FunctionDef]:
    return {n.name: n for n in STUB.body if isinstance(n, ast.FunctionDef)}


def _stub_class(name: str) -> ast.ClassDef:
    return next(n for n in STUB.body if isinstance(n, ast.ClassDef) and n.name == name)


def test_py_typed_marker_exists():
    """Without it PEP 561 tells type checkers to ignore the package entirely."""
    assert (PKG_DIR / "py.typed").is_file()


def test_stub_covers_exactly_the_extension_exports():
    exported = {n for n in dir(ext) if not n.startswith("_")}
    stubbed = set(_stub_functions()) | {
        n.name for n in STUB.body if isinstance(n, ast.ClassDef)
    }
    assert stubbed == exported


@pytest.mark.parametrize(
    "name", ["write_worksheet", "write_worksheets", "write_csv", "validate_sheet_name"]
)
def test_function_parameters_match(name):
    runtime = list(inspect.signature(getattr(ext, name)).parameters)
    assert _stub_params(_stub_functions()[name]) == runtime


def test_format_methods_match():
    runtime = {n for n in dir(ext.Format) if not n.startswith("_")}
    stubbed = {
        n.name
        for n in _stub_class("Format").body
        if isinstance(n, ast.FunctionDef) and not n.name.startswith("_")
    }
    assert stubbed == runtime


@pytest.mark.parametrize(
    "name", [n for n in dir(ext.Format) if not n.startswith("_")]
)
def test_format_setters_are_positional_only(name):
    """Arity and call style, not names.

    The macro that generates these setters calls every argument ``value``,
    and pyo3 makes it positional-only, so the stub's descriptive name is
    documentation the caller can never use as a keyword. What the stub must
    get right is the count, and that the argument is positional — otherwise
    a type checker blesses ``set_font_size(size=12)``, a runtime TypeError.

    Call style is asserted behaviourally below, not from ``inspect``:
    pyo3's ``__text_signature__`` reports the argument as ordinary, even
    though the call itself rejects keywords.
    """
    runtime = inspect.signature(getattr(ext.Format, name)).parameters
    node = next(
        n
        for n in _stub_class("Format").body
        if isinstance(n, ast.FunctionDef) and n.name == name
    )
    assert len(node.args.posonlyargs) + len(node.args.args) == len(runtime)
    assert not node.args.args[1:], f"{name}: argument must be marked positional-only"


def test_setter_keyword_call_is_rejected():
    """The behaviour the positional-only marks above exist to describe."""
    assert ext.Format().set_font_size(12) is not None
    with pytest.raises(TypeError, match="unexpected keyword argument"):
        ext.Format().set_font_size(size=12)
