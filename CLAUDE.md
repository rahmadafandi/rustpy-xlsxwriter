# CLAUDE.md

## Project Overview

RustPy-XlsxWriter is a high-performance Excel and CSV file generation library for Python, powered by Rust via PyO3. It achieves ~7-9x faster Excel and ~5x faster CSV than Python equivalents.

## Build & Development

```bash
pip install -e ".[dev]"   # tests + maturin + formatters

maturin develop           # development build
maturin develop --release # release build (with LTO)
maturin build --release   # production wheel
```

Dev dependencies live in `pyproject.toml` extras (`tests`, `dev`). There is no
`requirements.txt` — it was a `pip freeze` dump nothing consumed.

## Testing

```bash
# Unit tests only (fast, ~4 seconds)
pytest tests/ -m "not benchmark"

# All tests including benchmarks
pytest tests/

# Standalone benchmark script
python benchmark.py
```

Two opt-in markers: `benchmark` (deselected above) and `recalc`, which opens
output in LibreOffice and self-skips when it is not installed.

## Key Architecture

Source layout is in `src/` — `lib.rs` names every module the extension exports.

- **Output format**: `FastExcel(target, output_format=...)` wins; otherwise the
  target's extension decides (`.csv` → CSV, `.tsv` → TSV, anything else and
  every buffer → Excel)
- **Data input detection** (`data_types.rs`): `__arrow_c_stream__` → Arrow zero-copy, `get_column` → Polars fallback, `columns` → Pandas fallback, else → Records
- **Arrow path**: Manual Arrow C Data Interface via `arrow_ffi.rs` (no `pyo3-arrow` — avoids chrono-tz cross-compilation issues)
- **Records path**: First-row type caching — detect column types from row 1, skip type cascade for subsequent rows
- **CSV path**: Rust `Vec<u8>` buffer with `ryu`/`itoa` number formatting, proper CSV escaping
- **Constant memory mode**: All Excel paths write row-by-row for `rust_xlsxwriter` compatibility. It restricts *ordering* — row `n` closes every row below it — not which features are available
- **Format caching**: `Format` objects created once, reused across all cells
- **Targets**: `helpers.rs` extracts `PathBuf` (so `str` and any `os.PathLike` work), else falls back to a `.write()` method. Never coerce paths on the Python side — that layer is what hid the signatures from type checkers

## Coding Conventions

- Propagate every `rust_xlsxwriter` error with `.map_err(xlsx_err)?`. A bare
  `let _ =` is only for infallible writes into an in-memory buffer
- Check `PyBool` before `PyInt` (Python bool is subclass of int)
- Use `value.cast::<T>()` for Python native types, `value.extract::<T>()` for numpy scalar fallback
- Use `chars().count()` not `len()` for Unicode string length validation
- The Python-value type cascade lives ONCE in `cell.rs` (`classify_and_write` / `try_cached` over the `CellWriter` trait); Excel and CSV each implement a sink (`ExcelCell`, `CsvCell`). Change detection order there, not per-path.
- Tests must verify actual cell content via `openpyxl`, not just file existence
- CSV tests verify raw file content via string comparison

## Typing

The package ships `py.typed`, so everything below is what type checkers see.

- `rustpy_xlsxwriter.pyi` stubs **only** the compiled extension — what `lib.rs`
  exports, nothing else. `FastExcel` and the metadata helpers are annotated
  inline in `__init__.py` and deliberately have no stub
- Never wrap an extension function in a Python `(*args, **kwargs)` shim: it
  erases the signature the stub provides
- `Format` setters are positional-only (pyo3 names every macro-generated
  argument `value`), so the stub marks them `/`
- `test_type_stubs.py` compares the stub against the extension's own
  introspection — a new `#[pyfunction]` or setter fails it until stubbed

## Version Bumping

`Cargo.toml` → `version` is the only place. `pyproject.toml` declares the
version `dynamic`, and `get_version()` reads installed package metadata.
