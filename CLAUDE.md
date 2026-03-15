# CLAUDE.md

## Project Overview

RustPy-XlsxWriter is a high-performance Excel file generation library for Python, powered by Rust via PyO3. It achieves ~7-9x faster performance than Python's xlsxwriter.

## Build & Development

```bash
# Development build
maturin develop

# Release build (with LTO)
maturin develop --release

# Production wheel
maturin build --release
```

## Testing

```bash
# Unit tests only (fast, ~1 second)
pytest tests/ -m "not benchmark"

# All tests including benchmarks (~8 minutes)
pytest tests/

# Run benchmarks with codspeed
pytest tests/test_benchmark.py --codspeed

# Standalone benchmark script
python benchmark.py
```

## Project Structure

```
src/
├── lib.rs              # PyO3 module entry point
├── worksheet.rs        # Core write logic (Records, Pandas, Polars, Arrow)
├── arrow_writer.rs     # Arrow zero-copy batch writer
├── data_types.rs       # WorksheetData enum (ArrowStream, Records, Pandas, Polars)
├── metadata.rs         # Package metadata functions
└── utils.rs            # Sheet name validation

rustpy_xlsxwriter/
├── __init__.py         # FastExcel builder class + Python API
└── rustpy_xlsxwriter.pyi  # Type stubs for IDE support

tests/
├── conftest.py            # Shared fixtures
├── test_metadata.py       # Package metadata
├── test_validation.py     # Sheet name validation
├── test_write_single.py   # Single sheet writing
├── test_write_multi.py    # Multiple sheets
├── test_write_functional.py # Functional API
├── test_freeze_panes.py   # Freeze panes
├── test_password.py       # Password protection
├── test_bytesio.py        # In-memory buffer
├── test_dataframe.py      # Pandas DataFrame
├── test_polars.py         # Polars DataFrame
├── test_styling.py        # Float/datetime format, bold headers
└── test_benchmark.py      # Performance benchmarks
```

## Key Architecture

- **Data input detection** (`data_types.rs`): `__arrow_c_stream__` → Arrow zero-copy, `get_column` → Polars fallback, `columns` → Pandas fallback, else → Records (list of dicts / generator)
- **Arrow path**: Uses `pyo3-arrow` + `arrow-array` to read DataFrame memory directly in Rust — no Python object conversion
- **Records path**: First-row type caching — detect column types from row 1, use fast single-cast dispatch for subsequent rows
- **Constant memory mode**: All paths write row-by-row for `rust_xlsxwriter` constant memory compatibility
- **Format caching**: `Format` objects created once, reused across all cells

## Dependencies

- `pyo3` 0.28 — Rust-Python bindings
- `rust_xlsxwriter` 0.93 — Excel file generation (constant_memory, ryu, zlib)
- `pyo3-arrow` 0.17 + `arrow-array` 58 — Arrow zero-copy for DataFrames
- `indexmap` — Ordered maps for deterministic sheet ordering

## Coding Conventions

- Propagate all write errors with `.map_err(xlsx_err)?` — never use `let _ =`
- Check `PyBool` before `PyInt` (Python bool is subclass of int)
- Use `value.cast::<T>()` for Python native types, `value.extract::<T>()` for numpy scalar fallback
- Use `chars().count()` not `len()` for Unicode string length validation
- Keep `write_py_any_bound` and `write_py_any_bound_detect` in sync — they share the same type cascade logic
- Tests must verify actual cell content via `openpyxl`, not just file existence

## Version Bumping

Update version in both:
1. `Cargo.toml` → `version = "x.y.z"`
2. `rustpy_xlsxwriter/rustpy_xlsxwriter.pyi` → `get_version()` docstring example
