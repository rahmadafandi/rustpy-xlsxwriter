# Contributing to RustPy-XlsxWriter

Thanks for your interest in contributing! Here's how to get started.

## Development Setup

### Prerequisites

- Python 3.8+
- Rust (latest stable)
- [maturin](https://github.com/PyO3/maturin)

### Setup

```bash
git clone https://github.com/rahmadafandi/rustpy-xlsxwriter.git
cd rustpy-xlsxwriter
python -m venv .venv
source .venv/bin/activate
pip install maturin
pip install -e ".[tests]"
```

### Build

```bash
# Development (fast compile, no optimizations)
maturin develop

# Release (optimized, LTO — slower compile)
maturin develop --release
```

### Run Tests

```bash
# Unit tests only (~1 second)
pytest tests/ -m "not benchmark"

# All tests including benchmarks
pytest tests/

# Benchmark
python benchmark.py
```

## Project Structure

```
src/                          # Rust source
├── lib.rs                   # PyO3 module entry point
├── worksheet.rs             # Core write logic
├── arrow_ffi.rs             # Arrow C Data Interface bridge
├── arrow_writer.rs          # Arrow RecordBatch writer
├── data_types.rs            # Input type detection
├── metadata.rs              # Package metadata
└── utils.rs                 # Validation utilities

rustpy_xlsxwriter/           # Python package
├── __init__.py              # FastExcel class
└── rustpy_xlsxwriter.pyi    # Type stubs

tests/                       # Test suite (86 tests)
examples/                    # 13 runnable examples
```

## How to Contribute

### Reporting Bugs

- Open an [issue](https://github.com/rahmadafandi/rustpy-xlsxwriter/issues/new?template=bug_report.md)
- Include Python version, OS, and a minimal reproduction

### Suggesting Features

- Open an [issue](https://github.com/rahmadafandi/rustpy-xlsxwriter/issues/new?template=feature_request.md)
- Describe the use case and expected behavior

### Pull Requests

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Make your changes
4. Add tests for new functionality
5. Run `pytest tests/ -m "not benchmark"` and ensure all tests pass
6. Run `cargo check` to ensure no Rust warnings
7. Commit with a descriptive message
8. Push and open a Pull Request

### Coding Guidelines

**Rust:**
- Propagate errors with `.map_err(xlsx_err)?` — never use `let _ =`
- Check `PyBool` before `PyInt` (Python bool is subclass of int)
- Keep `write_py_any_bound` and `write_py_any_bound_detect` in sync
- Use `chars().count()` for Unicode string length

**Python:**
- Follow PEP 8
- Update `.pyi` type stubs when changing the Python API

**Tests:**
- Verify actual cell content via `openpyxl`, not just file existence
- Use `tmp_path` fixture for temporary files

### Version Bumping

Update in both files:
1. `Cargo.toml` — `version = "x.y.z"`
2. `rustpy_xlsxwriter/rustpy_xlsxwriter.pyi` — `get_version()` docstring
