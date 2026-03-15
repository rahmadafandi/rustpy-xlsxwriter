"""
Benchmark tests — separated from unit tests.

Run with: pytest tests/test_benchmark.py -m benchmark
"""

import os
import random
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List

import numpy as np
import pandas as pd
import polars as pl
import pytest
from faker import Faker

from rustpy_xlsxwriter import FastExcel


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _generate_large_records(count: int) -> List[Dict[str, Any]]:
    """Generate large record set with parallel chunking."""
    fake = Faker()
    fake.seed_instance(42)
    random.seed(42)

    chunk_size = 10_000
    num_chunks = (count + chunk_size - 1) // chunk_size

    def _make_record(f: Faker) -> Dict[str, Any]:
        return {
            "name": f.name(),
            "email": f.email(),
            "address": f.address() if random.random() > 0.2 else None,
            "phone": f.phone_number() if random.random() > 0.2 else None,
            "date": f.date() if random.random() > 0.2 else None,
            "numeric_int": random.randint(-1000, 1000),
            "numeric_float": round(random.uniform(-100.0, 100.0), 2),
            "text": f.text(max_nb_chars=50) if random.random() > 0.2 else None,
            "boolean": random.choice([True, False, None]),
            "datetime": f.date_time() if random.random() > 0.2 else None,
            "timestamp": f.date_time() if random.random() > 0.2 else None,
            "time": f.time() if random.random() > 0.2 else None,
            "dict": {"name": f.name(), "email": f.email()},
        }

    def _chunk(idx: int) -> List[Dict[str, Any]]:
        f = Faker()
        f.seed_instance(42 + idx)
        size = min(chunk_size, count - idx * chunk_size)
        base = [_make_record(f) for _ in range(20)]
        return (base * (size // len(base) + 1))[:size]

    with ThreadPoolExecutor() as pool:
        chunks = list(pool.map(_chunk, range(num_chunks)))
    return [r for c in chunks for r in c]


def _generate_large_dataframe(count: int) -> pd.DataFrame:
    """Generate large DataFrame with mixed typed columns."""
    np.random.seed(42)
    return pd.DataFrame(
        {
            "int_col": np.random.randint(0, 1000, count),
            "float_col": np.random.uniform(0, 100, count),
            "str_col": [f"row_{i}" for i in range(count)],
            "bool_col": np.random.choice([True, False], count),
        }
    )


def _xlsxwriter_write_records(records: List[Dict[str, Any]], path: str) -> None:
    """Baseline: write records using native Python XlsxWriter."""
    import xlsxwriter

    wb = xlsxwriter.Workbook(path, {"constant_memory": True})
    ws = wb.add_worksheet()

    headers = list(records[0].keys())
    for col, h in enumerate(headers):
        ws.write(0, col, h)

    for i, rec in enumerate(records, start=1):
        for col, h in enumerate(headers):
            val = rec[h]
            if isinstance(val, dict):
                ws.write_string(i, col, str(val))
            else:
                ws.write(i, col, val)

    wb.close()


def _xlsxwriter_write_dataframe(df: pd.DataFrame, path: str) -> None:
    """Baseline: write DataFrame using native Python XlsxWriter."""
    import xlsxwriter

    wb = xlsxwriter.Workbook(path, {"constant_memory": True})
    ws = wb.add_worksheet()

    headers = list(df.columns)
    for col, h in enumerate(headers):
        ws.write(0, col, h)

    for i, row in enumerate(df.itertuples(index=False, name=None), start=1):
        for col, val in enumerate(row):
            if isinstance(val, bool):
                ws.write_boolean(i, col, val)
            elif isinstance(val, (int, float, np.integer, np.floating)):
                ws.write_number(i, col, float(val))
            else:
                ws.write_string(i, col, str(val))

    wb.close()


# ---------------------------------------------------------------------------
# Records benchmarks
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
def test_records_500k_rustpy(tmp_path):
    """Benchmark: 500K rows records via rustpy-xlsxwriter."""
    records = _generate_large_records(500_000)
    path = str(tmp_path / "records_500k.xlsx")
    FastExcel(path, password="password").sheet("Benchmark", records).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_records_500k_xlsxwriter(tmp_path):
    """Baseline: 500K rows records via Python XlsxWriter."""
    records = _generate_large_records(500_000)
    path = str(tmp_path / "records_500k_baseline.xlsx")
    _xlsxwriter_write_records(records, path)
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_records_1m_rustpy(tmp_path):
    """Benchmark: 1M rows records via rustpy-xlsxwriter."""
    records = _generate_large_records(1_000_000)
    path = str(tmp_path / "records_1m.xlsx")
    FastExcel(path, password="password").sheet("Benchmark", records).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_records_1m_xlsxwriter(tmp_path):
    """Baseline: 1M rows records via Python XlsxWriter."""
    records = _generate_large_records(1_000_000)
    path = str(tmp_path / "records_1m_baseline.xlsx")
    _xlsxwriter_write_records(records, path)
    assert os.path.exists(path)


# ---------------------------------------------------------------------------
# DataFrame benchmarks
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
def test_dataframe_500k_rustpy(tmp_path):
    """Benchmark: 500K rows DataFrame via rustpy-xlsxwriter."""
    df = _generate_large_dataframe(500_000)
    path = str(tmp_path / "df_500k.xlsx")
    FastExcel(path, autofit=False).sheet("Benchmark", df).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_dataframe_500k_xlsxwriter(tmp_path):
    """Baseline: 500K rows DataFrame via Python XlsxWriter."""
    df = _generate_large_dataframe(500_000)
    path = str(tmp_path / "df_500k_baseline.xlsx")
    _xlsxwriter_write_dataframe(df, path)
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_dataframe_1m_rustpy(tmp_path):
    """Benchmark: 1M rows DataFrame via rustpy-xlsxwriter."""
    df = _generate_large_dataframe(1_000_000)
    path = str(tmp_path / "df_1m.xlsx")
    FastExcel(path, autofit=False).sheet("Benchmark", df).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_dataframe_1m_xlsxwriter(tmp_path):
    """Baseline: 1M rows DataFrame via Python XlsxWriter."""
    df = _generate_large_dataframe(1_000_000)
    path = str(tmp_path / "df_1m_baseline.xlsx")
    _xlsxwriter_write_dataframe(df, path)
    assert os.path.exists(path)


# ---------------------------------------------------------------------------
# Polars DataFrame benchmarks
# ---------------------------------------------------------------------------


def _generate_large_polars_dataframe(count: int) -> pl.DataFrame:
    """Generate large Polars DataFrame with mixed typed columns."""
    np.random.seed(42)
    return pl.DataFrame(
        {
            "int_col": np.random.randint(0, 1000, count),
            "float_col": np.random.uniform(0, 100, count),
            "str_col": [f"row_{i}" for i in range(count)],
            "bool_col": np.random.choice([True, False], count),
        }
    )


def _xlsxwriter_write_polars(df_pl: pl.DataFrame, path: str) -> None:
    """Baseline: write Polars DataFrame using native Python XlsxWriter."""
    import xlsxwriter

    wb = xlsxwriter.Workbook(path, {"constant_memory": True})
    ws = wb.add_worksheet()

    headers = df_pl.columns
    for col, h in enumerate(headers):
        ws.write(0, col, h)

    for i, row in enumerate(df_pl.iter_rows(), start=1):
        for col, val in enumerate(row):
            if val is None:
                pass
            elif isinstance(val, bool):
                ws.write_boolean(i, col, val)
            elif isinstance(val, (int, float, np.integer, np.floating)):
                ws.write_number(i, col, float(val))
            else:
                ws.write_string(i, col, str(val))

    wb.close()


@pytest.mark.benchmark
def test_polars_500k_rustpy(tmp_path):
    """Benchmark: 500K rows Polars DataFrame via rustpy-xlsxwriter."""
    df = _generate_large_polars_dataframe(500_000)
    path = str(tmp_path / "polars_500k.xlsx")
    FastExcel(path, autofit=False).sheet("Benchmark", df).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_polars_500k_xlsxwriter(tmp_path):
    """Baseline: 500K rows Polars DataFrame via Python XlsxWriter."""
    df = _generate_large_polars_dataframe(500_000)
    path = str(tmp_path / "polars_500k_baseline.xlsx")
    _xlsxwriter_write_polars(df, path)
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_polars_1m_rustpy(tmp_path):
    """Benchmark: 1M rows Polars DataFrame via rustpy-xlsxwriter."""
    df = _generate_large_polars_dataframe(1_000_000)
    path = str(tmp_path / "polars_1m.xlsx")
    FastExcel(path, autofit=False).sheet("Benchmark", df).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_polars_1m_xlsxwriter(tmp_path):
    """Baseline: 1M rows Polars DataFrame via Python XlsxWriter."""
    df = _generate_large_polars_dataframe(1_000_000)
    path = str(tmp_path / "polars_1m_baseline.xlsx")
    _xlsxwriter_write_polars(df, path)
    assert os.path.exists(path)
