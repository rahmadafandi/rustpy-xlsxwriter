"""
Benchmark tests — separated from unit tests.

Run with: pytest tests/test_benchmark.py -m benchmark
"""

import os
import random
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List

import pytest
from faker import Faker

from rustpy_xlsxwriter import FastExcel


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


@pytest.mark.benchmark
def test_rustpy_xlsxwriter_1m_rows(tmp_path):
    """Benchmark: 1M rows single sheet via rustpy-xlsxwriter."""
    records = _generate_large_records(1_000_000)
    path = str(tmp_path / "bench_1m.xlsx")
    FastExcel(path, password="password").sheet("Benchmark", records).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
def test_xlsxwriter_baseline_1m_rows(tmp_path):
    """Baseline benchmark: 1M rows via native Python XlsxWriter."""
    import xlsxwriter

    records = _generate_large_records(1_000_000)
    path = str(tmp_path / "bench_1m_baseline.xlsx")
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
    assert os.path.exists(path)
