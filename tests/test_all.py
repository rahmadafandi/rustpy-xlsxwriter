import os
import random
from concurrent.futures import ThreadPoolExecutor
from importlib.metadata import version
from typing import Any, Dict, List

import pytest
import xlsxwriter
from faker import Faker

from rustpy_xlsxwriter import (
    FastExcel,
    get_authors,
    get_description,
    get_homepage,
    get_license,
    get_name,
    get_repository,
    get_version,
    validate_sheet_name,
)

# ---------------------------------------------------------------------------
# Fixtures & helpers
# ---------------------------------------------------------------------------

TMP = "tmp"


@pytest.fixture(autouse=True, scope="module")
def tmp_dir():
    """Ensure tmp directory is clean before tests."""
    os.makedirs(TMP, exist_ok=True)
    for f in os.listdir(TMP):
        if f != ".gitignore":
            os.remove(os.path.join(TMP, f))
    yield
    # cleanup after all tests
    for f in os.listdir(TMP):
        if f != ".gitignore":
            os.remove(os.path.join(TMP, f))


def _make_record(fake: Faker) -> Dict[str, Any]:
    return {
        "name": fake.name(),
        "email": fake.email(),
        "address": fake.address() if random.random() > 0.2 else None,
        "phone": fake.phone_number() if random.random() > 0.2 else None,
        "date": fake.date() if random.random() > 0.2 else None,
        "numeric_int": random.randint(-1000, 1000),
        "numeric_float": round(random.uniform(-100.0, 100.0), 2),
        "text": fake.text(max_nb_chars=50) if random.random() > 0.2 else None,
        "boolean": random.choice([True, False, None]),
        "datetime": fake.date_time() if random.random() > 0.2 else None,
        "timestamp": fake.date_time() if random.random() > 0.2 else None,
        "time": fake.time() if random.random() > 0.2 else None,
        "dict": {"name": fake.name(), "email": fake.email()},
    }


def generate_records(count: int) -> List[Dict[str, Any]]:
    fake = Faker()
    fake.seed_instance(42)
    random.seed(42)

    if count <= 1000:
        base = [_make_record(fake) for _ in range(min(20, count))]
        return (base * (count // len(base) + 1))[:count]

    chunk_size = 10_000
    num_chunks = (count + chunk_size - 1) // chunk_size

    def _chunk(idx: int) -> List[Dict[str, Any]]:
        f = Faker()
        f.seed_instance(42 + idx)
        size = min(chunk_size, count - idx * chunk_size)
        base = [_make_record(f) for _ in range(20)]
        return (base * (size // len(base) + 1))[:size]

    with ThreadPoolExecutor() as pool:
        chunks = list(pool.map(_chunk, range(num_chunks)))
    return [r for c in chunks for r in c]


# ---------------------------------------------------------------------------
# Metadata tests
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
def test_get_version():
    assert get_version() == version("rustpy-xlsxwriter")


@pytest.mark.benchmark
def test_get_name():
    assert get_name() == "rustpy-xlsxwriter"


@pytest.mark.benchmark
def test_get_authors():
    assert get_authors() == "Rahmad Afandi <rahmadafandiii@gmail.com>"


@pytest.mark.benchmark
def test_get_description():
    assert get_description() == "Rust Python bindings for rust_xlsxwriter"


@pytest.mark.benchmark
def test_get_repository():
    assert get_repository() == "https://github.com/rahmadafandi/rustpy-xlsxwriter"


@pytest.mark.benchmark
def test_get_homepage():
    assert get_homepage() == "https://github.com/rahmadafandi/rustpy-xlsxwriter"


@pytest.mark.benchmark
def test_get_license():
    assert get_license() == "MIT"


@pytest.mark.benchmark
def test_validate_sheet_name():
    assert validate_sheet_name("Test") is True
    for char in ["[", "]", ":", "*", "?", "/", "\\"]:
        assert validate_sheet_name(f"Test{char}") is False


# ---------------------------------------------------------------------------
# FastExcel — single sheet
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000000])
def test_write_single_sheet(record_count: int):
    """Benchmark: single sheet via FastExcel."""
    records = generate_records(record_count)
    path = f"{TMP}/test_{record_count}.xlsx"
    FastExcel(path, password="password").sheet(f"Test {record_count}", records).save()
    assert os.path.exists(path)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [10])
def test_error_invalid_sheet_name_single(record_count: int):
    """Single sheet: ValueError on invalid sheet name."""
    records = generate_records(record_count)
    with pytest.raises(ValueError, match=r"Invalid sheet name"):
        FastExcel(f"{TMP}/should_not_exist.xlsx", password="password").sheet(
            "Test[]", records
        ).save()


# ---------------------------------------------------------------------------
# FastExcel — multiple sheets
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000])
def test_write_multiple_sheets(record_count: int):
    """Benchmark: multiple sheets via FastExcel."""
    records = generate_records(record_count)
    path = f"{TMP}/test_{record_count}_multi.xlsx"
    writer = FastExcel(path, password="password")
    for i in range(record_count):
        writer.sheet(f"Sheet{i}", records)
    writer.save()
    assert os.path.exists(path)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [10])
def test_error_invalid_sheet_name_multi(record_count: int):
    """Multiple sheets: ValueError on invalid sheet name."""
    with pytest.raises(ValueError, match=r"Invalid sheet name"):
        FastExcel(f"{TMP}/should_not_exist.xlsx", password="password").sheet(
            "Test[]", generate_records(record_count)
        ).save()


# ---------------------------------------------------------------------------
# Freeze panes — single sheet
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [10, 100, 1000])
def test_single_worksheet_freeze_panes(record_count: int):
    records = generate_records(record_count)

    # rows only
    path = f"{TMP}/freeze_rows_{record_count}.xlsx"
    FastExcel(path).freeze(row=2).sheet("Sheet1", records).save()
    assert os.path.exists(path)

    # cols only
    path = f"{TMP}/freeze_cols_{record_count}.xlsx"
    FastExcel(path).freeze(col=1).sheet("Sheet1", records).save()
    assert os.path.exists(path)

    # both
    path = f"{TMP}/freeze_both_{record_count}.xlsx"
    FastExcel(path).freeze(row=2, col=1).sheet("Sheet1", records).save()
    assert os.path.exists(path)

    # custom sheet name
    path = f"{TMP}/freeze_custom_{record_count}.xlsx"
    FastExcel(path).freeze(row=2, col=1).sheet("CustomSheet", records).save()
    assert os.path.exists(path)


# ---------------------------------------------------------------------------
# Freeze panes — multiple sheets
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [10, 100])
def test_multiple_worksheets_freeze_panes(record_count: int):
    records = generate_records(record_count)

    def _multi_writer(path, **freeze_kw):
        w = FastExcel(path)
        for k, v in freeze_kw.items():
            if k == "general":
                w.freeze(**v)
            else:
                w.freeze(sheet=k, **v)
        for i in range(record_count):
            w.sheet(f"Sheet{i}", records)
        w.save()
        assert os.path.exists(path)

    # general
    _multi_writer(
        f"{TMP}/multi_freeze_gen_{record_count}.xlsx", general={"row": 2, "col": 1}
    )

    # sheet-specific
    _multi_writer(
        f"{TMP}/multi_freeze_spec_{record_count}.xlsx",
        Sheet1={"row": 2, "col": 1},
        Sheet2={"row": 3, "col": 2},
        Sheet3={"row": 1, "col": 3},
    )

    # mixed
    _multi_writer(
        f"{TMP}/multi_freeze_mix_{record_count}.xlsx",
        general={"row": 2, "col": 1},
        Sheet2={"row": 3, "col": 2},
    )

    # row only
    _multi_writer(f"{TMP}/multi_freeze_row_{record_count}.xlsx", general={"row": 2})

    # col only
    _multi_writer(f"{TMP}/multi_freeze_col_{record_count}.xlsx", general={"col": 2})


# ---------------------------------------------------------------------------
# Freeze panes — edge cases
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
def test_freeze_panes_edge_cases():
    records = [{"col1": "v1", "col2": "v2"}, {"col1": "v3", "col2": "v4"}]

    # zero values
    FastExcel(f"{TMP}/edge_zero.xlsx").freeze(row=0, col=0).sheet(
        "Sheet1", records
    ).save()
    assert os.path.exists(f"{TMP}/edge_zero.xlsx")

    # no freeze (None is default)
    FastExcel(f"{TMP}/edge_none.xlsx").sheet("Sheet1", records).save()
    assert os.path.exists(f"{TMP}/edge_none.xlsx")

    # large values
    FastExcel(f"{TMP}/edge_large.xlsx").freeze(row=1000, col=100).sheet(
        "Sheet1", records
    ).save()
    assert os.path.exists(f"{TMP}/edge_large.xlsx")


# ---------------------------------------------------------------------------
# Baseline benchmark: xlsxwriter
# ---------------------------------------------------------------------------


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000000])
def test_xlsxwriter(record_count: int):
    """Baseline benchmark using native XlsxWriter library."""
    path = f"{TMP}/test_{record_count}_xlsxwriter.xlsx"
    wb = xlsxwriter.Workbook(path, {"constant_memory": True})
    ws = wb.add_worksheet()
    records = generate_records(record_count)

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
