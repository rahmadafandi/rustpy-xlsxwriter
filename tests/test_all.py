import os
import random
from concurrent.futures import ThreadPoolExecutor
from importlib.metadata import version
from typing import Any, Dict, List

import pytest
import xlsxwriter
from faker import Faker

import rustpy_xlsxwriter


def setup_module():
    """Create tmp directory if it doesn't exist"""
    if not os.path.exists("tmp"):
        os.makedirs("tmp", exist_ok=True)
    else:
        # clean up tmp directory except for .gitignore
        for file in os.listdir("tmp"):
            if file not in [".gitignore"]:
                os.remove(os.path.join("tmp", file))


@pytest.mark.benchmark
def test_get_version() -> None:
    """Test that version matches package metadata"""
    assert rustpy_xlsxwriter.get_version() == version("rustpy-xlsxwriter")


def _generate_base_record(fake: Faker) -> Dict[str, Any]:
    """Generate a single base record"""
    return {
        "name": fake.name(),
        "email": fake.email(),
        "address": (
            fake.address() if random.random() > 0.2 else random.choice([None, ""])
        ),
        "phone": (
            fake.phone_number() if random.random() > 0.2 else random.choice([None, ""])
        ),
        "date": fake.date() if random.random() > 0.2 else None,
        "numeric_int": random.randint(-1000, 1000),
        "numeric_float": round(random.uniform(-100.0, 100.0), 2),
        "text": (
            fake.text(max_nb_chars=50)
            if random.random() > 0.2
            else random.choice([None, ""])
        ),
        "boolean": random.choice([True, False, None]),
        "datetime": fake.date_time() if random.random() > 0.2 else None,
        "timestamp": fake.date_time() if random.random() > 0.2 else None,
        "time": fake.time() if random.random() > 0.2 else None,
        "dict": {"name": fake.name(), "email": fake.email()},
    }


@pytest.mark.benchmark
def generate_test_records(count: int) -> List[Dict[str, Any]]:
    """Generate test records using parallel processing for large datasets"""
    fake = Faker()
    fake.seed_instance(42)
    random.seed(42)

    # For small counts, use direct generation
    if count <= 1000:
        base_records = [_generate_base_record(fake) for _ in range(min(20, count))]
        multiplier = count // len(base_records) + 1
        return (base_records * multiplier)[:count]

    # For large counts, use parallel processing
    chunk_size = 10000
    num_chunks = (count + chunk_size - 1) // chunk_size

    def generate_chunk(chunk_idx: int) -> List[Dict[str, Any]]:
        chunk_fake = Faker()
        chunk_fake.seed_instance(42 + chunk_idx)
        size = min(chunk_size, count - chunk_idx * chunk_size)
        base_records = [_generate_base_record(chunk_fake) for _ in range(20)]
        multiplier = size // len(base_records) + 1
        return (base_records * multiplier)[:size]

    with ThreadPoolExecutor() as executor:
        chunks = list(executor.map(generate_chunk, range(num_chunks)))

    return [record for chunk in chunks for record in chunk]


@pytest.mark.benchmark
def generate_test_records_with_sheet_name(
    count: int, error_sheet_name: bool = False
) -> List[Dict[str, List[Dict[str, Any]]]]:
    """Generate test records with sheet names"""
    shared_records = generate_test_records(count)
    records = []

    # Pre-generate all sheet names
    sheet_names = ["Test[]" if error_sheet_name else f"Sheet{i}" for i in range(count)]

    # Create records list with references to shared_records
    for sheet_name in sheet_names:
        records.append({sheet_name: shared_records})

    return records


# Rest of the test functions remain unchanged
@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000])
def test_write_worksheets(record_count: int) -> None:
    """Test saving records to multiple sheets."""
    try:
        records = generate_test_records_with_sheet_name(record_count)
        filename = f"tmp/test_{record_count}_multiple_sheets.xlsx"
        assert rustpy_xlsxwriter.write_worksheets(records, filename, "password") is None
        assert os.path.exists(filename)
    except Exception as e:
        print(e)
        raise e


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000000])
def test_save_error_name_sheet_records_single_sheet(record_count: int) -> None:
    """Test error handling for invalid sheet names in single sheet mode."""
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}_single_sheet_error.xlsx"
    sheet_name = "Test[]"
    with pytest.raises(ValueError) as e:
        rustpy_xlsxwriter.write_worksheet(records, filename, sheet_name, "password")
    assert (
        str(e.value)
        == "Invalid sheet name 'Test[]'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\"
    )
    assert not os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000])
def test_save_error_name_sheet_records_multiple_sheets(record_count: int) -> None:
    """Test error handling for invalid sheet names in multiple sheet mode."""
    records = generate_test_records_with_sheet_name(record_count, True)
    filename = f"tmp/test_{record_count}_multiple_sheets_error.xlsx"
    with pytest.raises(ValueError) as e:
        rustpy_xlsxwriter.write_worksheets(records, filename, "password")
    assert (
        str(e.value)
        == "Invalid sheet name 'Test[]'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\"
    )
    assert not os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000000])
def test_write_worksheet_single_sheet(record_count: int) -> None:
    """Test saving records to a single sheet."""
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}.xlsx"
    sheet_name = f"Test {record_count}"
    assert (
        rustpy_xlsxwriter.write_worksheet(records, filename, sheet_name, "password")
        is None
    )
    assert os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [1000000])
def test_xlsxwriter(record_count: int) -> None:
    """Benchmark test using native XlsxWriter library."""
    filename = f"tmp/test_{record_count}_xlsxwriter.xlsx"
    workbook = xlsxwriter.Workbook(filename, {"constant_memory": True})
    worksheet = workbook.add_worksheet()
    records = generate_test_records(record_count)

    # Write headers
    headers = list(records[0].keys())
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write data in chunks for better memory management
    chunk_size = 10000
    for chunk_start in range(0, len(records), chunk_size):
        chunk_end = min(chunk_start + chunk_size, len(records))
        chunk = records[chunk_start:chunk_end]
        for i, record in enumerate(chunk, start=chunk_start + 1):
            for col, header in enumerate(headers):
                if isinstance(record[header], (dict)):
                    worksheet.write_string(i, col, str(record[header]))
                else:
                    worksheet.write(i, col, record[header])

    workbook.close()
    assert os.path.exists(filename)


@pytest.mark.benchmark
def test_get_name() -> None:
    """Test get_name returns correct package name"""
    assert rustpy_xlsxwriter.get_name() == "rustpy-xlsxwriter"


@pytest.mark.benchmark
def test_get_authors() -> None:
    """Test get_authors returns correct author info"""
    assert rustpy_xlsxwriter.get_authors() == "Rahmad Afandi <rahmadafandiii@gmail.com>"


@pytest.mark.benchmark
def test_get_description() -> None:
    """Test get_description returns correct package description"""
    assert (
        rustpy_xlsxwriter.get_description()
        == "Rust Python bindings for rust_xlsxwriter"
    )


@pytest.mark.benchmark
def test_get_repository() -> None:
    """Test get_repository returns correct repository URL"""
    assert (
        rustpy_xlsxwriter.get_repository()
        == "https://github.com/rahmadafandi/rustpy-xlsxwriter"
    )


@pytest.mark.benchmark
def test_get_homepage() -> None:
    """Test get_homepage returns correct homepage URL"""
    assert (
        rustpy_xlsxwriter.get_homepage()
        == "https://github.com/rahmadafandi/rustpy-xlsxwriter"
    )


@pytest.mark.benchmark
def test_get_license() -> None:
    """Test get_license returns correct license"""
    assert rustpy_xlsxwriter.get_license() == "MIT"


@pytest.mark.benchmark
def test_validate_sheet_name() -> None:
    """Test sheet name validation logic"""
    assert rustpy_xlsxwriter.validate_sheet_name("Test") is True

    invalid_chars = ["[", "]", ":", "*", "?", "/", "\\"]
    for char in invalid_chars:
        assert rustpy_xlsxwriter.validate_sheet_name(f"Test{char}") is False
