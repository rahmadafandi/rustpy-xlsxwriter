import os
from importlib.metadata import version
import pytest
import xlsxwriter
from typing import List, Dict, Any

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


@pytest.mark.benchmark
def generate_test_records(count: int) -> List[Dict[str, Any]]:
    """Generate test records using multiple threads.

    Args:
        count: Number of records to generate

    Returns:
        List of dictionaries containing test data
    """
    import math
    from concurrent.futures import ThreadPoolExecutor
    from faker import Faker
    import random

    fake = Faker()

    def generate_chunk(start: int, end: int) -> List[Dict[str, str]]:
        """Generate a chunk of test records.

        Args:
            start: Start index
            end: End index

        Returns:
            List of dictionaries with test data
        """
        return [
            {
                "name": fake.name(),
                "address": (
                    fake.address()
                    if random.random() > 0.3
                    else random.choice([None, ""])
                ),
                "something": (
                    ". ".join(fake.words(5))
                    if random.random() > 0.3
                    else random.choice([None, ""])
                ),
                "numeric_data": random.randint(0, 100),
            }
            for _ in range(start, end)
        ]

    num_threads = os.cpu_count() or 1
    chunk_size = math.ceil(count / num_threads)

    with ThreadPoolExecutor(max_workers=num_threads) as executor:
        futures = []
        for i in range(0, count, chunk_size):
            end = min(i + chunk_size, count)
            futures.append(executor.submit(generate_chunk, i, end))

        records = []
        for future in futures:
            records.extend(future.result())

    return records


@pytest.mark.benchmark
def generate_test_records_with_sheet_name(
    count: int, error_sheet_name: bool = False
) -> List[Dict[str, List[Dict[str, Any]]]]:
    """Generate test records with sheet names.

    Args:
        count: Number of records per sheet
        error_sheet_name: Whether to generate invalid sheet names

    Returns:
        List of dictionaries mapping sheet names to records
    """
    records = []
    for i in range(count):
        sheet_name = "Test[]" if error_sheet_name else f"Sheet{i}"
        records.append({sheet_name: generate_test_records(count)})
    return records


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [100] * 3)
def test_save_records_multiple_sheets(record_count: int) -> None:
    """Test saving records to multiple sheets.

    Args:
        record_count: Number of records per sheet
    """
    records = generate_test_records_with_sheet_name(record_count)
    filename = f"tmp/test_{record_count}_multiple_sheets.xlsx"
    assert (
        rustpy_xlsxwriter.save_records_multiple_sheets(records, filename, "password")
        is None
    )
    assert os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [100000] * 3)
def test_save_error_name_sheet_records_single_sheet(record_count: int) -> None:
    """Test error handling for invalid sheet names in single sheet mode.

    Args:
        record_count: Number of records to generate
    """
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}_single_sheet_error.xlsx"
    sheet_name = "Test[]"
    with pytest.raises(ValueError) as e:
        rustpy_xlsxwriter.save_records(records, filename, sheet_name, "password")
    assert (
        str(e.value)
        == "Invalid sheet name 'Test[]'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\"
    )
    assert not os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [100] * 3)
def test_save_error_name_sheet_records_multiple_sheets(record_count: int) -> None:
    """Test error handling for invalid sheet names in multiple sheet mode.

    Args:
        record_count: Number of records per sheet
    """
    records = generate_test_records_with_sheet_name(record_count, True)
    filename = f"tmp/test_{record_count}_multiple_sheets_error.xlsx"
    with pytest.raises(ValueError) as e:
        rustpy_xlsxwriter.save_records_multiple_sheets(records, filename, "password")
    assert (
        str(e.value)
        == "Invalid sheet name 'Test[]'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\"
    )
    assert not os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [100000] * 3)
def test_save_records_single_sheet(record_count: int) -> None:
    """Test saving records to a single sheet.

    Args:
        record_count: Number of records to generate
    """
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}.xlsx"
    sheet_name = f"Test {record_count}"
    assert (
        rustpy_xlsxwriter.save_records(records, filename, sheet_name, "password")
        is None
    )
    assert os.path.exists(filename)


@pytest.mark.benchmark
@pytest.mark.parametrize("record_count", [100000] * 3)
def test_xlsxwriter(record_count: int) -> None:
    """Benchmark test using native XlsxWriter library.

    Args:
        record_count: Number of records to generate
    """
    filename = f"tmp/test_{record_count}_xlsxwriter.xlsx"
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    records = generate_test_records(record_count)

    # Write headers
    headers = list(records[0].keys())
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write data
    for row, record in enumerate(records, start=1):
        for col, header in enumerate(headers):
            worksheet.write(row, col, record[header])

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
def test_get_repositorvy() -> None:
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
