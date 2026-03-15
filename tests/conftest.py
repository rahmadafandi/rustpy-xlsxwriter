import random
from typing import Any, Dict, List

import pytest
from faker import Faker

XLSX_MAGIC = b"PK\x03\x04"


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
    base = [_make_record(fake) for _ in range(min(20, count))]
    return (base * (count // len(base) + 1))[:count]


@pytest.fixture
def small_records():
    """10 records for quick unit tests."""
    return generate_records(10)


@pytest.fixture
def medium_records():
    """1000 records for moderate tests."""
    return generate_records(1000)
