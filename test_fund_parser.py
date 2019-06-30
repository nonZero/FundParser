import csv
from collections import Counter
from pathlib import Path

import numpy as np

from fund_parser.extract import extract_fund_data


def collect_tests():
    data_path = Path(__file__).parent / "test_data"
    with (data_path / "expected.csv").open() as f:
        reader = csv.DictReader(f)
        for d in reader:
            yield (
                str(data_path / d['filename']),
                int(d['sheet_count']),
                int(d['none_count']),
            )


def pytest_generate_tests(metafunc):
    tests = list(collect_tests())

    metafunc.parametrize(
        "xls_filename,expected_sheets,expected_empty_sheets", tests)


def test_xls(xls_filename, expected_sheets, expected_empty_sheets):
    general, sheets = extract_fund_data(xls_filename)
    assert isinstance(general, np.ndarray)
    c = Counter(type(v).__name__ for v in sheets.values())
    assert c['DataFrame'] == expected_sheets
    assert c['NoneType'] == expected_empty_sheets
