from pathlib import Path

import pandas as pd
import pytest

from fund_parser.extract import extract_fund_data


@pytest.fixture
def data_path():
    return Path(__file__).parent / "test_data"


def test_extract(data_path):
    filename = data_path / "520004078_b259010_p418.xlsx"
    general, sheets = extract_fund_data(filename)
    assert isinstance(general, list)
    assert len(sheets) == 29
    for k, v in sheets.items():
        assert v is None or isinstance(v, pd.DataFrame)
