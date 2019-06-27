import pandas as pd

from fund_parser.consts import MAIN_SHEET_NAMES
from fund_parser.parsing import extract_rows
from fund_parser.xl import extract_sheets


def extract_fund_data(filename, always_return_dataframes=False,
                      main_sheet_names=MAIN_SHEET_NAMES):
    sheets = extract_sheets(filename)
    main_sheet = None
    data = {}
    for name, raw in sheets.items():
        if name in main_sheet_names:
            main_sheet = raw
            continue
        if not raw:
            continue
        header, v = extract_rows(raw)
        data[name] = (pd.DataFrame(v, columns=header)
                      if v or always_return_dataframes else None)

    return main_sheet, data
