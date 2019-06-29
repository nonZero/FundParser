import numpy as np

from fund_parser.consts import EMPTY, BAD_COLS, BAD_NAMES


def remove_empty_cols(a: np.ndarray):
    """Removes empty cols from an array.

    >>> print(remove_empty_cols(np.array([
    ...     ["a", "", "b"],
    ...     ["c", "", "d"],
    ...     ["", "", ""],
    ...     ["e", "", "f"],
    ... ])))
    [['a' 'b']
     ['c' 'd']
     ['' '']
     ['e' 'f']]
    """
    valid_cols = (a != "").sum(axis=0) > 0
    return a[:, valid_cols]


def clean_from_blacklist(a: np.ndarray, invalid_strings=EMPTY):
    """Removes empty cols from an array.

    >>> print(clean_from_blacklist(np.array(
    ...     ["a", "123$$456", "", "XXX stuff", "good", "XXX"]
    ... ), ["XXX", "$$"]))
    ['a' '' '' '' 'good' '']
    """
    result = a.copy()
    for s in invalid_strings:
        empty = np.char.count(result, s) > 0
        result = np.where(empty, "", result)
    return result


def raw_to_array(raw):
    data = np.array(raw, dtype=str)
    data = clean_from_blacklist(data)
    data = remove_empty_cols(data)
    return data


def get_best_dtype(x):
    if isinstance(x, str) and x.endswith(".0"):
        x = x[:-2]
    try:
        return int(x)
    except ValueError:
        try:
            return float(x)
        except ValueError:
            return x


def extract_rows(raw):
    data = raw_to_array(raw)

    # remove short lines
    value_counts = (data != "").sum(axis=1)
    good_rows = value_counts > value_counts.max() / 1.5
    data = data[good_rows]
    values = remove_empty_cols(data)  # yes, again.

    header = []
    untitled = 0
    for c in values[0]:
        v = c.strip().strip("*")
        if not v:
            untitled += 1
            v = f"untitled{untitled}"
        header.append(v)

    records = []

    for row in values[1:]:
        if row[row != ""][0] in BAD_COLS:
            continue
        vals = [get_best_dtype(v.strip()) for v in row]
        if vals[0] in BAD_NAMES:
            continue
        records.append(vals)

    return header, records
