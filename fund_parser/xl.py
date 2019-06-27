import xlrd


def extract_cells(sh):
    return [sh.row_values(r) for r in range(sh.nrows)]


def extract_sheets(fn: str):
    b = xlrd.open_workbook(fn)
    it = zip(b.sheet_names(), b.sheets())
    return {n: extract_cells(sh) for n, sh in it}
