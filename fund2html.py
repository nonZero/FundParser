import argparse
import html
import random
import urllib.request
from pathlib import Path

from fund_parser.extract import extract_fund_data

CSS_URL = "https://cdn.rtlcss.com/bootstrap/v4.2.1/css/bootstrap.css"


def _row(row):
    return "\t<tr>{}</tr>".format(
        "\n".join("\t\t<td>{}</td>".format(
            html.escape(str(cell))
        ) for cell in row)
    )


def table(data):
    return '<table class="table table-sm table-bordered table-hover">{}</table>'.format(
        "\n".join(_row(row) for row in data)
    )


def xlsx_to_html(source: Path, target: Path):
    main, sheets = extract_fund_data(str(source))
    nodata = []
    with target.open("w") as w:
        print('<link rel="stylesheet" href="bootstrap.css"/>', file=w)
        print('<div class="container-fluid" dir="rtl">', file=w)
        print(f'<h1 dir="ltr">{html.escape(str(source))}</h1>', file=w)

        print(table(main) if main is not None else "?", file=w)

        for name, df in sheets.items():
            if df is None:
                nodata.append(name)
                continue
            print(f"<h2>{html.escape(name)}</h2>", file=w)
            print(df.to_html(
                index=False,
                float_format="{:n}".format,
                classes="table table-striped table-bordered table-hover table-sm",
            ), file=w)

        if nodata:
            print("<hr/>", file=w)
            print("; ".join(f"<s>{x}</s>" for x in nodata), file=w)


def process_folder(sources, target, shuffle=False, preview=False, overwrite=True):
    sources_path = Path(sources)
    target_path = Path(target)
    target_path.mkdir(exist_ok=True)

    css = target_path / "bootstrap.css"
    if not css.exists():
        print("downloading css...")
        urllib.request.urlretrieve(CSS_URL, css)

    paths = list(sources_path.rglob("*.[xX][lL][sS][xX]"))

    if shuffle:
        random.shuffle(paths)

    for i, source in enumerate(paths, 1):
        name = source.name.replace(" ", "_")
        target = target_path / f"{name}.html"
        if not overwrite:
            if target.exists():
                continue

        print(i, len(paths), target)

        xlsx_to_html(source, target)

        if preview:
            import webbrowser
            webbrowser.open(str(target))
            break


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert fund xlsx files to HTML')
    parser.add_argument('sources', type=str,
                        help='path to folder containing xlsx files')
    parser.add_argument('target', type=str,
                        help='path to folder for generated html files')
    parser.add_argument('--shuffle', action='store_true',
                        help='randomize order of files before processing')
    parser.add_argument('--preview', action='store_true',
                        help='Process only the first file, open it in a browser window and quit')
    parser.add_argument('--skip-existing', dest='overwrite',
                        action='store_false',
                        help='Skip existing files')

    args = parser.parse_args()

    process_folder(**args.__dict__)
