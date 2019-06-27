import html
import random
import urllib.request
from pathlib import Path

from fund_parser.extract import extract_fund_data
from fund_parser.parsing import raw_to_array

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

        print(table(raw_to_array(main)) if main else "?", file=w)

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


def main(sources, target, shuffle=False, just_show_one=False):
    sources_path = Path(sources)
    target_path = Path(target)
    target_path.mkdir(exist_ok=True)

    css = target_path / "bootstrap.css"
    if not css.exists():
        print("downloading css...")
        urllib.request.urlretrieve(CSS_URL, css)

    paths = [p for p in sources_path.glob("**/*") if
             p.suffix.lower() == '.xlsx']

    if shuffle:
        random.shuffle(paths)

    for i, source in enumerate(paths, 1):
        name = source.name.replace(" ", "_")
        target = target_path / f"{name}.html"
        if target.exists():
            continue

        print(i, len(paths), target)

        xlsx_to_html(source, target)

        if just_show_one:
            import webbrowser
            webbrowser.open(str(target))
            break


if __name__ == '__main__':
    main("./sources/", "./out/", True, True)
