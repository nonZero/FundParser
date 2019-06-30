import argparse
import logging
import random
import urllib.request
from functools import lru_cache
from pathlib import Path

from jinja2 import Template

from fund_parser.extract import extract_fund_data

logger = logging.getLogger(__name__)

CSS_URL = "https://cdn.rtlcss.com/bootstrap/v4.2.1/css/bootstrap.css"


@lru_cache()
def load_template():
    with open("template.html") as f:
        return Template(f.read())


def xlsx_to_html(source: Path, target: Path):
    main, sheets = extract_fund_data(str(source))
    template = load_template()
    nodata = [name for name, df in sheets.items() if df is None]
    tables = [(name, df.to_html(
        index=False,
        float_format="{:n}".format,
        classes="table table-striped table-bordered table-hover table-sm",
    )) for name, df in sheets.items() if df is not None]

    with target.open("w") as w:
        w.write(template.render(
            filename=str(source),
            main=main,
            sheets=tables,
            nodata=nodata,
        ))


def process_folder(sources, target, shuffle=False, preview=False,
                   overwrite=True):
    sources_path = Path(sources)
    target_path = Path(target)
    target_path.mkdir(exist_ok=True)

    css = target_path / "bootstrap.css"
    if not css.exists():
        logger.warning("downloading css...")
        urllib.request.urlretrieve(CSS_URL, css)

    paths = list(sources_path.rglob("*.[xX][lL][sS][xX]"))

    logger.info(f"Found {len(paths)} files in {sources!r}")

    if shuffle:
        random.shuffle(paths)

    for i, source in enumerate(paths, 1):
        name = source.name.replace(" ", "_")
        target = target_path / f"{name}.html"
        if not overwrite:
            if target.exists():
                continue

        logger.info(f"[{i}/{len(paths)}] {str(target)!r}")

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

    logging.basicConfig(
        format='[%(levelname).1s %(asctime).19s %(module)s:%(lineno)d] %(message)s',
        level=logging.INFO,
    )

    process_folder(**args.__dict__)
