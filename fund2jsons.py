import argparse
import json
import logging
import random
from pathlib import Path

from fund_parser.extract import extract_fund_data

logger = logging.getLogger(__name__)


def xlsx_to_json(source: Path, target: Path, info=None, keep_none=False):
    def extract(df):
        if df is None:
            return None
        d = df.to_dict(orient="split")
        del d['index']
        return d

    main, sheets = extract_fund_data(str(source))
    data = {name: extract(df) for name, df in sheets.items()
            if keep_none or df is not None}

    with target.open("w") as w:
        json.dump({
            'info': info,
            'sheets': data,
        }, w, ensure_ascii=False, indent=2)

    return data


def process_folder(sources, target, shuffle=False, preview=False,
                   overwrite=True):
    sources_path = Path(sources)
    target_path = Path(target)
    target_path.mkdir(exist_ok=True)

    paths = list(sources_path.rglob("*.[xX][lL][sS][xX]"))

    logger.info(f"Found {len(paths)} files in {sources!r}")

    if shuffle:
        random.shuffle(paths)

    for i, source in enumerate(paths, 1):
        name = source.relative_to(sources_path)
        target = target_path / name.with_suffix('.json')
        if not overwrite:
            if target.exists():
                continue
        if not target.parent.exists():
            target.parent.mkdir(parents=True)

        logger.info(f"[{i}/{len(paths)}] {str(target)!r}")

        xlsx_to_json(source, target, info={'key': target.stem})

        if preview:
            import webbrowser
            webbrowser.open(str(target))
            break


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert fund xlsx files to JSON')
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
