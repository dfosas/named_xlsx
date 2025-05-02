# coding=utf-8
import argparse
import shutil
from pathlib import Path
from multiprocessing import Pool
from tempfile import TemporaryDirectory

import xlwings


def refresh_path(path: Path):
    modified_time_old = path.stat().st_mtime
    wb = xlwings.Book(path)
    wb.save()
    wb.close()
    modified_time_new = path.stat().st_mtime
    if modified_time_old >= modified_time_new:
        raise ValueError("File was not updated.")


def refresh_paths_in_tempdir(paths: list[Path], folder=None, parallel: bool = False):
    with xlwings.App() as _, TemporaryDirectory(dir=folder) as tempdir:
        paths_tmp = [Path(tempdir) / path.name for path in paths]
        for path, path_tmp in zip(paths, paths_tmp):
            shutil.copy(path, path_tmp)
            assert path_tmp.exists()
        refresh_paths(paths_tmp, parallel=parallel)

        for path, path_tmp in zip(paths, paths_tmp):
            shutil.copy(path_tmp, path)


def refresh_paths(paths: list[Path], parallel: bool = False):
    with xlwings.App() as _:
        if parallel:
            with Pool() as p:
                p.map(refresh_path, paths)
        else:
            for path in paths:
                refresh_path(path)


def refresher():
    parser = argparse.ArgumentParser(description="Refresh cached values of xlsx.")
    parser.add_argument("--root", type=Path, default=Path(__file__).parent)
    parser.add_argument("--inplace", action="store_true")
    parser.add_argument("--parallel", action="store_true")
    args = parser.parse_args()
    print(args)

    _paths = list(sorted(args.root.glob("*.xlsx")))
    with xlwings.App() as _:
        if args.inplace:
            # For cases where it is fine to use the original paths (avoids copies)
            refresh_paths(_paths, parallel=args.parallel)
        else:
            # Using a local temporal folder helps with network path issues in Excel
            refresh_paths_in_tempdir(_paths, parallel=args.parallel)
