# coding=utf-8
"""Routines to refresh calculated values."""

import shutil
from pathlib import Path
from multiprocessing import Pool
from tempfile import TemporaryDirectory

DEFAULT_REFRESH_ROOT = Path(__file__).parent


def _require_xlwings():
    try:
        import xlwings
    except ImportError as exc:
        raise RuntimeError(
            "The refresh command requires the optional 'xlsx' dependencies. "
            "Install them with `named_xlsx[xlsx]`."
        ) from exc
    return xlwings


def refresh_path(path: Path) -> None:
    """

    Parameters
    ----------
    path

    """
    xlwings = _require_xlwings()
    modified_time_old = path.stat().st_mtime
    with xlwings.Book(path) as wb:
        wb.save()
    modified_time_new = path.stat().st_mtime
    if modified_time_old >= modified_time_new:
        raise ValueError("File was not updated.")


def refresh_paths_in_tempdir(paths: list[Path], folder=None, parallel: bool = False):
    xlwings = _require_xlwings()
    with xlwings.App() as _, TemporaryDirectory(dir=folder) as tempdir:
        paths_tmp = [Path(tempdir) / path.name for path in paths]
        for path, path_tmp in zip(paths, paths_tmp):
            shutil.copy(path, path_tmp)
            assert path_tmp.exists()
        refresh_paths(paths_tmp, parallel=parallel)

        for path, path_tmp in zip(paths, paths_tmp):
            shutil.copy(path_tmp, path)


def refresh_paths(paths: list[Path], parallel: bool = False):
    xlwings = _require_xlwings()
    with xlwings.App() as _:
        if parallel:
            with Pool() as p:
                p.map(refresh_path, paths)
        else:
            for path in paths:
                refresh_path(path)


def refresh(
    root: Path = DEFAULT_REFRESH_ROOT,
    *,
    inplace: bool = False,
    parallel: bool = False,
) -> None:
    """Refresh cached values for all xlsx files in a folder."""
    xlwings = _require_xlwings()
    paths = list(sorted(root.glob("*.xlsx")))
    with xlwings.App() as _:
        if inplace:
            # For cases where it is fine to use the original paths (avoids copies)
            refresh_paths(paths, parallel=parallel)
        else:
            # Using a local temporal folder helps with network path issues in Excel
            refresh_paths_in_tempdir(paths, parallel=parallel)


def refresher():
    """Backward-compatible Python entry point for refresh."""
    refresh()
