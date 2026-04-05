# coding=utf-8
"""Commandline Interfaces."""

from contextlib import closing
from pathlib import Path
import shutil
import fire
import toml

from named_xlsx.engines import (
    ENGINES,
    Engine,
    OpenPYXL,
    MaybeStr,
    SPECIFICATION_COLUMNS,
)


def _resolve_engine(
    engine: str | Engine | None, *, require_writable: bool = False
) -> Engine:
    resolved: Engine
    if engine is None:
        resolved = OpenPYXL
    elif isinstance(engine, str):
        try:
            resolved = ENGINES[engine]
        except KeyError as exc:
            available = ", ".join(sorted(ENGINES))
            raise ValueError(
                f"Unknown engine {engine!r}. Available engines: {available}."
            ) from exc
    else:
        resolved = engine

    if require_writable and getattr(resolved, "read_only", False):
        raise ValueError(f"Engine {resolved.__name__!r} is read-only.")
    return resolved


def load(
    p_toml: Path,
    p_xlsx: Path,
    p_out: Path,
    engine: str | Engine | None = None,
) -> Path:
    """
    Load configuration to a spreadsheet and save it to a file.

    Parameters
    ----------
    p_toml
        Path to the TOML file.
    p_xlsx
        Path to the XLSX file.
    p_out
        Path to the output file.
    engine
        Workbook engine to use.

    Returns
    -------
    Path to the output file.

    """
    engine = _resolve_engine(engine, require_writable=True)
    cfg = {
        name: value
        for sheet, mapping in toml.load(p_toml).items()
        for name, value in mapping.items()
    }
    shutil.copy(p_xlsx, p_out)
    with closing(engine.from_file(p_out)) as m:
        for addr, vals in cfg.items():
            m.write_via_name(addr, vals)
        m.save()

    return p_out


def save(
    p_ini: Path,
    p_out: Path | None = None,
    filter_prefix: MaybeStr = None,
    engine: str | Engine | None = None,
) -> None:
    engine = _resolve_engine(engine)
    with closing(engine.from_file(p_ini, data_only=True)) as m:
        txt = m.export(filter_prefix=filter_prefix)
    if p_out is None:
        print(txt)
    else:
        p_out.write_text(txt)


def specifications(
    p_xlsx: Path,
    p_out: Path | None = None,
    filter_prefix: MaybeStr = None,
    engine: str | Engine | None = None,
) -> None:
    engine = _resolve_engine(engine)
    with closing(engine.from_file(p_xlsx)) as m:
        df = (
            m.specifications(filter_prefix=filter_prefix)
            .sort_values(by=["sheet", "coord", "name"])
            .reindex(columns=[col for col in SPECIFICATION_COLUMNS if col != "addr"])
        )
    if p_out is None:
        print(df)
    else:
        df.to_csv(p_out, index=False)


def specifications_cli():
    fire.Fire(specifications)


def load_cli():
    fire.Fire(load)


def save_cli():
    fire.Fire(save)
