# coding=utf-8
"""Commandline Interfaces."""

from typing import Annotated
from contextlib import closing
from pathlib import Path
import shutil
import toml
import typer

from named_xlsx.engines import (
    ENGINES,
    Engine,
    OpenPYXL,
    MaybeStr,
    SPECIFICATION_COLUMNS,
)
from named_xlsx.refresh import refresh

DEFAULT_REFRESH_ROOT = Path(__file__).parent


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


app = typer.Typer(
    add_completion=False,
    pretty_exceptions_enable=False,
    help="Work with Excel named cells and tables from the command line.",
)

PathArgument = Annotated[Path, typer.Argument()]
OptionalPathOption = Annotated[
    Path | None, typer.Option("--p-out", "--p_out", help="Output file path.")
]
FilterPrefixOption = Annotated[
    MaybeStr,
    typer.Option(
        "--filter-prefix",
        "--filter_prefix",
        help="Only include names with this prefix.",
    ),
]
EngineOption = Annotated[
    str | None,
    typer.Option("--engine", help="Workbook engine name to use."),
]
RootOption = Annotated[
    Path, typer.Option("--root", help="Folder containing xlsx files to refresh.")
]


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


@app.command("save")
def save_command(
    p_ini: PathArgument,
    p_out: OptionalPathOption = None,
    filter_prefix: FilterPrefixOption = None,
    engine: EngineOption = None,
) -> None:
    try:
        save(p_ini, p_out=p_out, filter_prefix=filter_prefix, engine=engine)
    except ValueError as exc:
        raise typer.BadParameter(str(exc)) from exc


@app.command("spec")
def specifications_command(
    p_xlsx: PathArgument,
    p_out: OptionalPathOption = None,
    filter_prefix: FilterPrefixOption = None,
    engine: EngineOption = None,
) -> None:
    try:
        specifications(p_xlsx, p_out=p_out, filter_prefix=filter_prefix, engine=engine)
    except ValueError as exc:
        raise typer.BadParameter(str(exc)) from exc


@app.command("load")
def load_command(
    p_toml: PathArgument,
    p_xlsx: PathArgument,
    p_out: PathArgument,
    engine: EngineOption = None,
) -> None:
    try:
        load(p_toml, p_xlsx, p_out, engine=engine)
    except ValueError as exc:
        raise typer.BadParameter(str(exc)) from exc


@app.command("refresh")
def refresh_command(
    root: RootOption = DEFAULT_REFRESH_ROOT,
    inplace: Annotated[
        bool, typer.Option("--inplace", help="Refresh workbooks in place.")
    ] = False,
    parallel: Annotated[
        bool, typer.Option("--parallel", help="Refresh files in parallel.")
    ] = False,
) -> None:
    try:
        refresh(root=root, inplace=inplace, parallel=parallel)
    except RuntimeError as exc:
        raise typer.BadParameter(str(exc)) from exc


def main() -> None:
    app()
