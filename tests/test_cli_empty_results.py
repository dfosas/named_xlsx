from pathlib import Path

import openpyxl as xl
import pandas as pd
import pytest
from openpyxl.workbook.defined_name import DefinedName

from named_xlsx.cli import save, specifications
from named_xlsx.engines import OpenPYXL


@pytest.fixture()
def workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "cli.xlsx"
    wb = xl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 1
    wb.defined_names.add(DefinedName("i.a", attr_text="Sheet1!$A$1"))
    wb.save(path)
    wb.close()
    return path


def test_engine_specifications_empty_filter_has_stable_columns(workbook_path: Path):
    with OpenPYXL.from_file(workbook_path) as engine:
        df = engine.specifications(filter_prefix="missing.")
    assert list(df.columns) == ["name", "addr", "sheet", "coord", "value"]
    assert df.empty


def test_engine_export_empty_filter_returns_empty_string(workbook_path: Path):
    with OpenPYXL.from_file(workbook_path, data_only=True) as engine:
        assert engine.export(filter_prefix="missing.") == ""


def test_save_cli_empty_filter_prints_nothing(workbook_path: Path, capsys):
    save(workbook_path, filter_prefix="missing.")
    out = capsys.readouterr().out
    assert out == "\n"


def test_specifications_cli_empty_filter_prints_empty_schema(
    workbook_path: Path, capsys
):
    specifications(workbook_path, filter_prefix="missing.")
    out = capsys.readouterr().out
    expected = pd.DataFrame(columns=["name", "sheet", "coord", "value"])
    assert out == f"{expected}\n"


def test_specifications_cli_empty_filter_writes_csv(
    workbook_path: Path, tmp_path: Path
):
    out_path = tmp_path / "specs.csv"
    specifications(workbook_path, p_out=out_path, filter_prefix="missing.")
    assert out_path.read_text() == "name,sheet,coord,value\n"
