from pathlib import Path

import openpyxl as xl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from named_xlsx.cli import load, save, specifications
from named_xlsx.engines import OpenPYXL


@pytest.fixture()
def workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "base.xlsx"
    wb = xl.Workbook()
    ws1 = wb.active
    ws1.title = "sheet_1"
    ws1["A1"] = 2
    ws1["B1"] = 4
    ws2 = wb.create_sheet("sheet_2")
    ws2["A1"] = 3
    ws2["B1"] = 5

    for name, ref in {
        "i.a": "sheet_1!$A$1",
        "i.b": "sheet_1!$B$1",
        "i.x": "sheet_2!$A$1",
        "i.y": "sheet_2!$B$1",
    }.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    wb.save(path)
    wb.close()
    return path


def test_save_prints_toml(workbook_path: Path, capsys):
    save(workbook_path)
    out = capsys.readouterr().out
    assert (
        out == '[sheet_1]\n"i.a" = 2\n"i.b" = 4\n\n[sheet_2]\n"i.x" = 3\n"i.y" = 5\n\n'
    )


def test_save_filter_and_file_output(workbook_path: Path, tmp_path: Path):
    out_path = tmp_path / "saved.toml"
    save(workbook_path, p_out=out_path, filter_prefix="i.")
    assert (
        out_path.read_text()
        == '[sheet_1]\n"i.a" = 2\n"i.b" = 4\n\n[sheet_2]\n"i.x" = 3\n"i.y" = 5\n'
    )


def test_specifications_print_and_csv_output(
    workbook_path: Path, tmp_path: Path, capsys
):
    specifications(workbook_path, filter_prefix="i.")
    out = capsys.readouterr().out
    assert "name" in out
    assert "sheet_1" in out
    assert "sheet_2" in out

    out_path = tmp_path / "specs.csv"
    specifications(workbook_path, p_out=out_path, filter_prefix="i.")
    assert out_path.read_text() == (
        "name,sheet,coord,value\n"
        "i.a,sheet_1,A1,2\n"
        "i.b,sheet_1,B1,4\n"
        "i.x,sheet_2,A1,3\n"
        "i.y,sheet_2,B1,5\n"
    )


def test_load_updates_copy_without_modifying_source(
    workbook_path: Path, tmp_path: Path
):
    cfg_path = tmp_path / "update.toml"
    out_path = tmp_path / "modified.xlsx"
    cfg_path.write_text('[sheet_1]\n"i.a" = 8\n\n[sheet_2]\n"i.y" = 9\n')

    returned = load(cfg_path, workbook_path, out_path)

    assert returned == out_path

    with OpenPYXL.from_file(workbook_path, data_only=True) as engine:
        assert engine.read_via_name("i.a") == 2
        assert engine.read_via_name("i.y") == 5

    with OpenPYXL.from_file(out_path, data_only=True) as engine:
        assert engine.read_via_name("i.a") == 8
        assert engine.read_via_name("i.b") == 4
        assert engine.read_via_name("i.x") == 3
        assert engine.read_via_name("i.y") == 9
