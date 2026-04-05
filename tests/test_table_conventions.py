from pathlib import Path

import numpy as np
import openpyxl as xl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

from named_xlsx.engines import OpenPYXL


def build_table_workbook(path: Path) -> Path:
    wb = xl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["item", "value"])
    ws.append(["a", 10])
    ws.append(["b", 20])
    ws.append(["c", 30])
    ws.append(["Total", "=SUM(B2:B4)"])

    table = Table(displayName="tbl_demo", ref="A1:B5")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    wb.defined_names.add(DefinedName("series", attr_text="tbl_demo[value]"))
    wb.defined_names.add(DefinedName("t.series", attr_text="tbl_demo[value]"))
    wb.save(path)
    wb.close()
    return path


def test_table_defined_name_resolves_to_data_rows_without_special_prefix(
    tmp_path: Path,
):
    path = build_table_workbook(tmp_path / "table.xlsx")
    with OpenPYXL.from_file(path) as engine:
        addr = engine.name_address("series")
        assert addr.sheet == "Sheet1"
        assert addr.coord == "B2:B4"
        assert np.array_equal(engine.read_via_name("series"), np.array([10, 20, 30]))


def test_table_defined_name_legacy_prefix_still_works(tmp_path: Path):
    path = build_table_workbook(tmp_path / "table.xlsx")
    with OpenPYXL.from_file(path) as engine:
        assert np.array_equal(engine.read_via_name("t.series"), np.array([10, 20, 30]))


def test_table_defined_name_appears_in_export(tmp_path: Path):
    path = build_table_workbook(tmp_path / "table.xlsx")
    with OpenPYXL.from_file(path, data_only=True) as engine:
        exported = engine.export()
    assert (
        exported == '[Sheet1]\nseries = [ 10, 20, 30,]\n"t.series" = [ 10, 20, 30,]\n'
    )
