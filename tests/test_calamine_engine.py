from pathlib import Path

import numpy as np
import openpyxl as xl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from named_xlsx.engines import Calamine

pytest.importorskip("python_calamine")


@pytest.fixture()
def workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "calamine.xlsx"
    wb = xl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    for irow, row in enumerate(values, start=1):
        for icol, value in enumerate(row, start=1):
            ws.cell(irow, icol).value = value
    wb.defined_names.add(DefinedName("scalar", attr_text="Sheet1!$A$1"))
    wb.defined_names.add(DefinedName("matrix", attr_text="Sheet1!$A$1:$B$2"))
    wb.save(path)
    wb.close()
    return path


def test_calamine_reads_addresses_and_names(workbook_path: Path):
    with Calamine.from_file(workbook_path) as engine:
        assert engine.read("Sheet1!A1") == 1.0
        assert engine.read_via_name("scalar") == 1.0
        assert np.array_equal(
            engine.read("Sheet1!A1:B2"),
            np.array([[1.0, 2.0], [4.0, 5.0]]),
        )
        assert np.array_equal(
            engine.read_via_name("matrix"),
            np.array([[1.0, 2.0], [4.0, 5.0]]),
        )


def test_calamine_is_read_only(workbook_path: Path):
    with Calamine.from_file(workbook_path) as engine:
        with pytest.raises(NotImplementedError, match="read-only"):
            engine.write("Sheet1!A1", 10)
        with pytest.raises(NotImplementedError, match="read-only"):
            engine.save()
