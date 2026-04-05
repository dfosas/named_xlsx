from pathlib import Path

import numpy as np
import openpyxl as xl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from named_xlsx.engines import OpenPYXL


@pytest.fixture()
def workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "ranges.xlsx"
    wb = xl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    for irow, row in enumerate(values, start=1):
        for icol, value in enumerate(row, start=1):
            ws.cell(irow, icol).value = value

    wb.defined_names.add(DefinedName("scalar", attr_text="Sheet1!$A$1"))
    wb.defined_names.add(DefinedName("rowvec", attr_text="Sheet1!$A$1:$C$1"))
    wb.defined_names.add(DefinedName("colvec", attr_text="Sheet1!$A$1:$A$3"))
    wb.defined_names.add(DefinedName("matrix", attr_text="Sheet1!$A$1:$B$2"))
    wb.save(path)
    wb.close()
    return path


def test_read_range_shapes(workbook_path: Path):
    with OpenPYXL.from_file(workbook_path) as engine:
        assert engine.read_via_name("scalar") == 1
        assert np.array_equal(engine.read_via_name("rowvec"), np.array([1, 2, 3]))
        assert np.array_equal(engine.read_via_name("colvec"), np.array([1, 4, 7]))
        assert np.array_equal(
            engine.read_via_name("matrix"),
            np.array([[1, 2], [4, 5]]),
        )
        assert np.array_equal(
            engine.read("Sheet1!A1:B2"),
            np.array([[1, 2], [4, 5]]),
        )


def test_write_matrix_accepts_flat_values(workbook_path: Path):
    with OpenPYXL.from_file(workbook_path) as engine:
        engine.write_via_name("matrix", [10, 11, 12, 13])
        assert np.array_equal(
            engine.read_via_name("matrix"),
            np.array([[10, 11], [12, 13]]),
        )


def test_write_matrix_accepts_nested_values(workbook_path: Path):
    with OpenPYXL.from_file(workbook_path) as engine:
        engine.write("Sheet1!A1:B2", [[10, 11], [12, 13]])
        assert np.array_equal(
            engine.read("Sheet1!A1:B2"),
            np.array([[10, 11], [12, 13]]),
        )


def test_write_matrix_rejects_incompatible_values(workbook_path: Path):
    with OpenPYXL.from_file(workbook_path) as engine:
        match = "Cannot broadcast"
        with pytest.raises(ValueError, match=match):
            engine.write_via_name("matrix", [10, 11, 12])
