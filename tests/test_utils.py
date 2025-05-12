import pytest
import numpy as np

from named_xlsx.utils import XLSXAddress, nanaverage


def test_xlsx_functions():
    cases = {
        "A10": dict(size=1, is_range=False, sheet=None, cells="A10"),
        "Sheet!A10": dict(size=1, is_range=False, sheet="Sheet", cells="A10"),
        "Sheet!A10:D10": dict(size=4, is_range=True, sheet="Sheet", cells="A10:D10"),
        "Sheet!A10:B11": dict(size=4, is_range=True, sheet="Sheet", cells="A10:B11"),
    }

    for addr, out in cases.items():
        addr_ = XLSXAddress(addr)
        assert addr_.size == out["size"]
        assert addr_.is_range == out["is_range"]


def test_as_array():
    cases = [
        ("A10:D10", dict(a=np.array(["A10", "B10", "C10", "D10"]), squeeze=True)),
        (
            "A10:A13",
            dict(a=np.array(["A10", "A11", "A12", "A13"]), squeeze=True),
        ),
        ("A10:B11", dict(a=np.array([["A10", "B10"], ["A11", "B11"]]), squeeze=True)),
        ("A10:D10", dict(a=np.array([["A10", "B10", "C10", "D10"]]), squeeze=False)),
        (
            "A10:A13",
            dict(a=np.array([["A10", "A11", "A12", "A13"]]).T, squeeze=False),
        ),
        ("A10:B11", dict(a=np.array([["A10", "B10"], ["A11", "B11"]]), squeeze=False)),
    ]
    for addr, cfg in cases:
        addr_ = XLSXAddress(addr)
        cal = addr_.as_array(squeeze=cfg["squeeze"])
        val = addr_.as_array(squeeze=cfg["squeeze"], order="row")
        assert np.all(cal == val), "Unexpected change in defaults."
        val = cfg["a"]
        assert np.all(cal == val)
        cal = addr_.as_array(squeeze=cfg["squeeze"], order="col")
        val = cfg["a"].T
        assert np.all(cal == val)


def test_methods():
    cases = [
        dict(addr="A10", format="A10", shape=(1, 1), size=1),
        dict(addr="Sheet!A10", format="Sheet!A10", shape=(1, 1), size=1),
        dict(addr="A10:B15", format="A10:B15", shape=(6, 2), size=12),
        dict(addr="A!A10:B15", format="A!A10:B15", shape=(6, 2), size=12),
        dict(addr="A B!A10:B15", format="A B!A10:B15", shape=(6, 2), size=12),
    ]
    for case in cases:
        addr_ = XLSXAddress(case["addr"])
        assert addr_.format() == case["format"]
        assert addr_.shape == case["shape"]
        assert addr_.size == case["size"]


def test_nanaverage():
    cases = [
        dict(arr=np.array([1, 1, 1]), w=None, out=1),
        dict(arr=np.array([10, float("nan"), 999]), w=np.array([1, 1, 0]), out=10),
        dict(arr=np.array([10, float("nan"), 999]), w=np.array([0.5, 1, 0]), out=10),
    ]
    for case in cases:
        assert nanaverage(case["arr"], weights=case["w"]) == case["out"]
    cases = [
        dict(type="arr", arr=[10, float("nan")], w=np.array([1, 1])),
        dict(type="weights", arr=np.array([10, float("nan")]), w=[0.5, 1]),
    ]
    for case in cases:
        match = f"Argument `{case['type']}` must be a numpy array .*"
        with pytest.raises(ValueError, match=match):
            nanaverage(case["arr"], weights=case["w"])
