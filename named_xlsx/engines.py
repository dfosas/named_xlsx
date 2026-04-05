# coding=utf-8
"""Engines."""

from abc import abstractmethod
import importlib
import importlib.util
from typing import Any, Type

import toml
import numpy as np
import pandas as pd
import openpyxl as xl
from openpyxl.utils.cell import coordinate_to_tuple

from named_xlsx.utils import (
    Addresslike,
    MaybeStr,
    MaybePathlike,
    Pathlike,
    XLSXAddress,
    get_tables,
    get_destinations,
)

OptionalCFG = dict[str, Any] | None
SPECIFICATION_COLUMNS = ["name", "addr", "sheet", "coord", "value"]


__all__ = ["ENGINES", "OpenPYXL"]


def _module_available(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def _import_xlwings_book():
    try:
        from xlwings import Book
    except ImportError as exc:
        raise RuntimeError(
            "XLWings engine requires the optional 'xlsx' dependencies. "
            "Install them with `named_xlsx[xlsx]`."
        ) from exc
    return Book


def _import_calamine_workbook():
    try:
        from python_calamine import CalamineWorkbook
    except ImportError as exc:
        raise RuntimeError(
            "Calamine engine requires the optional 'calamine' dependencies. "
            "Install them with `named_xlsx[calamine]`."
        ) from exc
    return CalamineWorkbook


class AbstractEngine:
    """General Engine interface."""

    read_only = False

    def __init__(self, wb, path: MaybePathlike = None):
        self.wb = wb
        self.path = path

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @abstractmethod
    def name_address(self, name: str) -> XLSXAddress:
        pass

    @property
    @abstractmethod
    def names(self):
        pass

    @staticmethod
    def _load_address(x: Addresslike, /) -> XLSXAddress:
        if isinstance(x, str):
            out = XLSXAddress(x)
        else:
            out = x
        return out

    @staticmethod
    def _range_cell_addresses(addr: XLSXAddress) -> list[str]:
        coords = addr.as_array(squeeze=False).reshape(-1)
        return [XLSXAddress.from_parts(addr.sheet, coord).format() for coord in coords]

    @staticmethod
    def _range_values_as_array(addr: XLSXAddress, values: Any) -> np.ndarray:
        arr = np.asarray(values, dtype=object)
        if arr.size != addr.size:
            addrs = AbstractEngine._range_cell_addresses(addr)
            raise ValueError(f"Cannot broadcast {values=} to {addrs=}.")
        return arr.reshape(addr.shape)

    def read(self, addr: Addresslike, read_as=None, hook=None):
        addr = self._load_address(addr)
        if addr.is_range:
            v = self._read_range(addr, dtype=read_as)
        else:
            v = self._read(addr, dtype=read_as)
        if hook:
            v = hook(v)
        return v

    @abstractmethod
    def _read(self, addr, dtype=None):
        pass

    def _read_range(self, addr, dtype=None):
        addr = self._load_address(addr)
        coords = self._range_cell_addresses(addr)
        values = np.array([self.read(i) for i in coords], dtype=dtype)
        return np.squeeze(values.reshape(addr.shape))

    @abstractmethod
    def read_via_name(self, name, **kwargs):
        pass

    def _ensure_writable(self):
        if self.read_only:
            raise NotImplementedError(f"{self.__class__.__name__} engine is read-only.")

    def write(self, addr: Addresslike, value: Any):
        self._ensure_writable()
        addr = self._load_address(addr)
        if addr.is_range:
            self._write_range(addr.format(), value)
        else:
            self._write(addr.format(), value)

    @abstractmethod
    def write_via_name(self, name, value: Any):
        pass

    @abstractmethod
    def _write(self, addr: Addresslike, value: Any):
        pass

    def _write_range(self, addr: Addresslike, values: list[Any]):
        addr = self._load_address(addr)
        addr_range = self._range_cell_addresses(addr)
        values_array = self._range_values_as_array(addr, values).reshape(-1)
        for cell_addr, cell_value in zip(addr_range, values_array):
            self.write(cell_addr, cell_value)
        return self

    def save(self, f: MaybePathlike = None):
        self._ensure_writable()
        if f is None:
            if self.path is None:
                raise ValueError(f"Need a file path. {f=} and {self.path=}")
            out = self.path
        else:
            out = f
        self._save(out)

    @abstractmethod
    def close(self):
        pass

    @abstractmethod
    def _save(self, f):
        pass

    @classmethod
    @abstractmethod
    def from_file(cls, path: Pathlike, **kwargs):
        pass

    def __repr__(self):
        return f"{self.__class__.__name__}({self.path})"

    def names_as_dict(self, filter_prefix: MaybeStr = None):
        if filter_prefix is None:
            filter_prefix = ""
        out = {
            name: self.read_via_name(name)
            for name in self.names
            if name.startswith(filter_prefix)
        }
        return out

    def specifications(self, filter_prefix: MaybeStr = None) -> pd.DataFrame:
        names = self.names_as_dict(filter_prefix=filter_prefix)
        addrs = {name: self.name_address(name) for name in names}
        records = [
            dict(name=k, addr=v, sheet=v.sheet, coord=v.coord, value=names[k])
            for k, v in addrs.items()
        ]
        return pd.DataFrame.from_records(records, columns=SPECIFICATION_COLUMNS)

    def export(self, filter_prefix: MaybeStr = None) -> str:
        def normalize(value):
            if isinstance(value, np.ndarray):
                return [normalize(i) for i in value.tolist()]
            if isinstance(value, np.generic):
                return value.item()
            return value

        df = self.specifications(filter_prefix=filter_prefix)
        if df.empty:
            return ""
        df = df.assign(value=df["value"].map(normalize))
        parts = [
            {
                g: df_.set_index("name")["value"].to_dict()
                for g, df_ in df.groupby("sheet")
            }
        ]
        out = "\n\n".join([toml.dumps(part) for part in parts])
        return out


Engine = Type[AbstractEngine]


class OpenPYXL(AbstractEngine):
    """oenpyxl engine."""

    @property
    def names(self):
        return list(self.wb.defined_names)

    def name_address(self, name: str) -> XLSXAddress:
        available_names = self.names
        if name not in available_names:
            raise ValueError(f"{name=} not in {available_names=}")
        dn = self.wb.defined_names[name]
        tables = get_tables(self.wb)
        try:
            address, *_ = get_destinations(dn, tables=tables)
        except ValueError as e:
            raise ValueError(f"Cannot retrieve address for {name=}") from e
        if len(_) != 0:
            raise ValueError(f"Multiple destinations not implemented: {dn=}")
        return XLSXAddress(f"{address[0]}!{address[1]}")

    def _read(self, addr: Addresslike, dtype=None):
        addr = self._load_address(addr)
        v = self.wb[addr.sheet][addr.coord].value
        if dtype is not None:
            v = dtype(v)
        return v

    def _read_range(self, addr: Addresslike, dtype=None) -> np.ndarray:
        addr = self._load_address(addr)
        sheet = addr.sheet
        coord = addr.coord
        gen = (cell.value for row in self.wb[sheet][coord] for cell in row)
        values = np.array(list(gen), dtype=dtype)  # read `gen` before applying `dtype`
        return np.squeeze(values.reshape(addr.shape))

    def read_via_name(self, name, **kwargs):
        addr = self.name_address(name)
        return self.read(addr, **kwargs)

    def write_via_name(self, name: str, value: Any):
        addr = self.name_address(name)
        self.write(addr=addr, value=value)

    def _write(self, addr, value):
        addr_ = XLSXAddress(addr)
        self.wb[addr_.sheet][addr_.coord] = value
        return self

    def _save(self, f):
        self.wb.save(filename=f)
        return self

    def close(self):
        self.wb.close()

    @classmethod
    def from_file(cls, path: Pathlike, **kwargs):
        return cls(wb=xl.load_workbook(str(path), **kwargs), path=path)


ENGINES: dict[str, Engine] = {"OpenPYXL": OpenPYXL}

if _module_available("xlwings"):

    class XLWings(AbstractEngine):
        """xlwings engine."""

        @property
        def names(self):
            return [i.name for i in self.wb.names]

        def name_address(self, name: str) -> XLSXAddress:
            obj = self.wb.names(name).refers_to_range
            return XLSXAddress(f"{obj.sheet.name}!{obj.address}")

        def _read(self, addr: Addresslike, dtype=None):
            addr = self._load_address(addr)
            v = self.wb.sheets[addr.sheet].range(addr.coord).value
            if dtype is not None:
                v = dtype(v)
            return v

        def read_via_name(self, name, **kwargs):
            obj = self.wb.names(name).refers_to_range
            coords = obj.address
            sheet = obj.sheet.name
            return self.read(f"{sheet}!{coords}", **kwargs)

        def write_via_name(self, name: str, value: Any):
            addr = self.name_address(name).value
            return self.write(addr=addr, value=value)

        def _write(self, addr, value):
            addr_ = XLSXAddress(addr)
            ws = self.wb.sheets[addr_.sheet]
            ws.range(addr_.coord).value = value
            return self

        def _save(self, f):
            self.wb.save(path=f)
            return self

        def close(self):
            self.wb.close()

        @classmethod
        def from_file(cls, path: Pathlike, **kwargs):
            Book = _import_xlwings_book()
            return cls(wb=Book(str(path), **kwargs), path=path)

    ENGINES["XLWings"] = XLWings
    __all__.append("XLWings")

if _module_available("python_calamine"):

    class Calamine(AbstractEngine):
        """Read-only engine powered by python-calamine."""

        read_only = True

        def __init__(self, wb, path: MaybePathlike = None, names_wb=None):
            super().__init__(wb=wb, path=path)
            self.names_wb = names_wb
            self._sheet_data: dict[str, list[list[Any]]] = {}

        @property
        def names(self):
            return list(self.names_wb.defined_names)

        def _sheet_values(self, sheet_name: str) -> list[list[Any]]:
            if sheet_name not in self._sheet_data:
                sheet = self.wb.get_sheet_by_name(sheet_name)
                self._sheet_data[sheet_name] = sheet.to_python(skip_empty_area=False)
            return self._sheet_data[sheet_name]

        def name_address(self, name: str) -> XLSXAddress:
            available_names = self.names
            if name not in available_names:
                raise ValueError(f"{name=} not in {available_names=}")
            dn = self.names_wb.defined_names[name]
            tables = get_tables(self.names_wb)
            try:
                address, *_ = get_destinations(dn, tables=tables)
            except ValueError as e:
                raise ValueError(f"Cannot retrieve address for {name=}") from e
            if len(_) != 0:
                raise ValueError(f"Multiple destinations not implemented: {dn=}")
            return XLSXAddress(f"{address[0]}!{address[1]}")

        def _read(self, addr: Addresslike, dtype=None):
            addr = self._load_address(addr)
            row, col = coordinate_to_tuple(addr.coord)
            data = self._sheet_values(addr.sheet)
            if row > len(data):
                value = None
            else:
                row_data = data[row - 1]
                value = row_data[col - 1] if col <= len(row_data) else None
            if dtype is not None:
                value = dtype(value)
            return value

        def read_via_name(self, name, **kwargs):
            addr = self.name_address(name)
            return self.read(addr, **kwargs)

        def write_via_name(self, name: str, value: Any):
            self._ensure_writable()

        def _write(self, addr: Addresslike, value: Any):
            self._ensure_writable()

        def _save(self, f):
            self._ensure_writable()

        def close(self):
            self.wb.close()
            self.names_wb.close()

        @classmethod
        def from_file(cls, path: Pathlike, **kwargs):
            data_only = kwargs.pop("data_only", False)
            names_wb = xl.load_workbook(str(path), data_only=data_only)
            CalamineWorkbook = _import_calamine_workbook()
            wb = CalamineWorkbook.from_path(str(path))
            return cls(wb=wb, names_wb=names_wb, path=path)

    ENGINES["Calamine"] = Calamine
    __all__.append("Calamine")
