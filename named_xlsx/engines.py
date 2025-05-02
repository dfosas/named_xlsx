# coding=utf-8
from abc import abstractmethod
from contextlib import suppress
from typing import Any, Type

import numpy as np
import openpyxl as xl

from named_xlsx.utils import Pathlike, XLSXAddress, get_tables, get_destinations

OptionalCFG = dict[str, Any] | None


__all__ = ["ENGINES", "OpenPYXL"]


class AbstractEngine:
    def __init__(self, wb, path: Pathlike):
        self.wb = wb
        self.path = path

    def read(self, addr: str, read_as=None, hook=None):
        if XLSXAddress(addr).is_range:
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
        coords = XLSXAddress(addr).as_array()
        return np.array([self.read(i) for i in coords], dtype=dtype)

    @abstractmethod
    def read_from_name(self, name, **kwargs):
        pass

    def write(self, addr: str, value: Any):
        if XLSXAddress(addr).is_range:
            self._write_range(addr, value)
        else:
            self._write(addr, value)

    @abstractmethod
    def write_to_name(self, name, value: Any):
        pass

    @abstractmethod
    def _write(self, addr: str, value: Any):
        return self

    def _write_range(self, addr: str, values: list[Any]):
        addr_ = XLSXAddress(addr)
        addr_range = [f"{addr_.sheet}!{coord}" for coord in addr_.as_array()]
        if len(addr_range) != len(values):
            raise ValueError(f"Cannot broadcast {values=} to {addr_range=}.")
        for cell_addr, cell_value in zip(addr_range, values):
            self.write(cell_addr, cell_value)
        return self

    def save(self, f: Pathlike | None = None):
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


Engine = Type[AbstractEngine]


class OpenPYXL(AbstractEngine):
    def _read(self, addr: str, dtype=None):
        addr_ = XLSXAddress(addr)
        v = self.wb[addr_.sheet][addr_.coord].value
        if dtype is not None:
            v = dtype(v)
        return v

    def _read_range(self, addr: str, dtype=None) -> np.ndarray:
        addr_ = XLSXAddress(addr)
        sheet = addr_.sheet
        coord = addr_.coord
        gen = (cell.value for row in self.wb[sheet][coord] for cell in row)
        return np.array(list(gen), dtype=dtype)  # read `gen` before applying `dtype`

    def read_from_name(self, name, **kwargs):
        addr = self._get_name_address(name)
        return self.read(addr, **kwargs)

    def write_to_name(self, name: str, value: Any):
        addr = self._get_name_address(name)
        self.write(addr=addr, value=value)

    def _get_name_address(self, name: str) -> str:
        available_names = list(self.wb.defined_names)
        if name not in available_names:
            raise ValueError(f"{name=} not in {available_names=}")
        dn = self.wb.defined_names[name]
        tables = get_tables(self.wb)
        address, *_ = get_destinations(dn, tables=tables)
        if len(_) != 0:
            raise ValueError(f"Multiple destinations not implemented: {dn=}")
        return f"{address[0]}!{address[1]}"

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

with suppress(ImportError):
    from xlwings import Book

    class XLWings(AbstractEngine):
        def _read(self, addr: str, dtype=None):
            addr_ = XLSXAddress(addr)
            v = self.wb.sheets[addr_.sheet].range(addr_.coord).value
            if dtype is not None:
                v = dtype(v)
            return v

        def read_from_name(self, name, **kwargs):
            obj = self.wb.names(name).refers_to_range
            coords = obj.address
            sheet = obj.sheet.name
            return self.read(f"{sheet}!{coords}", **kwargs)

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
            return cls(wb=Book(str(path), **kwargs), path=path)

    ENGINES["XLWings"] = XLWings
    __all__.append("XLWings")
