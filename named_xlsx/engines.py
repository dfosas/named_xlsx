# coding=utf-8
from abc import abstractmethod
from contextlib import suppress
from typing import Any, Type

import toml
import numpy as np
import pandas as pd
import openpyxl as xl

from named_xlsx.utils import Addresslike, Pathlike, XLSXAddress, get_tables, get_destinations

OptionalCFG = dict[str, Any] | None


__all__ = ["ENGINES", "OpenPYXL"]


class AbstractEngine:
    def __init__(self, wb, path: Pathlike | None = None):
        self.wb = wb
        self.path = path

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
        coords = XLSXAddress(addr).as_array()
        return np.array([self.read(i) for i in coords], dtype=dtype)

    @abstractmethod
    def read_via_name(self, name, **kwargs):
        pass

    def write(self, addr: Addresslike, value: Any):
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
        addr_range = [f"{addr.sheet}!{coord}" for coord in addr.as_array()]
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

    def __repr__(self):
        return f"{self.__class__.__name__}({self.path})"

    def names_as_dict(self, filter_prefix: str | None = None):
        out = {name: self.read_via_name(name) for name in self.names}
        if filter_prefix is None:
            return out
        return {k: v for k, v in out.items() if k.startswith(filter_prefix)}

    def specifications(self, filter_prefix: str | None = None) -> pd.DataFrame:
        names = self.names_as_dict(filter_prefix=filter_prefix)
        addrs = {name: self.name_address(name) for name in names}
        records = [
            dict(name=k, addr=v, sheet=v.sheet, coord=v.coord, value=names[k])
            for k, v in addrs.items()
        ]
        return pd.DataFrame.from_records(records)

    def export(self, filter_prefix: str | None = None) -> str:
        df = self.specifications(filter_prefix=filter_prefix)
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

    @property
    def names(self):
        return list(self.wb.defined_names)

    def name_address(self, name: str) -> XLSXAddress:
        available_names = self.names
        if name not in available_names:
            raise ValueError(f"{name=} not in {available_names=}")
        dn = self.wb.defined_names[name]
        tables = get_tables(self.wb)
        address, *_ = get_destinations(dn, tables=tables)
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
        return np.array(list(gen), dtype=dtype)  # read `gen` before applying `dtype`

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

with suppress(ImportError):
    from xlwings import Book

    class XLWings(AbstractEngine):

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
            return cls(wb=Book(str(path), **kwargs), path=path)

    ENGINES["XLWings"] = XLWings
    __all__.append("XLWings")
