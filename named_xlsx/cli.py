# coding=utf-8
from contextlib import closing
from pathlib import Path
import shutil
import fire
import toml

from named_xlsx.engines import OpenPYXL

def load(p_toml: Path, p_xlsx: Path, p_out: Path) -> None:
    cfg = {
        name: value
        for sheet, mapping in toml.load(p_toml).items()
        for name, value in mapping.items()
    }
    shutil.copy(p_xlsx, p_out)
    with closing(OpenPYXL.from_file(p_out)) as m:
        for addr, vals in cfg.items():
            m.write_via_name(addr, vals)
        m.save()

def save(p_ini: Path, p_out: Path | None = None, filter_prefix: str | None = None) -> None:
    txt = OpenPYXL.from_file(p_ini).export(filter_prefix=filter_prefix)
    if p_out is None:
        print(txt)
    else:
        p_out.write_text(txt)

def load_cli():
    fire.Fire(load)

def save_cli():
    fire.Fire(save)
