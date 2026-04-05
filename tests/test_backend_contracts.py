import subprocess
import sys

from named_xlsx.engines import Calamine, OpenPYXL


def test_optional_backends_are_lazy_loaded():
    cmd = [
        sys.executable,
        "-c",
        (
            "import sys; "
            "import named_xlsx.engines as eng; "
            "print('xlwings' in sys.modules); "
            "print('python_calamine' in sys.modules); "
            "print(sorted(eng.ENGINES))"
        ),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, check=True)
    lines = proc.stdout.strip().splitlines()
    assert lines[0] == "False"
    assert lines[1] == "False"
    assert "OpenPYXL" in lines[2]


def test_engine_read_only_contract():
    assert OpenPYXL.read_only is False
    assert Calamine.read_only is True
