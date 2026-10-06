"""Runtime purity — importing the package must not pull python-docx or pywin32.

dcc-mcp-core enforces a zero-runtime-dependency rule so adapters stay
installable inside DCCs that pin their own library versions. python-docx is
an opt-in extra, imported only inside `compiler` / `readback`, exactly as
openpyxl is in dcc-mcp-excel.
"""

from __future__ import annotations

import subprocess
import sys

FORBIDDEN = ("docx", "win32com", "pythonwin")


def test_package_import_stays_stdlib_only() -> None:
    code = (
        "import sys; import dcc_mcp_word;"
        "bad = [m for m in ('" + "','".join(FORBIDDEN) + "') if m in sys.modules];"
        "print(','.join(bad))"
    )
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    assert result.stdout.strip() == "", f"importing the package pulled: {result.stdout.strip()}"


def test_capabilities_and_host_matrix_are_stdlib_only() -> None:
    code = (
        "import sys; from dcc_mcp_word import capabilities, host_matrix, validate;"
        "bad = [m for m in ('" + "','".join(FORBIDDEN) + "') if m in sys.modules];"
        "print(','.join(bad))"
    )
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    assert result.stdout.strip() == "", f"stdlib-only modules pulled: {result.stdout.strip()}"
