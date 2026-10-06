"""Version consistency — pyproject, the package and release-please must agree."""

from __future__ import annotations

import json
import re
from pathlib import Path

import dcc_mcp_word

ROOT = Path(__file__).resolve().parent.parent


def _pyproject_version() -> str:
    source = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    match = re.search(r'^version = "([0-9]+\.[0-9]+\.[0-9]+)"', source, re.MULTILINE)
    assert match is not None, "pyproject.toml has no PEP 440 version"
    return match.group(1)


def test_package_version_matches_pyproject() -> None:
    assert dcc_mcp_word.__version__ == _pyproject_version()


def test_release_please_manifest_matches_pyproject() -> None:
    manifest = json.loads((ROOT / ".release-please-manifest.json").read_text(encoding="utf-8"))
    assert manifest["."] == _pyproject_version()


def test_cli_reports_the_same_version() -> None:
    import subprocess
    import sys

    result = subprocess.run(
        [sys.executable, "-m", "dcc_mcp_word._standalone_entry", "--version"],
        capture_output=True,
        text=True,
        check=True,
    )
    assert _pyproject_version() in result.stdout
