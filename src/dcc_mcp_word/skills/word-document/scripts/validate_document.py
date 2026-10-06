"""word-document / validate_document — validate a Word IR or produced artifacts.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"input": ...}
2. CLI flags: --input
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from dcc_mcp_word.validate import validate_artifacts, validate_envelope
from dcc_mcp_word.word_ir import load_word_ir


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    target = Path(params["input"])
    if target.is_dir():
        paths = sorted(str(p) for p in target.rglob("*") if p.is_file())
        report = validate_artifacts(paths)
        print(
            json.dumps(
                {
                    "success": report["ok"],
                    "message": f"checked {len(paths)} artifact(s) under {target.name}",
                    "context": report,
                },
                ensure_ascii=False,
            )
        )
        return

    envelope = load_word_ir(target)
    report = validate_envelope(envelope)
    print(
        json.dumps(
            {
                "success": report["ok"],
                "message": f"word IR '{envelope.document_id}' validation {'passed' if report['ok'] else 'failed'}",
                "context": report,
            },
            ensure_ascii=False,
        )
    )


def main() -> None:
    _force_utf8_stdio()
    params: dict = {}
    if not sys.stdin.isatty():
        raw = sys.stdin.read()
        if raw.strip():
            try:
                params = json.loads(raw)
            except json.JSONDecodeError:
                params = {}
    if not params:
        parser = argparse.ArgumentParser(description="Validate a Word IR envelope or artifact directory")
        parser.add_argument("--input", required=True, help="Word IR JSON path, or a directory of artifacts")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
