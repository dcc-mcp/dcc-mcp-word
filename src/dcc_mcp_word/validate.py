"""Structural validation for documents and artifacts (proposal §17/§18.1).

Two layers, kept separate on purpose:

- `validate_envelope` — contract + structural checks on the Word IR, before
  any file is produced.
- `validate_artifacts` — produced files exist and are non-empty.

Honest reporting: every check returns pass/fail with a reason, and
host-limited features are reported as declared-but-unverified rather than
checked and passed.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .capabilities import HOST_LIMITED
from .word_ir import BUILTIN_STYLES, DOCUMENT_KIND, IR_VERSION, WordEnvelope


@dataclass(frozen=True)
class Check:
    name: str
    ok: bool
    detail: str = ""


def validate_envelope(envelope: WordEnvelope) -> dict[str, Any]:
    """Validate a WordEnvelope; returns {ok, checks, warnings}."""
    checks: list[Check] = []
    warnings: list[str] = []

    checks.append(Check("schema_version", envelope.schema_version == IR_VERSION, envelope.schema_version))
    checks.append(Check("document_kind", envelope.kind == DOCUMENT_KIND, envelope.kind))
    has_content = bool(
        envelope.document.paragraphs
        or envelope.document.sections
        or envelope.document.lists
        or envelope.document.tables
    )
    checks.append(Check("has_content", has_content, "paragraphs / sections / lists / tables all empty"))
    checks.append(Check("has_title", bool(envelope.metadata.title.strip()), envelope.metadata.title))

    # A style the default template does not carry is not an error — the target
    # template may define it — but it is the single most common reason a
    # compiled document looks wrong, so it is reported.
    for style in envelope.document.styles:
        if style not in BUILTIN_STYLES:
            warnings.append(
                f"style '{style}' is not a built-in python-docx style; "
                "compiling falls back to Normal unless the target template provides it"
            )

    for table_index, table in enumerate(envelope.document.tables):
        widths = {len(row) for row in table.rows}
        if len(widths) > 1:
            warnings.append(f"table[{table_index}]: ragged rows (widths {sorted(widths)}); short rows leave trailing cells empty")
        if table.header and not table.rows:
            warnings.append(f"table[{table_index}]: header requested but the table has no rows")

    for figure in envelope.document.figures:
        known = {r.id for r in envelope.resources}
        if figure.resource not in known:
            warnings.append(
                f"figure resource '{figure.resource}' is not declared in envelope.resources; "
                f"inline figures are unimplemented in v0.1.0"
            )

    if envelope.document.content_controls:
        warnings.append(
            f"{len(envelope.document.content_controls)} content control(s) declared: parsed and validated, "
            "but not materialised by the headless writer in v0.1.0"
        )
    if any(f.kind == "toc" for f in envelope.document.fields):
        warnings.append(
            f"TOC field declared: the field instruction is written, but {HOST_LIMITED} "
            "'word.toc.rebuild' means the entries appear only when Word refreshes it"
        )
    for field in envelope.document.fields:
        if field.kind in ("page", "date"):
            warnings.append(
                f"'{field.kind}' field declared: results are produced by Word on open, not by the headless writer"
            )
    if envelope.document.review_policy.track_changes:
        warnings.append("review_policy.track_changes=true: tracked-change inspection is host_limited")

    ok = all(c.ok for c in checks)
    return {"ok": ok, "checks": [c.__dict__ for c in checks], "warnings": warnings}


def validate_artifacts(paths: list[str | Path]) -> dict[str, Any]:
    """Check produced artifacts exist and are non-empty."""
    results = []
    for raw in paths:
        p = Path(raw)
        results.append({"path": str(p), "ok": p.is_file() and p.stat().st_size > 0})
    return {"ok": all(r["ok"] for r in results), "artifacts": results}
