"""Capability grading — the honest boundary of the v0.1.0 surface (PIP-4276).

Two grades, and the difference is evidence, not confidence:

- `verified` — covered by a CI-green headless test. No Word installation, no
  COM, no Windows-only path.
- `host_limited` — requires the desktop Word application (COM). The artifact
  is structurally valid without it, but the value only materializes when Word
  opens it. Never reported as verified.

Anything not listed here is unimplemented, not silently degraded: the
compiler raises on specs it cannot write rather than dropping them.

Note on `word.document.reflow`: the *compile* half (Word IR → DOCX, styles,
sections, tables, headers/footers, field instructions) is headless and
therefore `verified`. The *reflow* half — reflowing an existing document,
recomputing pagination, updating field results — needs Word's layout engine
and is `host_limited`. They are graded separately because claiming the whole
tool as verified would misrepresent what CI actually proves.
"""

from __future__ import annotations

from dataclasses import dataclass

VERIFIED = "verified"
HOST_LIMITED = "host_limited"
UNIMPLEMENTED = "unimplemented"


@dataclass(frozen=True)
class Capability:
    name: str
    grade: str
    summary: str
    evidence: str
    requires_office: bool

    @property
    def is_verified(self) -> bool:
        return self.grade == VERIFIED


CAPABILITIES: tuple[Capability, ...] = (
    Capability(
        name="word.document.compile",
        grade=VERIFIED,
        summary="Word IR (office-ir/1.0) → DOCX through the headless Open XML backend.",
        evidence="tests/test_compiler.py + tests/test_readback.py, runs without Word installed.",
        requires_office=False,
    ),
    Capability(
        name="word.document.read_back",
        grade=VERIFIED,
        summary="Reopen a compiled DOCX and compare every paragraph and table cell against the source IR.",
        evidence="tests/test_readback.py; the 1.0 write-then-read-back gate.",
        requires_office=False,
    ),
    Capability(
        name="word.paragraphs.write",
        grade=VERIFIED,
        summary="Body paragraphs and section headings with resolved styles.",
        evidence="tests/test_compiler.py::test_compile_writes_paragraphs_and_sections.",
        requires_office=False,
    ),
    Capability(
        name="word.lists.write",
        grade=VERIFIED,
        summary="Bullet and numbered list blocks.",
        evidence="tests/test_compiler.py::test_compile_writes_lists.",
        requires_office=False,
    ),
    Capability(
        name="word.tables.write",
        grade=VERIFIED,
        summary="Grid tables with an optional bold header row.",
        evidence="tests/test_readback.py::test_read_back_compares_table_cells.",
        requires_office=False,
    ),
    Capability(
        name="word.headers_footers.write",
        grade=VERIFIED,
        summary="Per-section header and footer text, including a PAGE field in the footer.",
        evidence="tests/test_compiler.py::test_compile_writes_headers_footers_and_page_field.",
        requires_office=False,
    ),
    Capability(
        name="word.fields.write",
        grade=VERIFIED,
        summary="Field instructions (TOC / PAGE / DATE / custom) are written as real Word fields.",
        evidence="tests/test_compiler.py::test_compile_writes_toc_field_instruction.",
        requires_office=False,
    ),
    Capability(
        name="word.fields.update",
        grade=HOST_LIMITED,
        summary="Recomputing field results (page numbers, dates). Word computes these on open; python-docx has no layout engine.",
        evidence="Not covered by CI. Field instructions are written; their results are produced by Word.",
        requires_office=True,
    ),
    Capability(
        name="word.toc.rebuild",
        grade=HOST_LIMITED,
        summary="Building the table of contents entries from the heading structure.",
        evidence="Not covered by CI. The TOC field is written; Word populates it on open or on F9.",
        requires_office=True,
    ),
    Capability(
        name="word.document.reflow",
        grade=HOST_LIMITED,
        summary="Reflowing an existing document: repagination, style reapplication and layout reconciliation.",
        evidence="Not covered by CI; requires the desktop Word COM backend (dcc-mcp-office WordBackend).",
        requires_office=True,
    ),
    Capability(
        name="word.track_changes.inspect",
        grade=HOST_LIMITED,
        summary="Reading tracked changes and review state from a document.",
        evidence="Not covered by CI; requires the desktop Word COM backend.",
        requires_office=True,
    ),
    Capability(
        name="word.content_controls.fill",
        grade=UNIMPLEMENTED,
        summary="Anchored content controls declared in the IR.",
        evidence="Parsed and validated by the IR, not materialised by the headless writer in v0.1.0.",
        requires_office=False,
    ),
    Capability(
        name="word.figures.inline",
        grade=UNIMPLEMENTED,
        summary="Image figures resolved through envelope resources.",
        evidence="Parsed and validated by the IR, not inlined by the headless writer in v0.1.0.",
        requires_office=False,
    ),
)

_BY_NAME = {capability.name: capability for capability in CAPABILITIES}


def get(name: str) -> Capability | None:
    return _BY_NAME.get(name)


def by_grade(grade: str) -> tuple[Capability, ...]:
    return tuple(c for c in CAPABILITIES if c.grade == grade)


def report() -> dict[str, object]:
    """Machine-readable grading summary, exposed through the skill and CLI."""
    return {
        "schema": "dcc-mcp-capability-grade/1",
        "verified": [c.name for c in by_grade(VERIFIED)],
        "host_limited": [c.name for c in by_grade(HOST_LIMITED)],
        "unimplemented": [c.name for c in by_grade(UNIMPLEMENTED)],
        "capabilities": [
            {
                "name": c.name,
                "grade": c.grade,
                "summary": c.summary,
                "evidence": c.evidence,
                "requires_office": c.requires_office,
            }
            for c in CAPABILITIES
        ],
    }
