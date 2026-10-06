"""End-to-end gate: Word IR → DOCX → read-back → validation, headlessly.

This is the acceptance path named in PIP-4282: the compile half runs with no
Word installed, and the host-limited half is reported honestly rather than
faked.
"""

from __future__ import annotations

import json
from pathlib import Path

from dcc_mcp_word.capabilities import HOST_LIMITED, get
from dcc_mcp_word.compiler import compile_document
from dcc_mcp_word.readback import read_back
from dcc_mcp_word.validate import validate_artifacts, validate_envelope
from dcc_mcp_word.word_ir import load_word_ir

# A production-shaped document: the fpt → Word hand-off the issue asks for.
CALL_SHEET_IR = {
    "schema_version": "office-ir/1.0",
    "kind": "word",
    "document_id": "fpt:call-sheet-2026-10-07",
    "metadata": {"title": "Call sheet — 2026-10-07", "author": "pipeline", "language": "en"},
    "document": {
        "styles": ["Heading 1", "Normal"],
        "paragraphs": [{"text": "Unit 2 — exterior night", "style": "Normal"}],
        "sections": [
            {
                "title": "Schedule",
                "paragraphs": [{"text": "Call time 07:00, wrap 19:30", "style": "Normal"}],
            }
        ],
        "lists": [{"items": ["Bring rain cover", "Check generator"], "style": "List Bullet"}],
        "tables": [
            {"header": True, "rows": [["Shot", "Status", "VFX"], ["sh010", "ip", "yes"], ["sh012", "wip", "no"]]}
        ],
        "headers": [{"section_index": 0, "text": "Call sheet — 2026-10-07"}],
        "footers": [{"section_index": 0, "text": "Page"}],
        "fields": [{"kind": "toc"}, {"kind": "page"}],
        "review_policy": {"track_changes": False, "comments_locked": False},
    },
    "outputs": ["docx"],
}


def test_full_pipeline_from_ir_to_verified_artifact(tmp_path: Path) -> None:
    ir_path = tmp_path / "call-sheet.json"
    ir_path.write_text(json.dumps(CALL_SHEET_IR), encoding="utf-8")

    envelope = load_word_ir(ir_path)
    validation = validate_envelope(envelope)
    assert validation["ok"], validation["checks"]

    docx = compile_document(envelope, tmp_path / "call-sheet.docx")
    report = read_back(envelope, docx)
    assert report.ok, report.mismatches
    assert report.checked_tables == 1
    # 3 rows x 3 columns in the call-sheet table.
    assert report.checked_cells == 9
    assert report.fields_present == 2

    artifacts = validate_artifacts([str(docx)])
    assert artifacts["ok"] is True


def test_generated_document_survives_a_reopen_by_a_fresh_reader(tmp_path: Path) -> None:
    """The artifact is a real DOCX, not a python-docx private format."""
    import zipfile

    envelope = load_word_ir(CALL_SHEET_IR)
    docx = compile_document(envelope, tmp_path / "doc.docx")

    assert zipfile.is_zipfile(docx)
    with zipfile.ZipFile(docx) as archive:
        names = archive.namelist()
    assert "word/document.xml" in names
    assert "[Content_Types].xml" in names


def test_toc_and_page_results_are_host_limited() -> None:
    """The gate writes the instructions; Word produces the results."""
    for name in ("word.toc.rebuild", "word.fields.update"):
        capability = get(name)
        assert capability is not None
        assert capability.grade == HOST_LIMITED


def test_validation_warns_about_host_limited_fields() -> None:
    envelope = load_word_ir(CALL_SHEET_IR)
    validation = validate_envelope(envelope)
    joined = " ".join(validation["warnings"])
    assert "TOC" in joined
    assert "page" in joined


def test_validation_warns_about_undeclared_figure_resources() -> None:
    ir = json.loads(json.dumps(CALL_SHEET_IR))
    ir["document"]["figures"] = [{"resource": "missing", "caption": "c"}]
    validation = validate_envelope(load_word_ir(ir))
    assert any("not declared in envelope.resources" in w for w in validation["warnings"])


def test_validation_warns_about_ragged_table_rows() -> None:
    ir = json.loads(json.dumps(CALL_SHEET_IR))
    ir["document"]["tables"] = [{"header": True, "rows": [["a", "b"], ["a"]]}]
    validation = validate_envelope(load_word_ir(ir))
    assert any("ragged rows" in w for w in validation["warnings"])


def test_validation_warns_about_non_builtin_styles() -> None:
    ir = json.loads(json.dumps(CALL_SHEET_IR))
    ir["document"]["styles"] = ["Heading 1", "Studio Custom"]
    validation = validate_envelope(load_word_ir(ir))
    assert any("Studio Custom" in w for w in validation["warnings"])


def test_empty_document_fails_validation() -> None:
    ir = json.loads(json.dumps(CALL_SHEET_IR))
    ir["document"] = {}
    validation = validate_envelope(load_word_ir(ir))
    assert validation["ok"] is False
    assert any(c["name"] == "has_content" and not c["ok"] for c in validation["checks"])
