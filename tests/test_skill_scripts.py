"""Skill-script contract tests — scripts run as subprocesses, as the gateway does."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parent.parent / "src" / "dcc_mcp_word" / "skills"

IR = {
    "schema_version": "office-ir/1.0",
    "kind": "word",
    "document_id": "draft:call-sheet",
    "metadata": {"title": "Call sheet", "author": "pipeline", "language": "en"},
    "document": {
        "sections": [{"title": "Tomorrow", "paragraphs": [{"text": "Call time 07:00"}]}],
        "tables": [{"header": True, "rows": [["Shot", "Status"], ["sh010", "ip"]]}],
        "footers": [{"section_index": 0, "text": "Call sheet"}],
        "fields": [{"kind": "page"}],
    },
    "outputs": ["docx"],
}


def _run(script: Path, payload: dict, env: dict) -> dict:
    result = subprocess.run(
        [sys.executable, str(script)],
        input=json.dumps(payload),
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        check=False,
    )
    assert result.returncode == 0, result.stderr
    return json.loads(result.stdout)


def test_generate_document_compiles_and_verifies(tmp_path: Path, skill_env: dict) -> None:
    ir_path = tmp_path / "doc.json"
    ir_path.write_text(json.dumps(IR), encoding="utf-8")
    out = _run(
        SKILLS / "word-document" / "scripts" / "generate_document.py",
        {"input": str(ir_path), "output_dir": str(tmp_path / "out")},
        skill_env,
    )
    assert out["success"] is True, out
    context = out["context"]
    assert Path(context["artifact"]).is_file()
    assert context["read_back"]["ok"] is True
    assert context["backend"] == "openxml"
    assert "host_limited" in context["note"]


def test_generate_document_reports_validation_failure(tmp_path: Path, skill_env: dict) -> None:
    bad = {**IR, "document": {}}
    ir_path = tmp_path / "bad.json"
    ir_path.write_text(json.dumps(bad), encoding="utf-8")
    out = _run(
        SKILLS / "word-document" / "scripts" / "generate_document.py",
        {"input": str(ir_path), "output_dir": str(tmp_path / "out")},
        skill_env,
    )
    assert out["success"] is False
    assert "structural validation" in out["message"]


def test_validate_document_accepts_a_valid_ir(tmp_path: Path, skill_env: dict) -> None:
    ir_path = tmp_path / "doc.json"
    ir_path.write_text(json.dumps(IR), encoding="utf-8")
    out = _run(
        SKILLS / "word-document" / "scripts" / "validate_document.py",
        {"input": str(ir_path)},
        skill_env,
    )
    assert out["success"] is True


def test_validate_document_checks_artifact_directory(tmp_path: Path, skill_env: dict) -> None:
    artifacts = tmp_path / "artifacts"
    artifacts.mkdir()
    (artifacts / "out.docx").write_bytes(b"pk")
    out = _run(
        SKILLS / "word-document" / "scripts" / "validate_document.py",
        {"input": str(artifacts)},
        skill_env,
    )
    assert out["success"] is True
    assert out["context"]["artifacts"][0]["ok"] is True


def test_inspect_document_inventories_an_artifact(tmp_path: Path, skill_env: dict) -> None:
    ir_path = tmp_path / "doc.json"
    ir_path.write_text(json.dumps(IR), encoding="utf-8")
    generated = _run(
        SKILLS / "word-document" / "scripts" / "generate_document.py",
        {"input": str(ir_path), "output_dir": str(tmp_path / "out")},
        skill_env,
    )
    out = _run(
        SKILLS / "word-document" / "scripts" / "inspect_document.py",
        {"input": generated["context"]["artifact"]},
        skill_env,
    )
    assert out["success"] is True
    assert out["context"]["table_count"] == 1
    assert out["context"]["fields"] >= 1


def test_inspect_document_reports_a_missing_file(tmp_path: Path, skill_env: dict) -> None:
    out = _run(
        SKILLS / "word-document" / "scripts" / "inspect_document.py",
        {"input": str(tmp_path / "absent.docx")},
        skill_env,
    )
    assert out["success"] is False
    assert "not found" in out["message"]


def test_capabilities_script_reports_grades(skill_env: dict) -> None:
    out = _run(SKILLS / "word-capabilities" / "scripts" / "capabilities.py", {}, skill_env)
    assert out["success"] is True
    assert "word.document.compile" in out["context"]["verified"]
    assert "word.toc.rebuild" in out["context"]["host_limited"]


def test_preflight_script_never_fails_on_missing_word(skill_env: dict) -> None:
    out = _run(SKILLS / "word-capabilities" / "scripts" / "preflight.py", {}, skill_env)
    assert out["success"] is True
    assert out["context"]["app"] == "word"
