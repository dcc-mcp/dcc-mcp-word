"""Contract tests for the Word IR (mirrors dcc-mcp-office-ir `pub mod word`)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from dcc_mcp_word.word_ir import (
    IR_VERSION,
    IrValidationError,
    artifact_stem,
    load_word_ir,
    parse_envelope,
)


def envelope(**overrides: object) -> dict:
    base = {
        "schema_version": IR_VERSION,
        "kind": "word",
        "document_id": "draft:test",
        "metadata": {"title": "Test", "author": "pipeline", "language": "en"},
        "document": {},
        "outputs": ["docx"],
    }
    base.update(overrides)
    return base


def test_parses_minimal_envelope() -> None:
    parsed = parse_envelope(envelope())
    assert parsed.kind == "word"
    assert parsed.document_id == "draft:test"
    assert parsed.metadata.title == "Test"
    # The Rust struct is #[serde(default)] throughout: every collection empty.
    assert parsed.document.paragraphs == ()
    assert parsed.document.sections == ()
    assert parsed.document.review_policy.track_changes is False


def test_parses_full_document() -> None:
    parsed = parse_envelope(
        envelope(
            document={
                "styles": ["Heading 1"],
                "sections": [{"title": "Summary", "paragraphs": [{"text": "hi", "style": "Normal"}]}],
                "paragraphs": [{"text": "intro"}],
                "lists": [{"items": ["a", "b"], "style": "List Bullet"}],
                "tables": [{"header": True, "rows": [["h1", "h2"], ["a", "b"]]}],
                "figures": [{"resource": "img", "caption": "cap"}],
                "content_controls": [{"tag": "shot", "value": "sh010"}],
                "headers": [{"section_index": 0, "text": "H"}],
                "footers": [{"section_index": 0, "text": "F"}],
                "fields": [{"kind": "toc"}, {"kind": "page"}],
                "review_policy": {"track_changes": True, "comments_locked": False},
            }
        )
    )
    document = parsed.document
    assert document.styles == ("Heading 1",)
    assert document.sections[0].title == "Summary"
    assert document.lists[0].items == ("a", "b")
    assert document.tables[0].rows[0] == ("h1", "h2")
    assert document.figures[0].resource == "img"
    assert document.content_controls[0].tag == "shot"
    assert document.headers[0].text == "H"
    assert len(document.fields) == 2
    assert document.review_policy.track_changes is True


def test_rejects_wrong_schema_version() -> None:
    with pytest.raises(IrValidationError, match="schema_version"):
        parse_envelope(envelope(schema_version="office-ir/9.9"))


def test_rejects_wrong_kind() -> None:
    with pytest.raises(IrValidationError, match="kind"):
        parse_envelope(envelope(kind="workbook"))


def test_rejects_unknown_field_kind() -> None:
    with pytest.raises(IrValidationError, match="unknown field kind"):
        parse_envelope(envelope(document={"fields": [{"kind": "nope"}]}))


def test_custom_field_requires_code() -> None:
    with pytest.raises(IrValidationError, match="requires 'code'"):
        parse_envelope(envelope(document={"fields": [{"kind": "custom"}]}))


def test_custom_field_with_code_is_accepted() -> None:
    parsed = parse_envelope(envelope(document={"fields": [{"kind": "custom", "code": "NUMPAGES"}]}))
    assert parsed.document.fields[0].code == "NUMPAGES"


def test_rejects_negative_section_index() -> None:
    with pytest.raises(IrValidationError, match="section_index"):
        parse_envelope(envelope(document={"headers": [{"section_index": -1, "text": "H"}]}))


def test_rejects_non_string_paragraph_text() -> None:
    with pytest.raises(IrValidationError, match="'text' must be a string"):
        parse_envelope(envelope(document={"paragraphs": [{"text": 42}]}))


def test_rejects_missing_document() -> None:
    # document_id is checked before document, so it must be present for the
    # missing-document error to be the one raised.
    with pytest.raises(IrValidationError, match="missing required key 'document'"):
        parse_envelope(
            {"schema_version": IR_VERSION, "kind": "word", "document_id": "d", "metadata": {"title": "t"}}
        )


def test_rejects_unknown_review_policy_key() -> None:
    with pytest.raises(IrValidationError, match="unknown key"):
        parse_envelope(envelope(document={"review_policy": {"nope": True}}))


def test_rejects_non_list_collection() -> None:
    with pytest.raises(IrValidationError, match="must be a list"):
        parse_envelope(envelope(document={"paragraphs": "not-a-list"}))


def test_artifact_stem_sanitizes_document_id() -> None:
    assert artifact_stem("draft:call sheet") == "draft-call-sheet"
    assert artifact_stem(":::") == "document"


def test_loads_from_file(tmp_path: Path) -> None:
    path = tmp_path / "doc.json"
    path.write_text(json.dumps(envelope()), encoding="utf-8")
    assert load_word_ir(path).document_id == "draft:test"


def test_load_rejects_missing_file(tmp_path: Path) -> None:
    with pytest.raises(IrValidationError, match="input file not found"):
        load_word_ir(tmp_path / "absent.json")


def test_load_rejects_invalid_json(tmp_path: Path) -> None:
    path = tmp_path / "bad.json"
    path.write_text("{not json", encoding="utf-8")
    with pytest.raises(IrValidationError, match="invalid JSON"):
        load_word_ir(path)
