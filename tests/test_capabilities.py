"""Capability grading tests — no host_limited capability may be graded verified.

This is the guard for PIP-4276: the whole point of grading is that the
adapter never claims a capability it cannot prove headlessly.
"""

from __future__ import annotations

from dcc_mcp_word.capabilities import (
    CAPABILITIES,
    HOST_LIMITED,
    UNIMPLEMENTED,
    VERIFIED,
    by_grade,
    get,
    report,
)

# Tools that need Word's layout engine or COM object model. They are named
# after the dcc-mcp-office tool registry so a rename upstream is caught here.
REQUIRES_OFFICE = {
    "word.fields.update",
    "word.toc.rebuild",
    "word.document.reflow",
    "word.track_changes.inspect",
}


def test_every_capability_has_a_known_grade() -> None:
    for capability in CAPABILITIES:
        assert capability.grade in {VERIFIED, HOST_LIMITED, UNIMPLEMENTED}


def test_host_limited_capabilities_require_office() -> None:
    for capability in by_grade(HOST_LIMITED):
        assert capability.requires_office is True, f"{capability.name} is host_limited but needs no Office"


def test_verified_capabilities_never_require_office() -> None:
    """The `verified` grade means provable headlessly — Office must not be needed."""
    for capability in by_grade(VERIFIED):
        assert capability.requires_office is False, f"{capability.name} is verified but requires Office"


def test_verified_capabilities_carry_evidence() -> None:
    for capability in CAPABILITIES:
        assert capability.evidence.strip(), f"{capability.name} has no evidence"


def test_no_host_limited_capability_is_graded_verified() -> None:
    verified = {c.name for c in by_grade(VERIFIED)}
    for name in REQUIRES_OFFICE:
        assert name not in verified, f"{name} needs Word's engine but is graded verified"


def test_reflow_is_host_limited() -> None:
    reflow = get("word.document.reflow")
    assert reflow is not None
    assert reflow.grade == HOST_LIMITED
    assert reflow.is_verified is False


def test_compile_is_verified() -> None:
    compile_cap = get("word.document.compile")
    assert compile_cap is not None
    assert compile_cap.grade == VERIFIED
    assert compile_cap.is_verified is True


def test_report_shape() -> None:
    body = report()
    assert body["schema"] == "dcc-mcp-capability-grade/1"
    assert "verified" in body and "host_limited" in body and "unimplemented" in body
    assert "word.document.compile" in body["verified"]
    assert "word.toc.rebuild" in body["host_limited"]


def test_report_is_json_serialisable() -> None:
    import json

    json.dumps(report())


def test_get_unknown_returns_none() -> None:
    assert get("word.does.not.exist") is None
