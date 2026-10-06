"""office-rpc/1 host client tests — protocol boundary without a live host.

The client is stdlib-only and must report a missing host rather than raising
or silently falling back to a second local writer (ADR 006: one docx
implementation, in the shared core / the headless compiler).
"""

from __future__ import annotations

from pathlib import Path

from dcc_mcp_word import host_client
from dcc_mcp_word.host_client import (
    HOST_EXE,
    handshake,
    ping,
    rebuild_toc,
    rpc,
    update_fields,
)

OFFICE_HOST_NOT_FOUND = "OFFICE_HOST_NOT_FOUND"


def test_find_host_binary_prefers_env(monkeypatch, tmp_path: Path) -> None:
    fake = tmp_path / HOST_EXE
    fake.write_bytes(b"")
    monkeypatch.setenv("DCC_OFFICE_HOST", str(fake))
    assert host_client.find_host_binary() == str(fake)


def test_find_host_binary_env_ignored_when_not_a_file(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("DCC_OFFICE_HOST", str(tmp_path / "absent.exe"))
    monkeypatch.setattr(host_client.shutil, "which", lambda _name: None)
    assert host_client.find_host_binary() is None


def test_rpc_reports_missing_host_without_raising(monkeypatch) -> None:
    monkeypatch.setattr("dcc_mcp_word.host_client.find_host_binary", lambda: None)
    result = rpc("office.host.ping", {})
    assert result["success"] is False
    assert OFFICE_HOST_NOT_FOUND in result["reason"]
    assert result["backend"] is None


def test_client_helpers_report_missing_host(monkeypatch) -> None:
    monkeypatch.setattr("dcc_mcp_word.host_client.find_host_binary", lambda: None)
    for result in (ping(), handshake(), update_fields("a.docx"), rebuild_toc("a.docx")):
        assert result["success"] is False
        assert OFFICE_HOST_NOT_FOUND in result["reason"]


def test_matching_response_selects_by_request_id() -> None:
    stdout = (
        '{"jsonrpc": "2.0", "id": "req", "result": {"ok": true}}\n'
        '{"jsonrpc": "2.0", "id": "other", "result": {"ok": false}}\n'
    )
    assert host_client._matching_response(stdout, "req")["result"] == {"ok": True}


def test_matching_response_rejects_missing_id() -> None:
    try:
        host_client._matching_response('{"jsonrpc": "2.0", "id": "nope"}', "req")
    except ValueError as exc:
        assert "did not contain response id" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("expected ValueError")


def test_matching_response_rejects_invalid_json() -> None:
    try:
        host_client._matching_response("not json", "req")
    except ValueError as exc:
        assert "invalid JSON on host output line 1" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("expected ValueError")


def test_matching_response_rejects_duplicate_ids() -> None:
    stdout = '{"id": "req"}\n{"id": "req"}\n'
    try:
        host_client._matching_response(stdout, "req")
    except ValueError as exc:
        assert "contained 2 responses" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("expected ValueError")


def test_abs_normalizes_relative_paths(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.chdir(tmp_path)
    assert host_client._abs("doc.docx") == str((tmp_path / "doc.docx").resolve())


def test_rpc_work_dir_creates_a_directory() -> None:
    work = host_client._rpc_work_dir()
    try:
        assert work.is_dir()
    finally:
        import shutil

        shutil.rmtree(work, ignore_errors=True)
