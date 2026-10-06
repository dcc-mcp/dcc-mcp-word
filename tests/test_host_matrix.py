"""Host matrix + preflight tests (the 1.0 start-up self-check gate)."""

from __future__ import annotations

from pathlib import Path

from dcc_mcp_word import host_matrix
from dcc_mcp_word.capabilities import HOST_LIMITED, VERIFIED, by_grade


def test_matrix_covers_headless_and_desktop_hosts() -> None:
    statuses = {entry.status for entry in host_matrix.HOST_MATRIX}
    assert "verified" in statuses
    assert "host_limited" in statuses
    assert any(entry.app == "headless" for entry in host_matrix.HOST_MATRIX)
    assert any(entry.app == "Word" for entry in host_matrix.HOST_MATRIX)


def test_find_word_executable_returns_path_or_none() -> None:
    result = host_matrix.find_word_executable()
    assert result is None or isinstance(result, Path)


def test_headless_available_matches_import() -> None:
    try:
        import docx  # noqa: F401
    except ImportError:
        assert host_matrix.headless_available() is False
    else:
        assert host_matrix.headless_available() is True


def test_preflight_reports_without_raising(monkeypatch) -> None:
    monkeypatch.setattr(host_matrix, "find_word_executable", lambda: None)
    status = host_matrix.preflight()
    assert status["app"] == "word"
    assert status["checks"]["desktop_word"] is False
    assert status["degraded_without_office"] == [c.name for c in by_grade(HOST_LIMITED)]


def test_preflight_ok_depends_on_headless_backend(monkeypatch) -> None:
    monkeypatch.setattr(host_matrix, "headless_available", lambda: True)
    monkeypatch.setattr(host_matrix, "find_word_executable", lambda: None)
    assert host_matrix.preflight()["ok"] is True

    monkeypatch.setattr(host_matrix, "headless_available", lambda: False)
    # A missing headless backend is a real failure, not a degradation.
    assert host_matrix.preflight()["ok"] is False


def test_preflight_with_desktop_word_has_no_degradation(monkeypatch, tmp_path: Path) -> None:
    word = tmp_path / "WINWORD.EXE"
    word.write_bytes(b"")
    monkeypatch.setattr(host_matrix, "find_word_executable", lambda: word)
    status = host_matrix.preflight()
    assert status["degraded_without_office"] == []
    assert status["word_path"] == str(word)


def test_preflight_lists_grades() -> None:
    status = host_matrix.preflight()
    assert set(status["verified"]) == {c.name for c in by_grade(VERIFIED)}
    assert len(status["host_matrix"]) == len(host_matrix.HOST_MATRIX)
