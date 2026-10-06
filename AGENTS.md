# AGENTS.md — dcc-mcp-word

> Progressive disclosure: this file is a **map**, not an encyclopedia.

## 30-Second Summary

`dcc-mcp-word` is the **thin Word adapter** over `dcc-mcp-office`. It owns
Word application semantics only: the Word IR contract, the headless DOCX
compile + read-back path, capability grading and the office-host launcher.
Shared machinery (protocol, IR envelope, C# COM runtime, jobs, security
policy) comes from `dcc-mcp-office` + `dcc-mcp-core`.

**Current status:** v0.1.0 — headless path live. Word IR → Open XML compile →
read-back verification runs end to end with **no Word installation**, and the
gate runs headless in CI. Field/TOC refresh, document reflow and
track-changes inspection are `host_limited`: declared, graded, and never
claimed as verified.

## Repo Map

| Path | What it is |
|---|---|
| `src/dcc_mcp_word/word_ir.py` | Word IR contract (mirrors dcc-mcp-office-ir `pub mod word`) |
| `src/dcc_mcp_word/compiler.py` | headless Open XML compiler: Word IR → DOCX (python-docx) |
| `src/dcc_mcp_word/readback.py` | write-then-read-back verification (the 1.0 gate) |
| `src/dcc_mcp_word/capabilities.py` | verified / host_limited / unimplemented grading |
| `src/dcc_mcp_word/host_matrix.py` | host matrix + preflight self-check |
| `src/dcc_mcp_word/validate.py` | structural validation reports (stdlib-only) |
| `src/dcc_mcp_word/host_client.py` | stdlib-only JSON-RPC client for the shared office host |
| `src/dcc_mcp_word/server_launcher.py` | registers the bundled skills with `dcc-mcp-server` |
| `src/dcc_mcp_word/skills/word-document/` | SKILL.md + tools.yaml + scripts (generate/validate/inspect) |
| `src/dcc_mcp_word/skills/word-capabilities/` | capability report + preflight scripts |
| `examples/` | call-sheet Word IR + generated DOCX + COM verification |
| `tests/` | pytest (headless; COM only via subprocess boundary) |

## Upstream Dependencies

- `dcc-mcp-core` (pip) — gateway, skills runtime, sidecar lifecycle.
- `dcc-mcp-office` — Rust crates (`dcc-mcp-office-protocol`,
  `dcc-mcp-office-ir`, `dcc-mcp-office-tools`) + the `office-host` runtime.

## Capabilities owned here

- `word.document.compile` — Word IR → DOCX (headless) → read-back verification.
- `word.document.read_back` — reopen the artifact and compare every paragraph
  and table cell.
- `word.fields.update` / `word.toc.rebuild` / `word.document.reflow` /
  `word.track_changes.inspect` — `host_limited` in v0.1.0 (COM path not wired).

## Dependency policy

- **Package import is stdlib-only.** python-docx is **opt-in**: `compiler` /
  `readback` import it at module level, but nothing else in the package does.
  Two tests pin this (`tests/test_runtime_purity.py`).
- Keep the style constants in `word_ir.py`, not `compiler.py`, so stdlib-only
  surfaces such as `validate` can warn about unknown styles without importing
  python-docx. Moving them back breaks the purity tests.
- One docx implementation only (ADR 004 / ADR 006): no second writer inside
  the adapter, even as a fallback. A missing shared host is reported, not
  worked around.

## Grading rules

A capability is `verified` only if a headless CI test proves it. Anything
needing Word's layout engine or COM object model is `host_limited` — the
distinction is evidence, not confidence. `tests/test_capabilities.py` asserts
that no `host_limited` capability is graded `verified`.

## Test

```bash
pip install -e ".[dev]"
pytest
ruff check src tests
```

The headless gate must stay green **without Word installed** — that is the
proof the adapter does not depend on a desktop Word installation.

## Gotchas

- python-docx exposes `_Header`/`_Footer` objects with **no `.text`**
  attribute; join `part.paragraphs` instead. Hit twice already.
- `add_page_break()` emits an extra empty body paragraph. The read-back skips
  formatting-only paragraphs (page breaks and the TOC field paragraph) so
  counts line up with the IR.
