# dcc-mcp-word

Word adapter for the DCC-MCP ecosystem — thin application layer over
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office).

**Status: v0.1.0 — headless compile path live.** Word IR → Open XML compile →
read-back verification runs end to end with **no Word installation**, and the
gate runs headless in CI. Field/TOC refresh, document reflow and
track-changes inspection are `host_limited`: declared, graded, and never
claimed as verified.

> The old blocker note ("blocked on the `dcc-mcp-office` M1 release") is
> **stale and removed**. `dcc-mcp-office` shipped M1 in v0.2.0 (2026-08-16)
> and is now at v0.2.3; the Word IR (`crates/office-ir`, `pub mod word`,
> proposal §13.3) and the `WordBackend.cs` COM backend have been present
> since then. The three placeholder repos stalled on an unreviewed blocker
> note, not on a technical wall.

## What v0.1.0 does

| Capability | Grade | Meaning |
|---|---|---|
| `word.document.compile` | `verified` | Word IR → DOCX, headless, CI-green |
| `word.document.read_back` | `verified` | Reopen the artifact; compare every paragraph and table cell |
| `word.paragraphs.write`, `word.lists.write`, `word.tables.write` | `verified` | Structured content writes |
| `word.headers_footers.write`, `word.fields.write` | `verified` | Per-section headers/footers; real Word field instructions |
| `word.fields.update`, `word.toc.rebuild` | `host_limited` | Word computes field results — needs the desktop app |
| `word.document.reflow`, `word.track_changes.inspect` | `host_limited` | Needs the Word COM object model |
| `word.content_controls.fill`, `word.figures.inline` | `unimplemented` | Parsed and validated by the IR, not emitted by the writer |

Run `dcc-mcp-word capabilities` for the machine-readable report, or
`dcc-mcp-word preflight` for the per-machine self-check.

## Install

```bash
pip install dcc-mcp-word          # adapter only (stdlib + dcc-mcp-core)
pip install "dcc-mcp-word[headless]"   # + python-docx, for the compile path
```

`python-docx` is an **opt-in extra**, matching how `dcc-mcp-excel` treats
`openpyxl`: importing the package never pulls it, so the adapter stays
installable inside DCCs that pin their own version. A test enforces this
(`tests/test_runtime_purity.py`).

## Quick start

```bash
# 1. a Word IR envelope (office-ir/1.0, kind: word)
cat > call-sheet.json <<'JSON'
{
  "schema_version": "office-ir/1.0",
  "kind": "word",
  "document_id": "fpt:call-sheet-2026-10-07",
  "metadata": {"title": "Call sheet — 2026-10-07", "author": "pipeline"},
  "document": {
    "sections": [{"title": "Schedule", "paragraphs": [{"text": "Call 07:00"}]}],
    "tables": [{"header": true, "rows": [["Shot", "Status"], ["sh010", "ip"]]}],
    "footers": [{"section_index": 0, "text": "Page"}],
    "fields": [{"kind": "toc"}, {"kind": "page"}]
  }
}
JSON

# 2. compile + read back
python -m dcc_mcp_word.skills.word-document.scripts.generate_document \
  --input call-sheet.json --out output
```

The script prints JSON: artifact path, read-back verdict (paragraphs and
cells checked, mismatches), validation checks and warnings.

## Input contract

The adapter mirrors the Rust `dcc-mcp-office-ir` word schema 1:1 — styles,
sections, paragraphs, lists, tables, figures, content controls,
headers/footers, fields and review policy. See
[`src/dcc_mcp_word/word_ir.py`](src/dcc_mcp_word/word_ir.py) for the
annotated contract, or `word-document/SKILL.md` for the agent-facing guide.

## The honest boundary

python-docx writes a valid DOCX but has **no layout engine**. That is why the
two halves of `word.document.reflow` are graded separately:

- *Compiling* (IR → DOCX) is headless and `verified`.
- *Reflowing, re-paginating, and refreshing field results* needs Word, and is
  `host_limited`.

A `TOC` field is written as a real field instruction. Word fills in the
entries when the document is opened — verified on a Word 365 host: before
`Fields.Update()` the TOC shows its placeholder, after it shows
`Schedule→1, Safety→2`. See
[`examples/com-verification.json`](examples/com-verification.json).

## Test

```bash
pip install -e ".[dev]"
pytest
ruff check src tests
```

The headless gate must stay green **without Word installed** — that is the
proof the adapter does not depend on a desktop Word installation.

## Upstream

- `dcc-mcp-office` — protocol, IR, C# COM runtime, jobs, security policy
- `dcc-mcp-core` — gateway, skills runtime, sidecar lifecycle

## License

MIT
