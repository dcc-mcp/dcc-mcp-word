---
name: word-document
description: >-
  Compile a Word IR JSON (office-ir/1.0, kind: word) into a native DOCX through
  the headless Open XML backend, then read the artifact back and verify every
  paragraph and table cell against the source IR. Use whenever the agent must
  produce a .docx from structured data — technical reports, call sheets,
  approval forms, shot summaries. No Word installation required.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: word
    layer: domain
    stage: authoring
    version: 0.1.0
    tags:
      - word
      - document
      - docx
      - generate
      - readback
      - openxml
    search-hint: >-
      generate docx, make word document, technical report, call sheet,
      approval form, compile document, read back document
    tools: tools.yaml
---

# word-document (Authoring stage)

Document generation through the designed pipeline:

1. data planner picks sections, paragraphs, lists and tables
2. Word IR document (contract: dcc-mcp-office-ir `pub mod word`, proposal §13.3)
3. headless Open XML compile → DOCX (no Word installation needed)
4. read-back verification: reopen the artifact and compare every paragraph
   and table cell against the source IR
5. structural validation report

The compile path is `verified`: it runs headless with no Word installed.
Reflow, field/TOC refresh and track-changes inspection are `host_limited`
and are reported, never faked — see `word-capabilities`.

## Related skills

- `word-capabilities` — the graded capability report for this adapter

## Input contract

- `input` — path to a Word IR JSON envelope
  (`schema_version: office-ir/1.0`, `kind: word`)
- `output_dir` — artifact directory
- `verify` — run the read-back verification step (default: true)

```json
{
  "schema_version": "office-ir/1.0",
  "kind": "word",
  "document_id": "draft:call-sheet",
  "metadata": {"title": "Call sheet", "author": "pipeline", "language": "en"},
  "document": {
    "sections": [
      {"title": "Tomorrow", "paragraphs": [{"text": "Call time 07:00", "style": "Normal"}]}
    ],
    "tables": [{"header": true, "rows": [["Shot", "Status"], ["sh010", "ip"]]}],
    "lists": [{"items": ["Bring rain cover", "Check generator"], "style": "List Bullet"}],
    "footers": [{"section_index": 0, "text": "Call sheet — page"}],
    "fields": [{"kind": "page"}]
  },
  "outputs": ["docx"]
}
```

## Decision rules

- model the document as sections and paragraphs first; never hand-place
  coordinates before the outline is settled
- use `sections[].title` for headings — it maps to `Heading 1` and anchors
  the TOC
- TOC and PAGE fields are written as real Word field instructions. Word
  fills in the results when the document is opened; say so in the hand-off
  rather than implying the page numbers already exist.
- `fields` with `kind: page` land in the footer automatically; a `toc` field
  is inserted at the top of the body.

## Scripts

- `generate_document` — IR → DOCX + read-back verification + validation report
- `validate_document` — validate a Word IR without generating
- `inspect_document` — read-only inventory of an existing DOCX

## Validation rules

- envelope contract enforced at load (`word_ir`): bad schema version, unknown
  field kinds, negative section indices and wrong value types are hard errors
  carrying a json-path hint
- read-back compares every paragraph (style + text) and every table cell; a
  mismatch fails the run instead of warning
- artifacts must exist and be non-empty

## Known limits

- python-docx has no layout engine: TOC entries and page numbers appear only
  after Word refreshes the fields (`word.toc.rebuild` / `word.fields.update`
  are `host_limited`)
- content controls and image figures are parsed and validated by the IR, but
  are not materialised by the headless writer in v0.1.0
- the compiler raises on specs it cannot write rather than silently dropping
  them
- non-built-in style names fall back to `Normal` unless the target template
  defines them — `validate_document` reports this as a warning

## Agent-visible summary

Result context: artifact path, paragraphs and tables written, read-back
verdict (checked counts, mismatches), validation checks + warnings, and the
capability grading for anything host-limited.
