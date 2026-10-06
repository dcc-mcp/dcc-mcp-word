---
name: word-capabilities
description: >-
  Report which Word capabilities are verified and which require a desktop Word
  installation (host_limited), plus the host matrix and the start-up preflight
  self-check. Use before promising a Word capability to a user, and when a Word
  operation degrades or is unavailable.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: word
    layer: infrastructure
    stage: diagnostics
    version: 0.1.0
    tags:
      - word
      - capabilities
      - host-matrix
      - preflight
      - diagnostics
    search-hint: >-
      word capability, is it verified, host limited, do I need Word,
      preflight, host matrix, why is the table of contents empty
    tools: tools.yaml
---

# word-capabilities (Diagnostics stage)

The grade of a capability is evidence, not confidence:

| Grade | Meaning |
|---|---|
| `verified` | Covered by a CI-green headless test. No Word installation involved. |
| `host_limited` | Requires desktop Word (COM). The artifact is structurally valid without it, but the value only materializes when Word opens the file. |
| `unimplemented` | Not built. Reported instead of silently degrading. |

v0.1.0: document compile, read-back verification, paragraph / list / table /
header-footer / field-instruction writes are `verified`. Field updates, TOC
rebuild, document reflow and track-changes inspection are `host_limited`.
Content controls and image figures are `unimplemented` in the headless
writer (parsed and validated by the IR, not materialised).

## Why the split matters

`word.document.reflow` is graded `host_limited` even though the *compile*
half of the same pipeline is `verified`. python-docx writes a valid DOCX but
has no layout engine: it cannot re-paginate an existing document, and it
cannot compute what a TOC or PAGE field displays. Claiming the whole tool as
verified would overstate what CI actually proves.

## Do not over-promise

- Do not tell a user the table of contents or page numbers are in the file
  until Word has opened and refreshed it. The field instruction is written;
  Word produces the result.
- Do not describe content controls or figures as "filled" or "inserted" —
  the IR declares and validates them; the headless writer does not emit them
  in v0.1.0.
- A `preflight` failure is a report, not a crash: the headless surface stays
  usable without desktop Word.

## Scripts

- `capabilities` — print the graded capability report
- `preflight` — host matrix + start-up self-check (what this machine can do)

## After Failure

- `preflight ok=false` means the headless backend itself is missing
  (python-docx not installed) — install the package with its `dev` extra.
- `desktop_word=false` is expected on Linux/macOS and on CI: only the
  `host_limited` capabilities degrade.
