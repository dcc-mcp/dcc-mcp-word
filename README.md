# dcc-mcp-word

Word adapter for the DCC-MCP ecosystem — thin application layer over
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office).

**Status: planned — not started.** This repository is a placeholder created
as part of the Office Automation Platform repo split (see
`dcc-mcp-office/docs/adr/006-shared-office-core-split.md`). Work starts in
Phase 2, blocked on the `dcc-mcp-office` M1 (COM MVP) release.

## Scope (proposal §11.2)

- - `word.document.reflow` — reflow documents (styles, sections,
  content controls, headers/footers, fields/TOC)
- `word.fields.update` / `word.toc.rebuild` / `word.track_changes.inspect`
- technical reports via the `office-generate-technical-report` skill

## Upstream

- `dcc-mcp-office` — protocol, IR, C# runtime, Open XML worker, security
  policy, generic skills.
- `dcc-mcp-core` — gateway, jobs, artifacts, skills runtime, lifecycle.

## License

MIT
