"""dcc-mcp-word — Word adapter for the DCC-MCP ecosystem.

Thin application layer over `dcc-mcp-office`. This package owns Word
semantics only: the Word IR contract, the headless Open XML compile +
read-back path, capability grading and the office-host launcher. The shared
machinery (protocol, IR envelope, C# COM runtime, Open XML worker) comes
from `dcc-mcp-office`.

Capability surface (v0.1.0):
- word_ir: the Word IR contract (mirrors dcc-mcp-office-ir `pub mod word`)
- compiler: Word IR -> DOCX (headless Open XML, python-docx; opt-in import)
- readback: write-then-read-back verification (opt-in import)
- capabilities: verified / host_limited grading (stdlib-only)
- host_matrix: host matrix + preflight self-check (stdlib-only)
- validate: structural validation reports (stdlib-only)
- host_client: stdlib-only JSON-RPC client for the shared office host

Grading boundary: the *compile* half is `verified` (headless, CI-green).
Reflow, field/TOC refresh and track-changes inspection are `host_limited` —
they need the desktop Word COM backend. Never reported as verified: see
`capabilities` for the per-tool evidence.

Dependency policy (mirrors dcc-mcp-excel): python-docx is opt-in. Importing
this package pulls neither it nor pywin32.
"""

from __future__ import annotations

from .capabilities import CAPABILITIES, Capability, report
from .host_matrix import HOST_MATRIX, preflight
from .validate import validate_artifacts, validate_envelope
from .word_ir import IrValidationError, WordEnvelope, load_word_ir

__version__ = "0.1.1"

__all__ = [
    "CAPABILITIES",
    "HOST_MATRIX",
    "Capability",
    "IrValidationError",
    "WordEnvelope",
    "__version__",
    "load_word_ir",
    "preflight",
    "report",
    "validate_artifacts",
    "validate_envelope",
]
