"""Word IR contract — mirrors the dcc-mcp-office-ir word schema 1:1.

Contract-first: this module is the *domain core* of the Word adapter. The
compiler (headless Open XML implementation) and the COM reflow path both
consume this contract, never raw Word object-model calls.

The Rust authoritative schema lives in `crates/office-ir/src/lib.rs`
(`pub mod word`, proposal §13.3). JSON shape (snake_case, office-ir/1.0):

    {
      "schema_version": "office-ir/1.0",
      "kind": "word",
      "document_id": "draft:tech-report",
      "metadata": {"title": "Tech report", "author": "...", "language": "en"},
      "document": {
        "styles": ["Heading 1", "Normal"],
        "sections": [
          {"title": "Summary", "paragraphs": [{"text": "...", "style": "Normal"}]}
        ],
        "paragraphs": [],
        "lists": [{"items": ["a", "b"], "style": "List Bullet"}],
        "tables": [{"header": true, "rows": [["h1", "h2"], ["a", "b"]]}],
        "figures": [],
        "content_controls": [{"tag": "shot", "value": "sh010"}],
        "headers": [{"section_index": 0, "text": "..."}],
        "footers": [],
        "fields": [{"kind": "toc"}, {"kind": "page"}],
        "review_policy": {"track_changes": false, "comments_locked": false}
      },
      "outputs": ["docx"]
    }
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

IR_VERSION: str = "office-ir/1.0"
DOCUMENT_KIND: str = "word"

# Field kinds the IR accepts. `toc` and `page` are the ones the headless
# writer can materialise as real Word field instructions; `date` and `custom`
# are parsed and validated, and reported as host_limited because their
# displayed value is produced by Word on open, not by the writer.
FIELD_KINDS = frozenset({"toc", "page", "date", "custom"})

# Style names python-docx can always resolve in its default template. Kept in
# the contract module (not in `compiler`) so stdlib-only surfaces such as
# `validate` can warn about an unknown style without importing python-docx.
BUILTIN_STYLES = frozenset(
    {
        "Normal",
        "Title",
        "Heading 1",
        "Heading 2",
        "Heading 3",
        "Heading 4",
        "Heading 5",
        "Heading 6",
        "List Bullet",
        "List Number",
        "List Paragraph",
        "Caption",
        "Quote",
        "Intense Quote",
    }
)

DEFAULT_FONT = "Calibri"

# The Rust schema types `kind` as a free String. We pin the known set so a
# typo fails at the contract edge instead of silently writing no field.
REVIEW_POLICY_DEFAULTS: dict[str, bool] = {"track_changes": False, "comments_locked": False}


class IrValidationError(ValueError):
    """A Word IR document violated the contract. Carries a json-path hint."""

    def __init__(self, path: str, message: str) -> None:
        super().__init__(f"[{path}] {message}")
        self.path = path
        self.message = message


@dataclass(frozen=True)
class TemplateRef:
    uri: str
    version: str


@dataclass(frozen=True)
class Resource:
    id: str
    uri: str
    mime: str | None = None


@dataclass(frozen=True)
class Metadata:
    title: str
    author: str = ""
    language: str = "en"


@dataclass(frozen=True)
class Paragraph:
    text: str
    style: str | None = None


@dataclass(frozen=True)
class Section:
    paragraphs: tuple[Paragraph, ...] = ()
    title: str | None = None
    page_break_before: bool | None = None


@dataclass(frozen=True)
class ListBlock:
    items: tuple[str, ...] = ()
    style: str | None = None


@dataclass(frozen=True)
class TableBlock:
    header: bool
    rows: tuple[tuple[str, ...], ...] = ()


@dataclass(frozen=True)
class Figure:
    resource: str
    caption: str | None = None


@dataclass(frozen=True)
class ContentControl:
    tag: str
    value: str


@dataclass(frozen=True)
class HeaderFooterBlock:
    section_index: int
    text: str


@dataclass(frozen=True)
class FieldSpec:
    kind: str
    code: str | None = None


@dataclass(frozen=True)
class ReviewPolicy:
    track_changes: bool = False
    comments_locked: bool = False


@dataclass(frozen=True)
class WordDocumentIr:
    """Word document payload (proposal §13.3).

    The Rust struct is `#[serde(default)]` throughout, so every collection
    defaults to empty and `review_policy` to all-false. Mirrored here: an
    omitted key is legal, and only wrong *types* are errors.
    """

    styles: tuple[str, ...] = ()
    sections: tuple[Section, ...] = ()
    paragraphs: tuple[Paragraph, ...] = ()
    lists: tuple[ListBlock, ...] = ()
    tables: tuple[TableBlock, ...] = ()
    figures: tuple[Figure, ...] = ()
    content_controls: tuple[ContentControl, ...] = ()
    headers: tuple[HeaderFooterBlock, ...] = ()
    footers: tuple[HeaderFooterBlock, ...] = ()
    fields: tuple[FieldSpec, ...] = ()
    review_policy: ReviewPolicy = field(default_factory=ReviewPolicy)


@dataclass(frozen=True)
class WordEnvelope:
    schema_version: str
    kind: str
    document_id: str
    metadata: Metadata
    document: WordDocumentIr
    template: TemplateRef | None = None
    resources: tuple[Resource, ...] = ()
    outputs: tuple[str, ...] = ("docx",)


# Annotations defer under `from __future__ import annotations`, but these are
# runtime assignments used by helpers below — written as Optional[Union[...]]
# so the 3.9 floor does not raise TypeError on `str | None`.
_STR = Optional[Union[str]]


def _require(mapping: dict[str, Any], key: str, path: str) -> Any:
    if key not in mapping:
        raise IrValidationError(path, f"missing required key '{key}'")
    return mapping[key]


def _require_str(mapping: dict[str, Any], key: str, path: str) -> str:
    value = _require(mapping, key, path)
    if not isinstance(value, str):
        raise IrValidationError(path, f"'{key}' must be a string, got {type(value).__name__}")
    return value


def _opt_str(mapping: dict[str, Any], key: str, path: str) -> _STR:
    value = mapping.get(key)
    if value is None:
        return None
    if not isinstance(value, str):
        raise IrValidationError(path, f"'{key}' must be a string when present, got {type(value).__name__}")
    return value


def _opt_bool(mapping: dict[str, Any], key: str, path: str) -> bool | None:
    value = mapping.get(key)
    if value is None:
        return None
    if not isinstance(value, bool):
        raise IrValidationError(path, f"'{key}' must be a boolean when present, got {type(value).__name__}")
    return value


def _seq(raw: Any, path: str, key: str) -> list[Any]:
    """Return a list for `key`, tolerating omission like the Rust default."""
    if raw is None:
        return []
    if not isinstance(raw, list):
        raise IrValidationError(f"{path}.{key}", f"'{key}' must be a list")
    return raw


def _text(value: Any, path: str, key: str) -> str:
    if not isinstance(value, str):
        raise IrValidationError(path, f"'{key}' must be a string, got {type(value).__name__}")
    return value


def parse_paragraph(raw: Any, index: int, path: str) -> Paragraph:
    ipath = f"{path}[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "paragraph must be an object")
    return Paragraph(
        text=_text(_require(raw, "text", ipath), ipath, "text"),
        style=_opt_str(raw, "style", ipath),
    )


def parse_section(raw: Any, index: int) -> Section:
    ipath = f"document.sections[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "section must be an object")
    return Section(
        paragraphs=tuple(
            parse_paragraph(item, i, f"{ipath}.paragraphs") for i, item in enumerate(_seq(raw.get("paragraphs"), ipath, "paragraphs"))
        ),
        title=_opt_str(raw, "title", ipath),
        page_break_before=_opt_bool(raw, "page_break_before", ipath),
    )


def parse_list(raw: Any, index: int) -> ListBlock:
    ipath = f"document.lists[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "list must be an object")
    items = _seq(raw.get("items"), ipath, "items")
    return ListBlock(
        items=tuple(_text(item, f"{ipath}.items[{i}]", f"items[{i}]") for i, item in enumerate(items)),
        style=_opt_str(raw, "style", ipath),
    )


def parse_table(raw: Any, index: int) -> TableBlock:
    ipath = f"document.tables[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "table must be an object")
    rows_raw = _seq(raw.get("rows"), ipath, "rows")
    rows: list[tuple[str, ...]] = []
    for row_index, row in enumerate(rows_raw):
        rpath = f"{ipath}.rows[{row_index}]"
        if not isinstance(row, list):
            raise IrValidationError(rpath, "row must be a list of strings")
        rows.append(tuple(_text(cell, f"{rpath}[{c}]", f"rows[{row_index}][{c}]") for c, cell in enumerate(row)))
    return TableBlock(header=bool(raw.get("header", False)), rows=tuple(rows))


def parse_figure(raw: Any, index: int) -> Figure:
    ipath = f"document.figures[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "figure must be an object")
    return Figure(
        resource=_text(_require(raw, "resource", ipath), ipath, "resource"),
        caption=_opt_str(raw, "caption", ipath),
    )


def parse_content_control(raw: Any, index: int) -> ContentControl:
    ipath = f"document.content_controls[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "content control must be an object")
    return ContentControl(
        tag=_text(_require(raw, "tag", ipath), ipath, "tag"),
        value=_text(_require(raw, "value", ipath), ipath, "value"),
    )


def parse_header_footer(raw: Any, index: int, key: str) -> HeaderFooterBlock:
    ipath = f"document.{key}[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, f"{key} entry must be an object")
    section_index = raw.get("section_index", 0)
    if isinstance(section_index, bool) or not isinstance(section_index, int) or section_index < 0:
        raise IrValidationError(ipath, f"'section_index' must be a non-negative integer, got {section_index!r}")
    return HeaderFooterBlock(
        section_index=section_index,
        text=_text(_require(raw, "text", ipath), ipath, "text"),
    )


def parse_field(raw: Any, index: int) -> FieldSpec:
    ipath = f"document.fields[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(ipath, "field must be an object")
    kind = _text(_require(raw, "kind", ipath), ipath, "kind")
    if kind not in FIELD_KINDS:
        known = ", ".join(sorted(FIELD_KINDS))
        raise IrValidationError(ipath, f"unknown field kind '{kind}'; known: {known}")
    code = _opt_str(raw, "code", ipath)
    if kind == "custom" and not code:
        raise IrValidationError(ipath, "field kind 'custom' requires 'code'")
    return FieldSpec(kind=kind, code=code)


def parse_review_policy(raw: Any) -> ReviewPolicy:
    path = "document.review_policy"
    if raw is None:
        return ReviewPolicy()
    if not isinstance(raw, dict):
        raise IrValidationError(path, "must be an object")
    unknown = set(raw) - set(REVIEW_POLICY_DEFAULTS)
    if unknown:
        raise IrValidationError(path, f"unknown key(s): {sorted(unknown)}")
    return ReviewPolicy(
        track_changes=bool(raw.get("track_changes", False)),
        comments_locked=bool(raw.get("comments_locked", False)),
    )


def parse_word_document(raw: Any) -> WordDocumentIr:
    path = "document"
    if not isinstance(raw, dict):
        raise IrValidationError(path, "must be an object")
    styles_raw = _seq(raw.get("styles"), path, "styles")
    styles = tuple(_text(s, f"{path}.styles[{i}]", f"styles[{i}]") for i, s in enumerate(styles_raw))
    return WordDocumentIr(
        styles=styles,
        sections=tuple(parse_section(item, i) for i, item in enumerate(_seq(raw.get("sections"), path, "sections"))),
        paragraphs=tuple(
            parse_paragraph(item, i, f"{path}.paragraphs") for i, item in enumerate(_seq(raw.get("paragraphs"), path, "paragraphs"))
        ),
        lists=tuple(parse_list(item, i) for i, item in enumerate(_seq(raw.get("lists"), path, "lists"))),
        tables=tuple(parse_table(item, i) for i, item in enumerate(_seq(raw.get("tables"), path, "tables"))),
        figures=tuple(parse_figure(item, i) for i, item in enumerate(_seq(raw.get("figures"), path, "figures"))),
        content_controls=tuple(
            parse_content_control(item, i) for i, item in enumerate(_seq(raw.get("content_controls"), path, "content_controls"))
        ),
        headers=tuple(parse_header_footer(item, i, "headers") for i, item in enumerate(_seq(raw.get("headers"), path, "headers"))),
        footers=tuple(parse_header_footer(item, i, "footers") for i, item in enumerate(_seq(raw.get("footers"), path, "footers"))),
        fields=tuple(parse_field(item, i) for i, item in enumerate(_seq(raw.get("fields"), path, "fields"))),
        review_policy=parse_review_policy(raw.get("review_policy")),
    )


def parse_envelope(raw: Any) -> WordEnvelope:
    if not isinstance(raw, dict):
        raise IrValidationError("$", "envelope must be an object")
    version = _require_str(raw, "schema_version", "$")
    if version != IR_VERSION:
        raise IrValidationError("$.schema_version", f"expected '{IR_VERSION}', got '{version}'")
    kind = _require_str(raw, "kind", "$")
    if kind != DOCUMENT_KIND:
        raise IrValidationError("$.kind", f"expected '{DOCUMENT_KIND}', got '{kind}'")
    metadata_raw = _require(raw, "metadata", "$")
    if not isinstance(metadata_raw, dict):
        raise IrValidationError("$.metadata", "must be an object")
    template_raw = raw.get("template")
    template = None
    if template_raw is not None:
        if not isinstance(template_raw, dict):
            raise IrValidationError("$.template", "must be an object")
        template = TemplateRef(
            uri=_require_str(template_raw, "uri", "$.template"),
            version=str(template_raw.get("version", "0.0.0")),
        )
    return WordEnvelope(
        schema_version=version,
        kind=kind,
        document_id=_require_str(raw, "document_id", "$"),
        metadata=Metadata(
            title=_require_str(metadata_raw, "title", "$.metadata"),
            author=str(metadata_raw.get("author", "")),
            language=str(metadata_raw.get("language", "en")),
        ),
        template=template,
        resources=tuple(
            Resource(id=str(r.get("id", i)), uri=str(r.get("uri", ""))) for i, r in enumerate(raw.get("resources", []))
        ),
        document=parse_word_document(_require(raw, "document", "$")),
        outputs=tuple(str(o) for o in raw.get("outputs", ["docx"])),
    )


def artifact_stem(document_id: str) -> str:
    """Safe filesystem stem for a document id (ids may contain ':')."""
    return re.sub(r"[^A-Za-z0-9._-]+", "-", document_id).strip("-") or "document"


def load_word_ir(source: str | Path | dict[str, Any]) -> WordEnvelope:
    """Load and validate a Word IR document from a JSON file or a mapping."""
    if isinstance(source, dict):
        return parse_envelope(source)
    path = Path(source)
    if not path.is_file():
        raise IrValidationError("$", f"input file not found: {path}")
    try:
        with path.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
    except json.JSONDecodeError as exc:
        raise IrValidationError("$", f"invalid JSON: {exc}") from exc
    return parse_envelope(raw)
