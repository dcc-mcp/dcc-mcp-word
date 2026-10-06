"""Headless Open XML compiler — Word IR → DOCX (ADR 004 / ADR 006).

Contract-first: the compiler is an *implementation* of the Word IR contract.
It writes structure — paragraphs, headings, lists, tables, headers/footers,
page fields and a TOC field. Two things it deliberately does not fake:

- **Field results.** A TOC or PAGE field is written as a real Word field
  instruction. Word computes the page numbers and the table of contents when
  the document is opened; python-docx has no layout engine, so any "result"
  we wrote would be a lie. `word.fields.update` / `word.toc.rebuild` are
  graded `host_limited` for exactly this reason.
- **Content controls and figures.** Anchored content controls and image
  figures are declared and validated by the IR, and reported as
  `host_limited` rather than dropped silently.

The dependency on python-docx is opt-in exactly like openpyxl in
dcc-mcp-excel: importing this package never pulls it, only this module does.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

from .word_ir import (
    DEFAULT_FONT,
    WordDocumentIr,
    WordEnvelope,
)

DOCX_SUFFIX = ".docx"


class UnsupportedFigureError(ValueError):
    """A figure references a resource the headless writer cannot inline."""


def _add_field(paragraph: Any, instruction: str) -> None:
    """Append a real Word field instruction to a paragraph.

    Word renders the result on open (or on F9). We write begin / instrText /
    separate / end so the field is genuine rather than literal text that only
    looks like one.
    """
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = _field_placeholder(instruction)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])


def _field_placeholder(instruction: str) -> str:
    """Cached result text shown before Word refreshes the field.

    A field with no cached result renders as an empty gap. `PAGE` is the one
    field whose absence is visibly wrong (a blank page number), so it gets a
    seed value; `TOC` legitimately shows "update fields" until refreshed.
    """
    stripped = instruction.strip().upper()
    if stripped == "PAGE":
        return "1"
    if stripped.startswith("TOC"):
        return "Right-click and choose 'Update Field' to build the table of contents."
    return ""


def _style_or_default(document: Any, style: str | None, default: str = "Normal") -> Any:
    """Resolve an IR style name against the document, falling back safely."""
    if not style:
        return document.styles[default]
    try:
        return document.styles[style]
    except KeyError:
        return document.styles[default]


def _add_paragraph(document: Any, paragraph_ir: Any) -> Any:
    style = _style_or_default(document, paragraph_ir.style)
    return document.add_paragraph(paragraph_ir.text, style=style.name)


def _write_paragraphs(document: Any, paragraphs: tuple[Any, ...]) -> int:
    for paragraph_ir in paragraphs:
        _add_paragraph(document, paragraph_ir)
    return len(paragraphs)


def _write_lists(document: Any, document_ir: WordDocumentIr) -> int:
    for list_block in document_ir.lists:
        style = _style_or_default(document, list_block.style, default="List Bullet")
        for item in list_block.items:
            document.add_paragraph(item, style=style.name)
    return sum(len(block.items) for block in document_ir.lists)


def _write_tables(document: Any, document_ir: WordDocumentIr) -> int:
    tables = 0
    for table_ir in document_ir.tables:
        if not table_ir.rows:
            continue
        column_count = max(len(row) for row in table_ir.rows)
        table = document.add_table(rows=0, cols=column_count)
        table.style = "Table Grid"
        for row_index, row in enumerate(table_ir.rows):
            cells = table.add_row().cells
            for column_index in range(column_count):
                value = row[column_index] if column_index < len(row) else ""
                cells[column_index].text = value
            if row_index == 0 and table_ir.header:
                for cell in cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True
        tables += 1
    return tables


def _write_fields(document: Any, document_ir: WordDocumentIr, *, where: Any) -> int:
    """Write field instructions into `where` (body paragraph or footer)."""
    count = 0
    for field in document_ir.fields:
        instruction = field.code or _default_instruction(field.kind)
        if instruction is None:
            continue
        paragraph = where if where is not None else document.add_paragraph()
        if field.kind == "page" and where is None:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_field(paragraph, instruction)
        count += 1
    return count


def _default_instruction(kind: str) -> str | None:
    if kind == "page":
        return "PAGE"
    if kind == "toc":
        return r'TOC \o "1-3" \h \z \u'
    if kind == "date":
        return "DATE"
    return None


def _write_toc_into_body(document: Any, document_ir: WordDocumentIr) -> None:
    """Insert a TOC field at the top of the body when the IR asks for one."""
    toc_fields = [f for f in document_ir.fields if f.kind == "toc"]
    if not toc_fields:
        return
    _write_fields(document, WordDocumentIr(fields=tuple(toc_fields)), where=None)


def _write_body_fields(document: Any, document_ir: WordDocumentIr) -> int:
    """Write the fields that belong in the body: date and custom codes.

    `toc` is handled by `_write_toc_into_body` (it must sit above the
    content it indexes) and `page` by `_write_page_fields` (a page number
    belongs in the footer), so neither is written here.
    """
    body_fields = tuple(f for f in document_ir.fields if f.kind in ("date", "custom"))
    if not body_fields:
        return 0
    return _write_fields(document, WordDocumentIr(fields=body_fields), where=None)


def _write_header_footer(document: Any, blocks: tuple[Any, ...], *, kind: str) -> int:
    """Write header/footer blocks onto their target section.

    Section index 0 is the document's only section for a freshly compiled
    document; later indices create the section they need.
    """
    written = 0
    for block in blocks:
        section = _section_for_index(document, block.section_index)
        target = section.header if kind == "header" else section.footer
        paragraph = target.paragraphs[0] if target.paragraphs else target.add_paragraph()
        if kind == "footer":
            paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
            left = paragraph.add_run(block.text)
            left.font.size = None
            separator = paragraph.add_run("\t")
            separator.font.size = None
        else:
            paragraph.text = block.text
        written += 1
    return written


def _section_for_index(document: Any, index: int) -> Any:
    """Return section `index`, creating the document sections to reach it."""
    while len(document.sections) <= index:
        document.add_section()
    return document.sections[index]


def _write_page_fields(document: Any, document_ir: WordDocumentIr) -> int:
    """PAGE fields belong in the footer; write them there, not in the body."""
    page_fields = tuple(f for f in document_ir.fields if f.kind == "page")
    if not page_fields:
        return 0
    section = _section_for_index(document, 0)
    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return _write_fields(document, WordDocumentIr(fields=page_fields), where=paragraph)


def _write_metadata(document: Any, envelope: WordEnvelope) -> None:
    properties = document.core_properties
    properties.title = envelope.metadata.title
    properties.author = envelope.metadata.author or "DCC-MCP"
    properties.subject = f"dcc-mcp-word document ({envelope.document_id})"
    properties.language = envelope.metadata.language


def _configure_base_styles(document: Any, document_ir: WordDocumentIr) -> None:
    """Touch the styles the IR declares so the names exist on the document.

    python-docx documents already ship the built-in names; this makes an IR
    that declares extra styles visible in `validate` as a warning rather than
    failing at compile time.
    """
    normal = document.styles["Normal"]
    if normal.font.name != DEFAULT_FONT:
        normal.font.name = DEFAULT_FONT


class WordCompiler:
    """Compiles a WordEnvelope into a DOCX file via python-docx."""

    def __init__(self, envelope: WordEnvelope, document: Any | None = None) -> None:
        self.envelope = envelope
        self.doc = document if document is not None else Document()
        self.summary: dict[str, Any] = {
            "paragraphs": 0,
            "sections": 0,
            "lists": 0,
            "tables": 0,
            "headers": 0,
            "footers": 0,
            "fields": 0,
        }

    def _apply_metadata(self) -> None:
        _write_metadata(self.doc, self.envelope)
        _configure_base_styles(self.doc, self.envelope.document)

    def compile(self, out_path: str | Path) -> Path:
        document_ir = self.envelope.document
        self._apply_metadata()

        # A TOC belongs above the content it indexes.
        _write_toc_into_body(self.doc, document_ir)

        self.summary["paragraphs"] += _write_paragraphs(self.doc, document_ir.paragraphs)
        self.summary["lists"] = _write_lists(self.doc, document_ir)
        self.summary["tables"] = _write_tables(self.doc, document_ir)

        for section_ir in document_ir.sections:
            if section_ir.page_break_before:
                self.doc.add_page_break()
            if section_ir.title is not None:
                heading = self.doc.add_paragraph(section_ir.title, style="Heading 1")
                heading.paragraph_format.keep_with_next = True
            self.summary["paragraphs"] += _write_paragraphs(self.doc, section_ir.paragraphs)
            self.summary["sections"] += 1

        self.summary["headers"] = _write_header_footer(self.doc, document_ir.headers, kind="header")
        self.summary["footers"] = _write_header_footer(self.doc, document_ir.footers, kind="footer")
        self.summary["fields"] = _write_page_fields(self.doc, document_ir) + _write_body_fields(self.doc, document_ir)

        out = Path(out_path)
        if out.suffix.lower() != DOCX_SUFFIX:
            out = out.with_suffix(DOCX_SUFFIX)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(out))
        return out


def compile_document(envelope: WordEnvelope, out_path: str | Path) -> Path:
    """Compile a WordEnvelope to DOCX via the headless Open XML backend."""
    return WordCompiler(envelope).compile(out_path)
