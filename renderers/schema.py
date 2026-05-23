"""
Structured document schema — the single source of truth that feeds
all three renderers (HTML, DOCX, PDF/LaTeX).

The schema intentionally mirrors the editorial structure of an academic
manuscript (title, authors, abstract, headings A/B/C, tables, figures,
references) so the GADING ruleset maps cleanly onto every output.

We also accept the legacy "flat" shape produced by main.py's
FormattingEngine.get_doc_json() — see normalize_input().
"""

from __future__ import annotations

from typing import List, Optional, Literal, Union, Dict, Any
from pydantic import BaseModel, Field


# ---------- inline runs (styled text within a paragraph) ----------

class RunSpan(BaseModel):
    """A styled run of text inside a paragraph (think of an HTML <span>)."""
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    superscript: bool = False
    subscript: bool = False


# ---------- block-level elements ----------

class HeadingBlock(BaseModel):
    type: Literal["heading"] = "heading"
    level: Literal["A", "B", "C"] = "A"   # GADING heading hierarchy
    text: str
    numbering: Optional[str] = None       # e.g. "1.", "1.1", "1.1.1"


class ParagraphBlock(BaseModel):
    type: Literal["paragraph"] = "paragraph"
    text: Optional[str] = None            # plain text (used if runs is empty)
    runs: List[RunSpan] = Field(default_factory=list)
    style: Optional[str] = None           # e.g. "Main Text", "Abstract", "Caption B"
    alignment: Optional[Literal["left", "right", "center", "justify"]] = None


class TableBlock(BaseModel):
    type: Literal["table"] = "table"
    caption: Optional[str] = None
    header: Optional[List[str]] = None
    rows: List[List[str]] = Field(default_factory=list)


class FigureBlock(BaseModel):
    type: Literal["figure"] = "figure"
    src: Optional[str] = None             # filename or URL
    caption: Optional[str] = None
    alt: Optional[str] = None


class ReferenceBlock(BaseModel):
    type: Literal["reference"] = "reference"
    raw: str                              # full reference string, APA-formatted
    runs: List[RunSpan] = Field(default_factory=list)   # optional pre-parsed runs


class EquationBlock(BaseModel):
    type: Literal["equation"] = "equation"
    latex: Optional[str] = None
    text: Optional[str] = None
    number: Optional[str] = None


Block = Union[HeadingBlock, ParagraphBlock, TableBlock, FigureBlock, ReferenceBlock, EquationBlock]


# ---------- document container ----------

class DocumentMetadata(BaseModel):
    title: Optional[str] = None
    authors: List[str] = Field(default_factory=list)
    affiliations: List[str] = Field(default_factory=list)
    abstract: Optional[str] = None
    keywords: Optional[str] = None
    journal: Optional[str] = None
    volume: Optional[str] = None
    issue: Optional[str] = None
    year: Optional[str] = None
    doi: Optional[str] = None


class StructuredDocument(BaseModel):
    metadata: DocumentMetadata = Field(default_factory=DocumentMetadata)
    blocks: List[Block] = Field(default_factory=list)

    # Free-form extras (figures dir, equation numbers, etc.)
    extras: Dict[str, Any] = Field(default_factory=dict)


# ---------- input normalization ----------

def normalize_input(payload: Any) -> StructuredDocument:
    """
    Accept either:
      (a) a StructuredDocument dict ({metadata, blocks: [...]}) — preferred
      (b) the legacy flat list produced by FormattingEngine.get_doc_json()
          ([{text, style, alignment, props, css, isHtml?}, ...])

    Returns a StructuredDocument either way.
    """
    if isinstance(payload, StructuredDocument):
        return payload

    if isinstance(payload, dict) and "blocks" in payload:
        return StructuredDocument(**payload)

    if isinstance(payload, list):
        return _from_flat(payload)

    raise ValueError(
        "Unrecognized document payload — expected StructuredDocument dict or flat list."
    )


def _from_flat(items: List[dict]) -> StructuredDocument:
    """Convert the legacy flat get_doc_json() shape into a StructuredDocument."""
    meta = DocumentMetadata()
    blocks: List[Block] = []

    for item in items:
        style = (item.get("style") or "").strip()
        text = item.get("text") or ""
        alignment = (item.get("alignment") or "").lower() or None
        if alignment == "justify_low":          # python-docx alias
            alignment = "justify"

        # Front-matter consumed into metadata
        if style == "Title" and not meta.title:
            meta.title = text
            continue
        if style == "Author":
            meta.authors.append(text)
            continue
        if style == "Affiliation":
            meta.affiliations.append(text)
            continue
        if style == "Abstract":
            meta.abstract = (meta.abstract + " " if meta.abstract else "") + text
            continue
        if text.lower().startswith("keywords:"):
            meta.keywords = text.split(":", 1)[1].strip()
            continue

        # Headings
        if style in ("Heading A", "Heading B", "Heading C"):
            blocks.append(HeadingBlock(level=style.split()[-1], text=text))
            continue

        # Tables (legacy format embeds rows under props.table_data)
        if style == "Table Placeholder":
            props = item.get("props") or {}
            rows = props.get("table_data") or []
            blocks.append(TableBlock(rows=rows))
            continue

        # Figures
        if style == "Image Placeholder":
            blocks.append(FigureBlock(caption=item.get("caption")))
            continue

        # References
        if style == "Reference":
            blocks.append(ReferenceBlock(raw=text))
            continue

        # Default → paragraph
        blocks.append(ParagraphBlock(text=text, style=style or "Main Text", alignment=alignment))

    return StructuredDocument(metadata=meta, blocks=blocks)
