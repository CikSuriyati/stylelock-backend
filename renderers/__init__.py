"""
StyleLock unified rendering pipeline.

The same structured document JSON + ruleset can be rendered to:
  - HTML/CSS (web preview)
  - PDF (via LaTeX, with WeasyPrint fallback)
  - DOCX (via python-docx, programmatic Word styles)

All three renderers consume the same StructuredDocument model defined in
schema.py and the same ruleset JSON (e.g. gading_ruleset.json).
"""

from .schema import StructuredDocument, Block, RunSpan, TableBlock, FigureBlock, ReferenceBlock, HeadingBlock, ParagraphBlock, normalize_input
from .html_renderer import render_html
from .docx_renderer import render_docx
from .latex_renderer import render_latex, render_pdf

__all__ = [
    "StructuredDocument",
    "Block",
    "RunSpan",
    "TableBlock",
    "FigureBlock",
    "ReferenceBlock",
    "HeadingBlock",
    "ParagraphBlock",
    "normalize_input",
    "render_html",
    "render_docx",
    "render_latex",
    "render_pdf",
]
