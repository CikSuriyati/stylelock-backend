"""
HTML/CSS renderer.

Produces a self-contained HTML document whose CSS is derived from the
ruleset (fonts, sizes, alignment, heading rules, table borders, hanging
indents for references). The same @page rules are emitted so a browser
print preview matches the LaTeX/DOCX outputs.

The web preview imports this HTML directly — no second styling pass.
"""

from __future__ import annotations

from html import escape
from typing import Dict, Any, List

from .schema import (
    StructuredDocument,
    HeadingBlock,
    ParagraphBlock,
    TableBlock,
    FigureBlock,
    ReferenceBlock,
    EquationBlock,
    RunSpan,
    normalize_input,
)


# ---------- ruleset -> CSS ----------

def _css_for_ruleset(ruleset: Dict[str, Any]) -> str:
    fonts = ruleset.get("fonts", {})
    main = fonts.get("main_text", {})
    table_font = fonts.get("table_text", {})
    spacing = ruleset.get("spacing", {})
    headings = ruleset.get("headings", {})
    tables = ruleset.get("tables", {})
    refs = ruleset.get("references", {})

    main_family = main.get("family", "Times New Roman")
    main_size = main.get("size_pt", 10)
    line_spacing = spacing.get("main_text_line_spacing", "single")
    line_height = "1.0" if line_spacing == "single" else ("1.5" if line_spacing == "1.5" else "2.0" if line_spacing == "double" else "1.0")
    para_gap = spacing.get("paragraph_spacing_before_after_pt", 0)

    table_family = table_font.get("family", main_family)
    table_size = table_font.get("size_pt", 8)

    # heading rules
    def heading_css(level_key: str, default_size: int) -> str:
        h = headings.get(level_key, {}) or {}
        size = h.get("font_size_pt", default_size)
        weight = "bold" if h.get("bold") else "normal"
        style = "italic" if h.get("italic") else "normal"
        align = h.get("alignment", "left")
        case_rule = h.get("case")
        transform = ""
        if case_rule == "upper":
            transform = "text-transform: uppercase;"
        elif case_rule == "title_case":
            transform = "text-transform: capitalize;"
        return (
            f"font-family: '{main_family}', serif; "
            f"font-size: {size}pt; font-weight: {weight}; font-style: {style}; "
            f"text-align: {align}; {transform}"
        )

    # table borders
    borders = tables.get("borders", {}) or {}
    has_vertical = borders.get("vertical", False)
    horizontal = borders.get("horizontal", ["top", "below_header", "bottom"]) or []

    # APA hanging indent
    hanging_in = refs.get("hanging_indent_in", 0.5)
    hanging_em = hanging_in   # 0.5in ≈ 0.5em-ish; we use em so it scales

    css = f"""
:root {{
  --main-font: '{main_family}', serif;
  --main-size: {main_size}pt;
  --line-height: {line_height};
}}

@page {{
  size: A4;
  margin: 1in;
}}

body.stylelock-doc {{
  font-family: var(--main-font);
  font-size: var(--main-size);
  line-height: var(--line-height);
  color: #111;
  max-width: 6.5in;
  margin: 1in auto;
  padding: 0 0.25in;
  background: #fff;
}}

.stylelock-doc .meta-title {{
  text-align: center;
  font-weight: bold;
  font-size: {main_size + 2}pt;
  margin-bottom: 12pt;
}}

.stylelock-doc .meta-authors {{
  text-align: center;
  margin-bottom: 4pt;
}}

.stylelock-doc .meta-affiliation {{
  text-align: center;
  font-style: italic;
  margin-bottom: 12pt;
}}

.stylelock-doc .meta-abstract {{
  text-align: justify;
  font-style: italic;
  margin: 12pt 0;
}}

.stylelock-doc .meta-keywords {{
  margin-bottom: 18pt;
}}

.stylelock-doc h2.heading-A {{ {heading_css("Heading A", main_size)} margin: 12pt 0 6pt; }}
.stylelock-doc h3.heading-B {{ {heading_css("Heading B", main_size)} margin: 10pt 0 4pt; font-weight: bold; }}
.stylelock-doc h4.heading-C {{ {heading_css("Heading C", main_size)} margin: 8pt 0 4pt; font-style: italic; }}

.stylelock-doc p.main-text {{
  text-align: justify;
  margin: 0 0 {para_gap}pt;
  text-indent: 0;
}}

.stylelock-doc table.data-table {{
  width: 100%;
  border-collapse: collapse;
  font-family: '{table_family}', serif;
  font-size: {table_size}pt;
  margin: 8pt 0;
}}

.stylelock-doc table.data-table th,
.stylelock-doc table.data-table td {{
  padding: 4pt 6pt;
  border: none;
  {"border-left: 1px solid #111; border-right: 1px solid #111;" if has_vertical else ""}
}}

.stylelock-doc table.data-table thead tr {{
  {"border-top: 1px solid #111;" if "top" in horizontal else ""}
  {"border-bottom: 1px solid #111;" if "below_header" in horizontal else ""}
}}

.stylelock-doc table.data-table tbody tr:last-child {{
  {"border-bottom: 1px solid #111;" if "bottom" in horizontal else ""}
}}

.stylelock-doc figure.figure-block {{
  text-align: center;
  margin: 12pt 0;
}}

.stylelock-doc figure.figure-block img {{ max-width: 100%; }}

.stylelock-doc figure.figure-block figcaption {{
  font-size: {max(table_size, 9)}pt;
  margin-top: 4pt;
  text-align: center;
}}

.stylelock-doc p.reference-entry {{
  text-align: left;
  padding-left: {hanging_em}in;
  text-indent: -{hanging_em}in;
  margin-bottom: 4pt;
}}

.stylelock-doc .equation {{
  text-align: right;
  font-style: italic;
  margin: 8pt 0;
}}
"""
    return css.strip()


# ---------- inline runs ----------

def _runs_to_html(runs: List[RunSpan]) -> str:
    out = []
    for r in runs:
        t = escape(r.text)
        if r.bold:
            t = f"<strong>{t}</strong>"
        if r.italic:
            t = f"<em>{t}</em>"
        if r.underline:
            t = f"<u>{t}</u>"
        if r.superscript:
            t = f"<sup>{t}</sup>"
        if r.subscript:
            t = f"<sub>{t}</sub>"
        out.append(t)
    return "".join(out)


def _paragraph_inner(text: str | None, runs: List[RunSpan]) -> str:
    if runs:
        return _runs_to_html(runs)
    return escape(text or "")


# ---------- block renderers ----------

def _render_heading(b: HeadingBlock) -> str:
    tag = {"A": "h2", "B": "h3", "C": "h4"}[b.level]
    klass = f"heading-{b.level}"
    prefix = f"{escape(b.numbering)} " if b.numbering else ""
    return f'<{tag} class="{klass}">{prefix}{escape(b.text)}</{tag}>'


def _render_paragraph(b: ParagraphBlock) -> str:
    align_attr = f' style="text-align: {b.alignment};"' if b.alignment else ""
    klass = "main-text"
    return f'<p class="{klass}"{align_attr}>{_paragraph_inner(b.text, b.runs)}</p>'


def _render_table(b: TableBlock) -> str:
    parts = ['<table class="data-table">']
    if b.header:
        parts.append("<thead><tr>")
        parts.extend(f"<th>{escape(c)}</th>" for c in b.header)
        parts.append("</tr></thead>")
    parts.append("<tbody>")
    for row in b.rows:
        parts.append("<tr>")
        parts.extend(f"<td>{escape(str(c))}</td>" for c in row)
        parts.append("</tr>")
    parts.append("</tbody></table>")
    if b.caption:
        parts.append(f'<p class="main-text" style="text-align:center;font-size:9pt;">{escape(b.caption)}</p>')
    return "".join(parts)


def _render_figure(b: FigureBlock) -> str:
    src = escape(b.src or "")
    alt = escape(b.alt or b.caption or "Figure")
    cap = f"<figcaption>{escape(b.caption)}</figcaption>" if b.caption else ""
    img = f'<img src="{src}" alt="{alt}" />' if src else f'<div style="padding:24pt;border:1px dashed #aaa;">[Figure: {alt}]</div>'
    return f'<figure class="figure-block">{img}{cap}</figure>'


def _render_reference(b: ReferenceBlock) -> str:
    inner = _runs_to_html(b.runs) if b.runs else escape(b.raw)
    return f'<p class="reference-entry">{inner}</p>'


def _render_equation(b: EquationBlock) -> str:
    body = escape(b.text or b.latex or "")
    num = f' <span class="eq-num">({escape(b.number)})</span>' if b.number else ""
    return f'<p class="equation">{body}{num}</p>'


_DISPATCH = {
    "heading": _render_heading,
    "paragraph": _render_paragraph,
    "table": _render_table,
    "figure": _render_figure,
    "reference": _render_reference,
    "equation": _render_equation,
}


# ---------- entrypoint ----------

def render_html(document: Any, ruleset: Dict[str, Any], *, standalone: bool = True) -> str:
    """
    Render a StructuredDocument (or compatible payload) to HTML.

    Args:
        document: StructuredDocument or dict/flat-list normalized via normalize_input().
        ruleset: parsed gading_ruleset.json (or compatible).
        standalone: if True, returns a full <html> document with <style> embedded;
                    if False, returns just the <body> fragment (for embedding in a page).
    """
    doc = normalize_input(document)
    css = _css_for_ruleset(ruleset)

    body_parts: List[str] = []

    # Front matter
    md = doc.metadata
    if md.title:
        body_parts.append(f'<div class="meta-title">{escape(md.title)}</div>')
    for author in md.authors:
        body_parts.append(f'<div class="meta-authors">{escape(author)}</div>')
    for aff in md.affiliations:
        body_parts.append(f'<div class="meta-affiliation">{escape(aff)}</div>')
    if md.abstract:
        body_parts.append(f'<div class="meta-abstract"><strong>Abstract.</strong> {escape(md.abstract)}</div>')
    if md.keywords:
        body_parts.append(f'<div class="meta-keywords"><em>Keywords:</em> {escape(md.keywords)}</div>')

    # Blocks
    for b in doc.blocks:
        renderer = _DISPATCH.get(b.type)
        if renderer:
            body_parts.append(renderer(b))

    body = "\n".join(body_parts)

    if not standalone:
        return f'<div class="stylelock-doc">{body}</div>'

    title = escape(md.title or "Document")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>{title}</title>
<style>{css}</style>
</head>
<body class="stylelock-doc">
{body}
</body>
</html>"""
