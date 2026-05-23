"""
HTML/CSS renderer — WeasyPrint-optimised for B5 JIS MJCET/GADING layout.

v2 improvements over the original:
  - B5 JIS page size (182 × 257 mm) + twips-derived margins from ruleset
  - Tinos font (Google Fonts CDN — Times New Roman metric substitute)
  - CSS @page margin boxes: italic running header + DOI/copyright footer
  - 9 pt body text with 0.236 in first-line indent (340 twips)
  - 17 pt title / 13 pt authors / 8 pt italic affiliations
  - Article-info table: history | gutter | abstract/keywords (MJCET layout)
  - Heading A: 10 pt bold UPPERCASE  /  B: bold title-case  /  C: italic
  - Tables: 8 pt, horizontal-only borders (top / below_header / bottom)
  - References: 9 pt, APA 7 0.5 in hanging indent
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

_TWIPS_PER_INCH = 1440


def _t2in(twips: float) -> float:
    return twips / _TWIPS_PER_INCH


def _t2mm(twips: float) -> float:
    return twips / _TWIPS_PER_INCH * 25.4


def _t2pt(twips: float) -> float:
    return twips / _TWIPS_PER_INCH * 72


# ─────────────────────────────────────────────────────────────────────────────
#  CSS generation
# ─────────────────────────────────────────────────────────────────────────────

def _build_css(ruleset: Dict[str, Any], header_line: str, footer_line: str) -> str:
    fonts        = ruleset.get("fonts", {})
    main_fi      = fonts.get("main_text", {})
    table_fi     = fonts.get("table_text", {})
    spacing      = ruleset.get("spacing", {})
    headings     = ruleset.get("headings", {})
    tables_cfg   = ruleset.get("tables", {})
    refs_cfg     = ruleset.get("references", {})
    title_block  = ruleset.get("title_block", {})
    ps           = ruleset.get("page_setup", {})
    margins      = ps.get("margins_twips", {})

    main_family = main_fi.get("family", "Times New Roman")
    main_size   = main_fi.get("size_pt", 9)
    table_size  = table_fi.get("size_pt", 8)

    # Line height — convert auto spacing (twips/240 = multiple)
    ls_twips = spacing.get("main_text_line_spacing_twips", 276)
    ls_rule  = spacing.get("main_text_line_spacing_rule", "auto")
    if ls_rule == "exact":
        line_height = f"{_t2pt(ls_twips):.1f}pt"
    else:
        line_height = f"{ls_twips / 240:.3f}"

    # First-line indent
    fi_in = _t2in(spacing.get("main_text_first_line_indent_twips", 340))

    # Page margins
    m_top    = _t2mm(margins.get("top",    754))
    m_bottom = _t2mm(margins.get("bottom", 1418))
    m_left   = _t2mm(margins.get("left",   1191))
    m_right  = _t2mm(margins.get("right",  1191))
    m_header = _t2mm(margins.get("header", 907))
    m_footer = _t2mm(margins.get("footer", 1202))

    # Ensure the margin is large enough to fit the header/footer band
    top_margin    = max(m_top + 10, m_header + 6)
    bottom_margin = max(m_bottom + 8, m_footer + 6)

    # Headings helper
    def h_style(key: str, default_size: int) -> str:
        h   = headings.get(key, {}) or {}
        sz  = h.get("font_size_pt", default_size)
        fw  = "bold"   if h.get("bold")   else "normal"
        fs  = "italic" if h.get("italic") else "normal"
        align = h.get("alignment", "left")
        case  = h.get("case", "")
        tt = {"upper": "uppercase", "title_case": "capitalize", "sentence_case": "none"}.get(case, "none")
        sp_b = _t2pt(h.get("spacing_before_twips", 360))
        sp_a = _t2pt(h.get("spacing_after_twips",  240))
        return (
            f"font-family: 'Tinos', '{main_family}', serif; font-size: {sz}pt; "
            f"font-weight: {fw}; font-style: {fs}; text-align: {align}; "
            f"text-transform: {tt}; margin-top: {sp_b:.1f}pt; margin-bottom: {sp_a:.1f}pt; "
            f"margin-left: 0; padding: 0;"
        )

    # Table borders
    borders      = tables_cfg.get("borders", {})
    horiz        = borders.get("horizontal", ["top", "below_header", "bottom"])
    border_pt    = borders.get("border_size_halfpts", 4) / 2.0

    # References
    hang_in  = refs_cfg.get("hanging_indent_in", 0.5)
    ref_size = refs_cfg.get("font_size_pt", main_size)
    ref_sp_a = _t2pt(refs_cfg.get("spacing_after_twips", 180))

    # Title block
    tb_title  = title_block.get("title",  {})
    tb_author = title_block.get("author", {})
    tb_affil  = title_block.get("affiliation", {})
    title_size  = tb_title.get("font_size_pt",  17)
    author_size = tb_author.get("font_size_pt", 13)
    affil_size  = tb_affil.get("font_size_pt",   8)
    title_bold  = "bold" if tb_title.get("bold", False) else "normal"

    # Escape header/footer strings for CSS content property
    def css_str(s: str) -> str:
        return s.replace("\\", "\\\\").replace('"', '\\"')

    h_esc = css_str(header_line)
    f_esc = css_str(footer_line)

    return f"""
@import url("https://fonts.googleapis.com/css2?family=Tinos:ital,wght@0,400;0,700;1,400;1,700&display=swap");

/* ── Page geometry ───────────────────────────────────────────── */
@page {{
  size: 182mm 257mm;
  margin: {top_margin:.1f}mm {m_right:.1f}mm {bottom_margin:.1f}mm {m_left:.1f}mm;

  @top-left {{
    content: "{h_esc}";
    font-family: 'Tinos', '{main_family}', serif;
    font-size: 8pt;
    font-style: italic;
    vertical-align: bottom;
    padding-bottom: 1.5mm;
    border-bottom: 0.5pt solid #000;
    width: 100%;
  }}

  @bottom-left {{
    content: "{f_esc}";
    font-family: 'Tinos', '{main_family}', serif;
    font-size: 7pt;
    vertical-align: top;
    padding-top: 1.5mm;
    border-top: 0.5pt solid #000;
    width: 80%;
  }}

  @bottom-right {{
    content: counter(page);
    font-family: 'Tinos', '{main_family}', serif;
    font-size: 8pt;
    vertical-align: top;
    padding-top: 1.5mm;
    text-align: right;
  }}
}}

/* First page: suppress running header (it has its own layout) */
@page :first {{
  @top-left {{ content: ""; border-bottom: none; }}
}}

/* ── Base ────────────────────────────────────────────────────── */
*, *::before, *::after {{ box-sizing: border-box; }}

body {{
  font-family: 'Tinos', '{main_family}', serif;
  font-size: {main_size}pt;
  line-height: {line_height};
  color: #000;
  margin: 0;
  padding: 0;
  background: #fff;
}}

/* ── Journal banner (first page only) ───────────────────────── */
.journal-banner {{
  font-size: 8pt;
  font-style: italic;
  border-bottom: 0.5pt solid #000;
  padding-bottom: 2pt;
  margin-bottom: 10pt;
  color: #222;
}}

/* ── Title block ─────────────────────────────────────────────── */
.meta-title {{
  font-size: {title_size}pt;
  font-weight: {title_bold};
  text-align: center;
  line-height: 1.25;
  margin: 0 0 8pt;
}}

.meta-authors {{
  font-size: {author_size}pt;
  text-align: center;
  margin: 0 0 4pt;
}}

.meta-affiliation {{
  font-size: {affil_size}pt;
  font-style: italic;
  text-align: center;
  margin: 0 0 3pt;
}}

/* ── Article-info table ──────────────────────────────────────── */
.article-info-table {{
  width: 100%;
  border-collapse: collapse;
  border-top: 2pt solid #000;
  border-bottom: 2pt solid #000;
  margin: 10pt 0 14pt;
  font-size: {main_size}pt;
  line-height: {line_height};
}}

.article-info-table td {{
  vertical-align: top;
  padding: 4pt 0;
  border: none;
}}

.article-info-left  {{ width: 32%; padding-right: 6pt; }}
.article-info-gutter {{ width: 2%; border-left: 0.5pt solid #bbb; }}
.article-info-right {{ width: 66%; padding-left: 8pt; }}

.info-section {{ margin-bottom: 6pt; }}
.info-label {{ font-weight: bold; font-size: {main_size}pt; display: block; margin-bottom: 1pt; }}
.info-date-row {{ display: block; font-size: {max(main_size - 1, 7)}pt; }}
.info-doi {{ font-size: {max(main_size - 1, 7)}pt; word-break: break-all; }}
.abstract-label {{ font-weight: bold; font-style: italic; }}

/* ── Headings ────────────────────────────────────────────────── */
h2.heading-A {{ {h_style("Heading A", main_size)} }}
h3.heading-B {{ {h_style("Heading B", main_size)} }}
h4.heading-C {{ {h_style("Heading C", main_size)} }}

/* ── Body paragraphs ─────────────────────────────────────────── */
p.main-text {{
  text-align: justify;
  text-indent: {fi_in:.4f}in;
  margin: 0;
  line-height: {line_height};
}}

p.no-indent {{
  text-indent: 0;
}}

/* ── Tables ──────────────────────────────────────────────────── */
table.data-table {{
  width: 100%;
  border-collapse: collapse;
  font-size: {table_size}pt;
  margin: 4pt 0 2pt;
  line-height: 1.2;
  page-break-inside: avoid;
}}

table.data-table th,
table.data-table td {{
  padding: 2pt 4pt;
  border: none;
  text-align: left;
}}

table.data-table thead tr th {{
  {"border-top: " + str(border_pt) + "pt solid #000;" if "top" in horiz else ""}
  {"border-bottom: " + str(border_pt) + "pt solid #000;" if "below_header" in horiz else ""}
  font-weight: bold;
}}

table.data-table tbody tr:last-child td {{
  {"border-bottom: " + str(border_pt) + "pt solid #000;" if "bottom" in horiz else ""}
}}

p.table-caption {{
  font-size: {table_size}pt;
  font-weight: bold;
  text-align: left;
  margin: 6pt 0 2pt;
  text-indent: 0;
}}

/* ── Figures ──────────────────────────────────────────────────── */
figure.figure-block {{
  text-align: center;
  margin: 10pt 0;
  page-break-inside: avoid;
}}

figure.figure-block img {{ max-width: 100%; height: auto; }}

figure.figure-block figcaption {{
  font-size: {main_size}pt;
  text-align: left;
  margin-top: 3pt;
}}

/* ── References ───────────────────────────────────────────────── */
p.reference-entry {{
  font-size: {ref_size}pt;
  text-align: justify;
  padding-left: {hang_in}in;
  text-indent: -{hang_in}in;
  margin: 0 0 {ref_sp_a:.1f}pt;
  line-height: {line_height};
}}

/* ── Equations ────────────────────────────────────────────────── */
p.equation {{
  text-align: right;
  font-style: italic;
  margin: 6pt 0;
  text-indent: 0;
}}
""".strip()


# ─────────────────────────────────────────────────────────────────────────────
#  Inline runs
# ─────────────────────────────────────────────────────────────────────────────

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


def _para_inner(text: str | None, runs: List[RunSpan]) -> str:
    if runs:
        return _runs_to_html(runs)
    return escape(text or "")


# ─────────────────────────────────────────────────────────────────────────────
#  Block renderers
# ─────────────────────────────────────────────────────────────────────────────

def _render_heading(b: HeadingBlock) -> str:
    tag   = {"A": "h2", "B": "h3", "C": "h4"}[b.level]
    klass = f"heading-{b.level}"
    prefix = f"{escape(b.numbering)} " if b.numbering else ""
    return f'<{tag} class="{klass}">{prefix}{escape(b.text)}</{tag}>'


def _render_paragraph(b: ParagraphBlock) -> str:
    align_attr = f' style="text-align: {b.alignment};"' if b.alignment else ""
    return f'<p class="main-text"{align_attr}>{_para_inner(b.text, b.runs)}</p>'


def _render_table(b: TableBlock) -> str:
    parts = ['<table class="data-table">']
    if b.header:
        parts.append("<thead><tr>")
        parts.extend(f"<th>{escape(str(c))}</th>" for c in b.header)
        parts.append("</tr></thead>")
    if b.rows:
        parts.append("<tbody>")
        for row in b.rows:
            parts.append("<tr>")
            parts.extend(f"<td>{escape(str(c))}</td>" for c in row)
            parts.append("</tr>")
        parts.append("</tbody>")
    parts.append("</table>")
    if b.caption:
        parts.append(f'<p class="table-caption">{escape(b.caption)}</p>')
    return "\n".join(parts)


def _render_figure(b: FigureBlock) -> str:
    src = escape(b.src or "")
    alt = escape(b.alt or b.caption or "Figure")
    cap = f"<figcaption>{escape(b.caption)}</figcaption>" if b.caption else ""
    if src:
        img = f'<img src="{src}" alt="{alt}" />'
    else:
        img = f'<div style="padding:18pt 0;border:1px dashed #aaa;font-size:8pt;color:#666;">[Figure: {alt}]</div>'
    return f'<figure class="figure-block">{img}{cap}</figure>'


def _render_reference(b: ReferenceBlock) -> str:
    inner = _runs_to_html(b.runs) if b.runs else escape(b.raw)
    return f'<p class="reference-entry">{inner}</p>'


def _render_equation(b: EquationBlock) -> str:
    body = escape(b.text or b.latex or "")
    num  = f' ({escape(b.number)})' if b.number else ""
    return f'<p class="equation">{body}{num}</p>'


_DISPATCH = {
    "heading":   _render_heading,
    "paragraph": _render_paragraph,
    "table":     _render_table,
    "figure":    _render_figure,
    "reference": _render_reference,
    "equation":  _render_equation,
}


# ─────────────────────────────────────────────────────────────────────────────
#  Article-info table (left: history/keywords/doi | gutter | right: abstract)
# ─────────────────────────────────────────────────────────────────────────────

def _build_article_info_table(md, ruleset: Dict[str, Any]) -> str:
    # ── Left column ──────────────────────────────────────────────
    left_html = ""

    # Article history (fields not in schema — skip gracefully)
    history_dates = []
    for attr, label in [("received","Received"), ("revised","Revised"),
                         ("accepted","Accepted"), ("published","Published")]:
        val = getattr(md, attr, None)
        if val:
            history_dates.append((label, val))

    if history_dates:
        rows = "".join(f'<span class="info-date-row">{label}: {escape(str(d))}</span>'
                       for label, d in history_dates)
        left_html += f'<div class="info-section"><span class="info-label">Article History</span>{rows}</div>'

    if md.keywords:
        kws = escape(md.keywords)
        left_html += f'<div class="info-section"><span class="info-label">Keywords</span>{kws}</div>'

    doi = getattr(md, "doi", None)
    if doi:
        left_html += f'<div class="info-section"><span class="info-label">DOI</span><span class="info-doi">{escape(doi)}</span></div>'

    # ── Right column ─────────────────────────────────────────────
    right_html = ""
    if md.abstract:
        right_html = f'<span class="abstract-label">Abstract</span> {escape(md.abstract)}'

    # Only emit the table if there's content on at least one side
    if not left_html and not right_html:
        return ""

    # If left is empty, make right span full width
    if not left_html:
        return f'<div style="border-top:2pt solid #000;border-bottom:2pt solid #000;margin:10pt 0 14pt;padding:4pt 0;"><span class="abstract-label">Abstract</span> {escape(md.abstract or "")}</div>'

    return f"""<table class="article-info-table">
  <tr>
    <td class="article-info-left">{left_html}</td>
    <td class="article-info-gutter"></td>
    <td class="article-info-right">{right_html}</td>
  </tr>
</table>"""


# ─────────────────────────────────────────────────────────────────────────────
#  Entrypoint
# ─────────────────────────────────────────────────────────────────────────────

def render_html(document: Any, ruleset: Dict[str, Any], *, standalone: bool = True) -> str:
    """
    Render a StructuredDocument (or compatible payload) to HTML.

    Args:
        document: StructuredDocument or dict/flat-list normalised via normalize_input().
        ruleset:  parsed ruleset JSON (mjcet / uitm / gading).
        standalone: True → full <html> document; False → bare <div> fragment.
    """
    doc = normalize_input(document)
    md  = doc.metadata

    # ── Build running header text ─────────────────────────────────
    journal_info = ruleset.get("journal", {})
    journal_name = journal_info.get("name", "")
    hf_cfg       = ruleset.get("header_footer", {})
    header_tmpl  = hf_cfg.get("header", {}).get("odd_page", "")
    footer_tmpl  = hf_cfg.get("footer", {}).get("odd_page", "")

    # First author surname
    first_author = "Author"
    if md.authors:
        parts = md.authors[0].split()
        first_author = parts[-1] if parts else md.authors[0]

    year   = str(md.year)   if md.year   else "Year"
    volume = str(md.volume) if md.volume else "X"
    issue  = str(md.issue)  if md.issue  else "X"

    if header_tmpl:
        header_line = (header_tmpl
            .replace("First Author", first_author)
            .replace("(Year)", f"({year})")
            .replace("Vol. X", f"Vol. {volume}")
            .replace("No. X",  f"No. {issue}"))
    else:
        short = journal_info.get("short_name", journal_name or "Journal")
        header_line = f"{first_author} et al. / {short} ({year}) Vol. {volume}, No. {issue}"

    # Footer
    doi_val = getattr(md, "doi", None) or journal_info.get("doi_pattern", "")
    copyright_text = journal_info.get("copyright_text", "©UiTM Press")
    if footer_tmpl:
        footer_line = footer_tmpl
    else:
        footer_line = f"{doi_val}   {copyright_text}" if doi_val else copyright_text

    css = _build_css(ruleset, header_line, footer_line)

    # ── Build body ────────────────────────────────────────────────
    body_parts: List[str] = []

    # First-page journal banner (mimics the Word first-page header)
    banner_journal = journal_name or journal_info.get("short_name", "")
    if banner_journal:
        body_parts.append(f'<div class="journal-banner">{escape(banner_journal)}</div>')

    # Title block
    if md.title:
        body_parts.append(f'<div class="meta-title">{escape(md.title)}</div>')
    for author in md.authors:
        body_parts.append(f'<div class="meta-authors">{escape(author)}</div>')
    for aff in md.affiliations:
        body_parts.append(f'<div class="meta-affiliation">{escape(aff)}</div>')

    # Article-info table (abstract + keywords + history)
    info_table = _build_article_info_table(md, ruleset)
    if info_table:
        body_parts.append(info_table)

    # Body blocks
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
<body>
{body}
</body>
</html>"""
