"""
LaTeX renderer + PDF compiler.

Emits a .tex document derived from the ruleset (geometry, font sizes,
heading commands, APA-style hanging-indent references). The PDF is
compiled by trying engines in this order:

  1. tectonic   (preferred — single binary, auto-downloads packages)
  2. xelatex / pdflatex  (if a system TeX is installed)
  3. WeasyPrint fallback (HTML → PDF, no LaTeX required)

If none of those are available, render_pdf raises RuntimeError with a
clear setup hint.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
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


# ---------- LaTeX escaping ----------

_LATEX_ESCAPES = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def _esc(s: str) -> str:
    if not s:
        return ""
    out = []
    for ch in s:
        out.append(_LATEX_ESCAPES.get(ch, ch))
    return "".join(out)


def _runs_to_tex(runs: List[RunSpan]) -> str:
    out = []
    for r in runs:
        t = _esc(r.text)
        if r.bold:
            t = r"\textbf{" + t + "}"
        if r.italic:
            t = r"\textit{" + t + "}"
        if r.underline:
            t = r"\underline{" + t + "}"
        if r.superscript:
            t = r"\textsuperscript{" + t + "}"
        if r.subscript:
            t = r"\textsubscript{" + t + "}"
        out.append(t)
    return "".join(out)


def _para_body(text: str | None, runs: List[RunSpan]) -> str:
    if runs:
        return _runs_to_tex(runs)
    return _esc(text or "")


# ---------- preamble from ruleset ----------

def _preamble(ruleset: Dict[str, Any]) -> str:
    fonts = ruleset.get("fonts", {}) or {}
    main = fonts.get("main_text", {}) or {}
    spacing = ruleset.get("spacing", {}) or {}
    refs = ruleset.get("references", {}) or {}

    main_size = main.get("size_pt", 10)
    line_spacing = spacing.get("main_text_line_spacing", "single")
    setspace = {"single": r"\singlespacing", "1.5": r"\onehalfspacing", "double": r"\doublespacing"}.get(line_spacing, r"\singlespacing")

    hanging_in = refs.get("hanging_indent_in", 0.5)

    # Times-like font; the user's ruleset asks for Times New Roman. We use
    # the `newtxtext`/`newtxmath` Times-equivalent which works under
    # pdflatex without needing TTF files. Tectonic will fetch it on demand.
    return rf"""\documentclass[{main_size}pt,a4paper]{{article}}
\usepackage[utf8]{{inputenc}}
\usepackage[T1]{{fontenc}}
\usepackage{{newtxtext,newtxmath}}
\usepackage[a4paper,margin=1in]{{geometry}}
\usepackage{{setspace}}
\usepackage{{titlesec}}
\usepackage{{enumitem}}
\usepackage{{booktabs}}
\usepackage{{array}}
\usepackage{{graphicx}}
\usepackage{{caption}}
\usepackage{{ragged2e}}
\usepackage[hidelinks]{{hyperref}}

% Spacing
{setspace}
\setlength{{\parindent}}{{0pt}}
\setlength{{\parskip}}{{0pt}}

% Headings (GADING A/B/C)
\titleformat{{\section}}{{\normalfont\bfseries}}{{}}{{0pt}}{{}}
\titleformat{{\subsection}}{{\normalfont\bfseries}}{{}}{{0pt}}{{}}
\titleformat{{\subsubsection}}{{\normalfont\itshape}}{{}}{{0pt}}{{}}
\titlespacing*{{\section}}{{0pt}}{{12pt}}{{6pt}}
\titlespacing*{{\subsection}}{{0pt}}{{10pt}}{{4pt}}
\titlespacing*{{\subsubsection}}{{0pt}}{{8pt}}{{4pt}}

% APA-style hanging indent for references
\newenvironment{{apareferences}}
  {{\begin{{list}}{{}}{{\setlength{{\leftmargin}}{{{hanging_in}in}}\setlength{{\itemindent}}{{-{hanging_in}in}}\setlength{{\itemsep}}{{4pt}}\setlength{{\parsep}}{{0pt}}}}}}
  {{\end{{list}}}}

% Justified main text
\AtBeginDocument{{\justifying}}
"""


# ---------- block renderers ----------

def _render_heading(b: HeadingBlock) -> str:
    cmd = {"A": r"\section*", "B": r"\subsection*", "C": r"\subsubsection*"}[b.level]
    prefix = f"{_esc(b.numbering)} " if b.numbering else ""
    return f"{cmd}{{{prefix}{_esc(b.text)}}}"


def _render_paragraph(b: ParagraphBlock) -> str:
    body = _para_body(b.text, b.runs)
    if b.alignment == "center":
        return r"\begin{center}" + body + r"\end{center}"
    if b.alignment == "right":
        return r"\begin{flushright}" + body + r"\end{flushright}"
    if b.alignment == "left":
        return r"\begin{flushleft}" + body + r"\end{flushleft}"
    return body + "\n"


def _render_table(b: TableBlock) -> str:
    if not b.rows and not b.header:
        return ""
    ncols = len(b.header) if b.header else len(b.rows[0])
    colspec = " ".join(["l"] * ncols)
    lines = [r"\begin{table}[h]", r"\centering", rf"\begin{{tabular}}{{{colspec}}}", r"\toprule"]
    if b.header:
        lines.append(" & ".join(_esc(c) for c in b.header) + r" \\")
        lines.append(r"\midrule")
    for row in b.rows:
        lines.append(" & ".join(_esc(str(c)) for c in row) + r" \\")
    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    if b.caption:
        lines.append(rf"\caption*{{{_esc(b.caption)}}}")
    lines.append(r"\end{table}")
    return "\n".join(lines)


def _render_figure(b: FigureBlock) -> str:
    if not b.src:
        # placeholder
        return r"\begin{center}\fbox{\parbox{0.5\linewidth}{\centering [Figure: " + _esc(b.alt or b.caption or "") + r"]}}\end{center}"
    lines = [r"\begin{figure}[h]", r"\centering", rf"\includegraphics[max width=\linewidth]{{{b.src}}}"]
    if b.caption:
        lines.append(rf"\caption*{{{_esc(b.caption)}}}")
    lines.append(r"\end{figure}")
    return "\n".join(lines)


def _render_reference(b: ReferenceBlock) -> str:
    inner = _runs_to_tex(b.runs) if b.runs else _esc(b.raw)
    return r"\item " + inner


def _render_equation(b: EquationBlock) -> str:
    body = b.latex if b.latex else _esc(b.text or "")
    if b.number:
        return rf"\begin{{flushright}}${body}$ \quad ({_esc(b.number)})\end{{flushright}}"
    return rf"\begin{{flushright}}${body}$\end{{flushright}}"


# ---------- entrypoint ----------

def render_latex(document: Any, ruleset: Dict[str, Any]) -> str:
    """Render a StructuredDocument to a complete LaTeX source string."""
    doc = normalize_input(document)
    parts: List[str] = [_preamble(ruleset), r"\begin{document}"]

    md = doc.metadata
    if md.title:
        parts.append(r"\begin{center}\bfseries\large " + _esc(md.title) + r"\end{center}")
    for author in md.authors:
        parts.append(r"\begin{center}" + _esc(author) + r"\end{center}")
    for aff in md.affiliations:
        parts.append(r"\begin{center}\textit{" + _esc(aff) + r"}\end{center}")
    if md.abstract:
        parts.append(r"\noindent\textit{Abstract.} " + _esc(md.abstract) + r"\par")
    if md.keywords:
        parts.append(r"\noindent\textit{Keywords:} " + _esc(md.keywords) + r"\par\vspace{6pt}")

    # Group consecutive references into a single hanging-indent env
    i = 0
    blocks = doc.blocks
    while i < len(blocks):
        b = blocks[i]
        if b.type == "reference":
            parts.append(r"\begin{apareferences}")
            while i < len(blocks) and blocks[i].type == "reference":
                parts.append(_render_reference(blocks[i]))
                i += 1
            parts.append(r"\end{apareferences}")
            continue

        if b.type == "heading":
            parts.append(_render_heading(b))
        elif b.type == "paragraph":
            parts.append(_render_paragraph(b))
        elif b.type == "table":
            parts.append(_render_table(b))
        elif b.type == "figure":
            parts.append(_render_figure(b))
        elif b.type == "equation":
            parts.append(_render_equation(b))
        i += 1

    parts.append(r"\end{document}")
    return "\n\n".join(parts)


# ---------- PDF compilation ----------

def _which(*names: str) -> str | None:
    for n in names:
        path = shutil.which(n)
        if path:
            return path
    return None


def _compile_with_engine(tex: str, engine: str, workdir: str) -> str:
    """Compile a .tex string with the given engine binary. Returns PDF path."""
    tex_path = os.path.join(workdir, "doc.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex)

    if engine.endswith("tectonic"):
        cmd = [engine, "--keep-logs", "--outdir", workdir, tex_path]
    else:
        cmd = [engine, "-interaction=nonstopmode", "-output-directory", workdir, tex_path]

    # Run twice for cross-refs (skip for tectonic which handles passes itself)
    passes = 1 if engine.endswith("tectonic") else 2
    for _ in range(passes):
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=workdir)
        if result.returncode != 0:
            log_tail = (result.stdout or "")[-2000:] + "\n---STDERR---\n" + (result.stderr or "")[-1000:]
            raise RuntimeError(f"LaTeX engine {engine} failed:\n{log_tail}")

    pdf_path = os.path.join(workdir, "doc.pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"LaTeX engine {engine} completed but produced no PDF.")
    return pdf_path


def _try_weasyprint(html: str, output_path: str) -> bool:
    try:
        from weasyprint import HTML  # type: ignore
    except Exception:
        return False
    HTML(string=html).write_pdf(output_path)
    return True


def render_pdf(document: Any, ruleset: Dict[str, Any], output_path: str) -> str:
    """
    Render a StructuredDocument to PDF at output_path.

    Tries (in order): tectonic, xelatex, pdflatex, WeasyPrint(html fallback).
    Returns the absolute output_path on success.
    """
    # Try LaTeX engines first; fall through to WeasyPrint if compile fails
    engine = _which("tectonic", "xelatex", "pdflatex")
    if engine:
        try:
            tex = render_latex(document, ruleset)
            with tempfile.TemporaryDirectory() as workdir:
                pdf_path = _compile_with_engine(tex, engine, workdir)
                shutil.copyfile(pdf_path, output_path)
            return output_path
        except Exception:
            pass  # engine found but failed (e.g. missing packages) — try WeasyPrint

    # WeasyPrint fallback (HTML -> PDF)
    from .html_renderer import render_html
    html = render_html(document, ruleset, standalone=True)
    if _try_weasyprint(html, output_path):
        return output_path

    raise RuntimeError(
        "No PDF engine available. Install one of: tectonic (recommended), "
        "TeX Live (xelatex/pdflatex), or `pip install weasyprint`."
    )
