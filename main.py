from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional, Any
import shutil
import os
import uuid
import json as _json

from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="StyleLock Formatting System")

# CORS: in production set FRONTEND_URL env var to your Vercel URL.
# Defaults to "*" for local dev.
_frontend_url = os.environ.get("FRONTEND_URL", "*")
_allowed_origins = ["*"] if _frontend_url == "*" else [o.strip() for o in _frontend_url.split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
PROCESSED_DIR = os.path.join(BASE_DIR, "processed")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
DEFAULT_TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "gading_template.docx")
os.makedirs(TEMPLATES_DIR, exist_ok=True)

from renderers import render_html, render_docx, render_pdf
from ingest import ingest_docx

# =====================================================================
# Ruleset registry — name -> filename in this directory.
# Add a new journal/template by dropping a JSON file here and
# registering it below.
# =====================================================================

RULESET_REGISTRY = {
    "gading": "gading_ruleset.json",
    "mjcet": "mjcet_ruleset.json",
    "uitm": "uitm_ruleset.json",
}
DEFAULT_RULESET = "mjcet"

# Back-compat: kept so any legacy code path referencing RULESET_PATH still works.
RULESET_PATH = os.path.join(BASE_DIR, RULESET_REGISTRY[DEFAULT_RULESET])


def _load_ruleset(ruleset_name: Optional[str] = None) -> dict:
    """Load a ruleset by name. Defaults to the MJCET ruleset (the authoritative
    forensic-faithful template). Valid names: 'gading', 'mjcet', 'uitm'.
    Unknown names return a 400 so callers see the typo instead of silently
    falling back."""
    name = (ruleset_name or DEFAULT_RULESET).strip().lower()
    if name not in RULESET_REGISTRY:
        raise HTTPException(
            status_code=400,
            detail=f"Unknown ruleset '{ruleset_name}'. Available: {sorted(RULESET_REGISTRY.keys())}",
        )
    path = os.path.join(BASE_DIR, RULESET_REGISTRY[name])
    if not os.path.exists(path):
        raise HTTPException(status_code=500, detail=f"Ruleset file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return _json.load(f)


class RenderRequest(BaseModel):
    document: Any                      # StructuredDocument dict OR legacy flat list
    ruleset_name: Optional[str] = None # e.g. "mjcet" — defaults to mjcet_ruleset.json
    use_template: bool = True          # for DOCX: open UiTM template as base


# =====================================================================
# Render endpoints
# =====================================================================

@app.post("/render/html", response_class=HTMLResponse)
async def render_to_html(req: RenderRequest):
    """Render structured doc -> self-contained HTML for the web preview."""
    try:
        ruleset = _load_ruleset(req.ruleset_name)
        html = render_html(req.document, ruleset, standalone=True)
        return HTMLResponse(content=html)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"HTML render failed: {e}")


@app.post("/render/docx")
async def render_to_docx(req: RenderRequest):
    """Render structured doc -> .docx, applying Word styles from the ruleset."""
    try:
        ruleset = _load_ruleset(req.ruleset_name)
        job_id = str(uuid.uuid4())
        out_path = os.path.join(PROCESSED_DIR, f"{job_id}_render.docx")
        template = DEFAULT_TEMPLATE_PATH if req.use_template and os.path.exists(DEFAULT_TEMPLATE_PATH) else None
        render_docx(req.document, ruleset, out_path, template_path=template)
        return FileResponse(
            out_path,
            filename="document.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"DOCX render failed: {e}")


@app.post("/render/pdf")
async def render_to_pdf(req: RenderRequest):
    """Render structured doc -> PDF via LaTeX with WeasyPrint fallback."""
    try:
        ruleset = _load_ruleset(req.ruleset_name)
        job_id = str(uuid.uuid4())
        out_path = os.path.join(PROCESSED_DIR, f"{job_id}_render.pdf")
        render_pdf(req.document, ruleset, out_path)
        return FileResponse(out_path, filename="document.pdf", media_type="application/pdf")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF render failed: {e}")


@app.get("/render/ruleset")
async def get_ruleset(ruleset_name: Optional[str] = None):
    """Expose the active ruleset so the frontend can mirror its styles in real time."""
    return _load_ruleset(ruleset_name)


# =====================================================================
# Ingest pipeline
# DOCX upload -> StructuredDocument JSON the editor and renderers consume.
# =====================================================================

@app.post("/ingest")
async def ingest_endpoint(manuscript: UploadFile = File(...)):
    """Parse an uploaded .docx into a StructuredDocument and return the JSON.

    Returns:
        {
          "job_id": "...",
          "document": StructuredDocument,
          "stats": { blocks, headings, references, tables, figures, ... }
        }
    """
    filename = (manuscript.filename or "").lower()
    if not filename.endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail=f"Expected a .docx upload, got {manuscript.filename!r}",
        )

    job_id = str(uuid.uuid4())
    saved_path = os.path.join(UPLOAD_DIR, f"{job_id}_manuscript.docx")

    try:
        with open(saved_path, "wb") as buf:
            shutil.copyfileobj(manuscript.file, buf)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Could not save upload: {e}")

    try:
        structured = ingest_docx(saved_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ingest failed: {e}")

    blocks = structured.blocks
    stats = {
        "blocks": len(blocks),
        "headings": sum(1 for b in blocks if b.type == "heading"),
        "references": sum(1 for b in blocks if b.type == "reference"),
        "tables": sum(1 for b in blocks if b.type == "table"),
        "figures": sum(1 for b in blocks if b.type == "figure"),
        "paragraphs": sum(1 for b in blocks if b.type == "paragraph"),
        "equations": sum(1 for b in blocks if b.type == "equation"),
    }

    return JSONResponse(
        {
            "job_id": job_id,
            "document": structured.model_dump(),
            "stats": stats,
        }
    )


@app.post("/ingest-and-render/{file_type}")
async def ingest_and_render(
    file_type: str,
    manuscript: UploadFile = File(...),
    ruleset_name: Optional[str] = None,
):
    """One-shot: upload a .docx, get back a formatted .docx, .pdf, or .html.
    Skips the editor — useful for the 'reformat my manuscript' button."""
    file_type = (file_type or "").lower().strip()
    if file_type not in ("docx", "pdf", "html"):
        raise HTTPException(
            status_code=400,
            detail=f"file_type must be one of docx | pdf | html (got {file_type!r})",
        )

    filename = (manuscript.filename or "").lower()
    if not filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Expected a .docx upload")

    job_id = str(uuid.uuid4())
    saved_path = os.path.join(UPLOAD_DIR, f"{job_id}_manuscript.docx")
    with open(saved_path, "wb") as buf:
        shutil.copyfileobj(manuscript.file, buf)

    try:
        structured = ingest_docx(saved_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ingest failed: {e}")

    ruleset = _load_ruleset(ruleset_name)
    doc_dict = structured.model_dump()

    if file_type == "html":
        html = render_html(doc_dict, ruleset, standalone=True)
        return HTMLResponse(content=html)

    if file_type == "docx":
        out_path = os.path.join(PROCESSED_DIR, f"{job_id}_render.docx")
        template = DEFAULT_TEMPLATE_PATH if os.path.exists(DEFAULT_TEMPLATE_PATH) else None
        render_docx(doc_dict, ruleset, out_path, template_path=template)
        return FileResponse(
            out_path,
            filename="document.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    # pdf
    out_path = os.path.join(PROCESSED_DIR, f"{job_id}_render.pdf")
    try:
        render_pdf(doc_dict, ruleset, out_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF render failed: {e}")
    return FileResponse(out_path, filename="document.pdf", media_type="application/pdf")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
