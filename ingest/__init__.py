"""
StyleLock ingest layer.

This is the SciSpace-style INPUT edge: take a user-authored .docx and
produce a StructuredDocument (the same model the renderers consume).

Public API:
    from ingest import ingest_docx
    structured: StructuredDocument = ingest_docx("/path/to/manuscript.docx")
"""

from .docx_ingest import ingest_docx, ingest_docx_to_dict

__all__ = ["ingest_docx", "ingest_docx_to_dict"]
