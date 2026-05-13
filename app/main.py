"""FastAPI backend for boss-pdf: PDF -> Excel converter + splitter."""
from __future__ import annotations

import io
import json
import logging
import os
import threading
import time
import uuid
import zipfile
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader, PdfWriter

from .converter import convert_pdf_bytes, parse_page_spec

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger("boss-pdf")

ROOT = Path(__file__).resolve().parent.parent
WEB = ROOT / "web"

app = FastAPI(title="boss-pdf", description="The Boss of all PDFs.")

MAX_BYTES = 50 * 1024 * 1024  # 50 MB

# In-memory cache for per-conversion diagnostics (log + OCR preview HTML).
# Keyed by a random log_id; expires after _RESULT_TTL seconds.
_RESULT_TTL = 3600  # 1 hour
_results_lock = threading.Lock()
_results_cache: dict[str, dict] = {}


def _stash_result(log_text: str, preview_html: Optional[str]) -> str:
    log_id = uuid.uuid4().hex
    now = time.time()
    with _results_lock:
        _results_cache[log_id] = {
            "log_text": log_text,
            "preview_html": preview_html,
            "expires_at": now + _RESULT_TTL,
        }
        # opportunistic cleanup
        for k in list(_results_cache.keys()):
            if _results_cache[k]["expires_at"] < now:
                del _results_cache[k]
    return log_id


def _get_result(log_id: str) -> Optional[dict]:
    with _results_lock:
        data = _results_cache.get(log_id)
        if data and data["expires_at"] < time.time():
            del _results_cache[log_id]
            return None
        return data


def _validate_pdf(pdf_bytes: bytes) -> None:
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="Empty file. The Boss expected something.")
    if len(pdf_bytes) > MAX_BYTES:
        raise HTTPException(status_code=413,
                            detail=f"File exceeds {MAX_BYTES // (1024*1024)} MB. Trim it down, kid.")
    if pdf_bytes[:5] != b"%PDF-":
        raise HTTPException(status_code=400, detail="That's not a PDF. Don't try to fool The Boss.")


def _safe_filename(name: str) -> str:
    base = os.path.basename(name or "file")
    return "".join(c for c in base if c.isalnum() or c in ("-", "_", ".", " ")).strip() or "file"


@app.get("/api/health")
def health() -> dict:
    return {"ok": True}


@app.post("/api/pdf-info")
async def pdf_info(file: UploadFile = File(...)) -> dict:
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")
    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise HTTPException(status_code=400,
                                    detail="PDF is password-protected. The Boss doesn't do passwords.")
        pages = len(reader.pages)
    except HTTPException:
        raise
    except Exception as e:
        log.exception("pdf-info failed")
        raise HTTPException(status_code=400, detail=f"Could not read PDF: {e}") from e
    return {"pages": pages, "filename": file.filename, "size": len(pdf_bytes)}


@app.post("/api/convert")
async def convert(file: UploadFile = File(...),
                  pages: Optional[str] = Form(None),
                  force_ocr: Optional[str] = Form(None)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")

    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)

    # Determine page count for validation
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        total = len(reader.pages)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read PDF: {e}") from e

    page_set: Optional[set[int]] = None
    if pages and pages.strip().lower() not in ("", "all"):
        try:
            page_set = parse_page_spec(pages, total)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e

    force_flag = (force_ocr or "").lower() in ("true", "1", "yes", "on")

    log.info("converting %s (%d bytes) pages=%s%s",
             file.filename, len(pdf_bytes), pages or "all",
             " FORCE_OCR" if force_flag else "")
    try:
        result = convert_pdf_bytes(
            pdf_bytes, pages=page_set,
            filename=file.filename, force_ocr=force_flag,
        )
    except Exception as e:
        log.exception("conversion failed")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}") from e

    # Stash log + preview server-side; client fetches them from /api/log/{id}
    # and /api/preview/{id} when the user clicks the corresponding button.
    log_id = _stash_result(result.log_text, result.preview_html)

    stem = os.path.splitext(_safe_filename(file.filename))[0]
    out_name = f"{stem}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{out_name}"',
        "X-Boss-OCR-Used": "true" if result.ocr_used else "false",
        "X-Boss-OCR-Engine": result.ocr_engine or "",
        "X-Boss-Log-Id": log_id,
        "X-Boss-Pages": str(len(result.pages_summary)),
        "Access-Control-Expose-Headers":
            "X-Boss-OCR-Used, X-Boss-OCR-Engine, X-Boss-Log-Id, X-Boss-Pages",
    }
    return StreamingResponse(
        io.BytesIO(result.xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/log/{log_id}")
def get_log(log_id: str):
    data = _get_result(log_id)
    if not data:
        raise HTTPException(status_code=404, detail="Log expired or not found.")
    return StreamingResponse(
        io.BytesIO(data["log_text"].encode("utf-8")),
        media_type="text/plain; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="boss-pdf-log.txt"'},
    )


@app.get("/api/preview/{log_id}")
def get_preview(log_id: str):
    data = _get_result(log_id)
    if not data or not data["preview_html"]:
        raise HTTPException(status_code=404, detail="Preview not found.")
    return HTMLResponse(content=data["preview_html"])


def _extract_pages_to_pdf(reader: PdfReader, page_indices: list[int]) -> bytes:
    writer = PdfWriter()
    for idx in page_indices:
        writer.add_page(reader.pages[idx])
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


@app.post("/api/split")
async def split(file: UploadFile = File(...),
                ranges: str = Form(...)):
    """Split a PDF into one or more output PDFs.

    `ranges` is a JSON array of strings, each a page-spec like "1-3" or "5".
    Returns a single PDF if one range, otherwise a ZIP of named PDFs.
    """
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")

    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise HTTPException(status_code=400,
                                    detail="PDF is password-protected.")
        total = len(reader.pages)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read PDF: {e}") from e

    try:
        spec_list = json.loads(ranges)
        if not isinstance(spec_list, list) or not spec_list:
            raise ValueError("ranges must be a non-empty list")
        spec_list = [str(s) for s in spec_list]
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid ranges: {e}") from e

    parsed: list[tuple[str, list[int]]] = []
    for spec in spec_list:
        try:
            pages = parse_page_spec(spec, total)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
        # 0-indexed for pypdf, sorted
        idx = sorted(p - 1 for p in pages)
        parsed.append((spec, idx))

    stem = os.path.splitext(_safe_filename(file.filename))[0]

    if len(parsed) == 1:
        spec, idx = parsed[0]
        out = _extract_pages_to_pdf(reader, idx)
        out_name = f"{stem}_p{spec.replace(',', '_')}.pdf"
        headers = {"Content-Disposition": f'attachment; filename="{out_name}"'}
        return StreamingResponse(io.BytesIO(out), media_type="application/pdf", headers=headers)

    # Multiple ranges -> zip them up
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (spec, idx) in enumerate(parsed, start=1):
            part = _extract_pages_to_pdf(reader, idx)
            label = spec.replace(",", "_")
            zf.writestr(f"{stem}_part{i}_p{label}.pdf", part)
    zip_buf.seek(0)
    out_name = f"{stem}_split.zip"
    headers = {"Content-Disposition": f'attachment; filename="{out_name}"'}
    return StreamingResponse(zip_buf, media_type="application/zip", headers=headers)


# Serve the static frontend at /
if WEB.exists():
    app.mount("/static", StaticFiles(directory=str(WEB)), name="static")

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(str(WEB / "index.html"))
