"""FastAPI backend for boss-pdf: PDF -> Excel converter."""
from __future__ import annotations

import io
import logging
import os
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from .converter import convert_pdf_bytes_to_xlsx_bytes

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger("boss-pdf")

ROOT = Path(__file__).resolve().parent.parent
WEB = ROOT / "web"

app = FastAPI(title="boss-pdf", description="PDF to Excel converter")

MAX_BYTES = 50 * 1024 * 1024  # 50 MB


@app.get("/api/health")
def health() -> dict:
    return {"ok": True}


@app.post("/api/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")

    pdf_bytes = await file.read()
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="Empty file.")
    if len(pdf_bytes) > MAX_BYTES:
        raise HTTPException(status_code=413, detail=f"File exceeds {MAX_BYTES // (1024*1024)} MB limit.")
    if not pdf_bytes[:5] == b"%PDF-":
        raise HTTPException(status_code=400, detail="Not a valid PDF (missing %PDF header).")

    log.info("converting %s (%d bytes)", file.filename, len(pdf_bytes))
    try:
        xlsx_bytes = convert_pdf_bytes_to_xlsx_bytes(pdf_bytes)
    except Exception as e:
        log.exception("conversion failed")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}") from e

    out_name = os.path.splitext(os.path.basename(file.filename))[0] + ".xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{out_name}"'}
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


# Serve the static frontend at /
if WEB.exists():
    app.mount("/static", StaticFiles(directory=str(WEB)), name="static")

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(str(WEB / "index.html"))
