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
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader, PdfWriter
from sqlalchemy.orm import Session

from . import usage as usage_mod
from .auth import CurrentUser, get_current_user, router as auth_router
from .converter import convert_pdf_bytes, parse_page_spec
from .db import get_db
from .usage import router as usage_router

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger("boss-pdf")

ROOT = Path(__file__).resolve().parent.parent
WEB = ROOT / "web"

app = FastAPI(title="boss-pdf", description="The Boss of all PDFs.")

# CORS — the spec calls for allow_origins=["*"]. Tokens travel in Authorization
# header (not cookies) so this is safe.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=[
        "X-Boss-OCR-Used",
        "X-Boss-OCR-Engine",
        "X-Boss-Log-Id",
        "X-Boss-Pages",
        "X-Boss-Usage-Id",
        # Without these, a cross-origin fetch() can't read the table-scrape counts.
        "X-Table-Rows",
        "X-Table-Pages",
    ],
)

app.include_router(auth_router, prefix="/api")
app.include_router(usage_router, prefix="/api")


@app.on_event("startup")
def _on_startup() -> None:
    usage_mod.ensure_table()


MAX_BYTES = 500 * 1024 * 1024  # 500 MB

# In-memory cache for per-conversion diagnostics (log + OCR preview HTML).
# Keyed by a random log_id; expires after _RESULT_TTL seconds.
_RESULT_TTL = 3600  # 1 hour
_results_lock = threading.Lock()
_results_cache: dict[str, dict] = {}


def _stash_result(log_text: str, preview_html: Optional[str],
                  usage_id: Optional[int]) -> str:
    log_id = uuid.uuid4().hex
    now = time.time()
    with _results_lock:
        _results_cache[log_id] = {
            "log_text": log_text,
            "preview_html": preview_html,
            "usage_id": usage_id,
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


def _ar_convert_sync(pdf_bytes: bytes) -> io.BytesIO:
    """Parse a JDE 'A/R Details with Aging' PDF into the template1.xlsx layout."""
    import sys
    import tempfile
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from ar_to_excel import export as ar_export, parse as ar_parse

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp = f.name
    try:
        data = ar_parse(tmp)
        buf = io.BytesIO()
        ar_export(data, buf, template_path=str(ROOT / "template1.xlsx"))
        buf.seek(0)
        return buf
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


@app.post("/api/table/detect")
async def table_detect(file: UploadFile = File(...), page: int = Form(1)):
    """Auto-detect a page's column/row separators. Lines come back normalized
    to [0,1] so the browser can render at any scale and still line up."""
    from .tablescrape import detect
    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)
    try:
        return await run_in_threadpool(detect, pdf_bytes, int(page))
    except Exception as e:
        log.exception("table detect failed")
        raise HTTPException(status_code=500, detail=f"Detection failed: {e}") from e


def _parse_lines(columns: str, rows: str) -> tuple[list[float], list[float]]:
    import json
    try:
        cols = [float(c) for c in (json.loads(columns) or [])]
        rws = [float(r) for r in (json.loads(rows) or [])]
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid lines: {e}") from e
    return cols, rws


@app.post("/api/table/scrape-all")
async def table_scrape_all(file: UploadFile = File(...),
                           columns: str = Form("[]"), rows: str = Form("[]"),
                           auto_rows: Optional[str] = Form("true"),
                           drop_empty: Optional[str] = Form("true")):
    """Preview: apply the calibrated grid to every page. Rows are page-prefixed."""
    from .tablescrape import scrape_all
    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)
    cols, rws = _parse_lines(columns, rows)
    auto = (auto_rows or "true").lower() in ("true", "1", "yes", "on")
    drop = (drop_empty or "true").lower() in ("true", "1", "yes", "on")
    if not cols:
        raise HTTPException(status_code=400, detail="No column lines given.")
    try:
        from .tablescrape import preview_rows
        pages = await run_in_threadpool(scrape_all, pdf_bytes, cols, rws, auto, drop)
    except Exception as e:
        log.exception("table scrape failed")
        raise HTTPException(status_code=500, detail=f"Scrape failed: {e}") from e
    combined = preview_rows(pages)
    total = sum(len(t) for p in pages for t in p.tables)
    return {"rows": combined, "n_rows": total, "page_count": len(pages)}


@app.post("/api/table/excel-all")
async def table_excel_all(file: UploadFile = File(...),
                          columns: str = Form("[]"), rows: str = Form("[]"),
                          auto_rows: Optional[str] = Form("true"),
                          drop_empty: Optional[str] = Form("true")):
    """Whole PDF -> one Excel, built by the CORE converter using the user's
    corrected column lines. Output matches 'Make it Excel' plus the fix."""
    from .tablescrape import build_xlsx, scrape_all
    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)
    cols, rws = _parse_lines(columns, rows)
    auto = (auto_rows or "true").lower() in ("true", "1", "yes", "on")
    drop = (drop_empty or "true").lower() in ("true", "1", "yes", "on")
    if not cols:
        raise HTTPException(status_code=400, detail="No column lines given.")
    try:
        pages = await run_in_threadpool(scrape_all, pdf_bytes, cols, rws, auto, drop)
        xlsx, total = await run_in_threadpool(build_xlsx, pages)
    except Exception as e:
        log.exception("table excel failed")
        raise HTTPException(status_code=500, detail=f"Excel failed: {e}") from e
    stem = os.path.splitext(_safe_filename(file.filename or "table"))[0]
    return StreamingResponse(
        io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{stem}.xlsx"',
            "X-Table-Rows": str(total),
            "X-Table-Pages": str(len(pages)),
        },
    )


@app.post("/api/datapuller/export")
async def datapuller_export(files: list[UploadFile] = File(...),
                            regions: str = Form(...)):
    """Pull named rectangular regions from many same-layout PDFs into one
    row-per-file Excel table. No auth — pure file-in/file-out."""
    import json

    from .datapuller import Region, build_xlsx, extract

    try:
        raw = json.loads(regions)
        regs = [
            Region(name=str(r["name"]).strip() or f"field{i+1}",
                   page=int(r.get("page", 1)),
                   x0=r["x0"], y0=r["y0"], x1=r["x1"], y1=r["y1"])
            for i, r in enumerate(raw)
        ]
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid regions: {e}") from e
    if not regs:
        raise HTTPException(status_code=400, detail="No datapoints were defined.")

    # column order = first appearance of each distinct name
    column_names = list(dict.fromkeys(r.name for r in regs))

    filenames: list[str] = []
    rows: list[dict] = []
    for uf in files:
        if not uf.filename or not uf.filename.lower().endswith(".pdf"):
            continue
        b = await uf.read()
        if b[:5] != b"%PDF-":
            continue
        try:
            row = await run_in_threadpool(extract, b, regs)
        except Exception:
            log.exception("datapuller extract failed for %s", uf.filename)
            row = {}
        filenames.append(uf.filename)
        rows.append(row)

    if not filenames:
        raise HTTPException(status_code=400, detail="No valid PDFs in the upload.")

    xlsx = await run_in_threadpool(build_xlsx, filenames, rows, column_names)
    return StreamingResponse(
        io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="datapuller.xlsx"'},
    )


@app.post("/api/convert-batch")
async def convert_batch(files: list[UploadFile] = File(...),
                        pages: Optional[str] = Form(None),
                        force_ocr: Optional[str] = Form(None),
                        x_session_id: Optional[str] = Header(None),
                        current_user: CurrentUser = Depends(get_current_user),
                        db: Session = Depends(get_db)):
    """Convert a whole folder of PDFs in one shot; returns a ZIP of .xlsx files.
    Same conversion as /api/convert, looped per file."""
    force_flag = (force_ocr or "").lower() in ("true", "1", "yes", "on")
    zip_buf = io.BytesIO()
    n_ok = 0
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for uf in files:
            if not uf.filename or not uf.filename.lower().endswith(".pdf"):
                continue
            pdf_bytes = await uf.read()
            if pdf_bytes[:5] != b"%PDF-" or len(pdf_bytes) > MAX_BYTES:
                continue
            try:
                result = await run_in_threadpool(
                    convert_pdf_bytes, pdf_bytes, None, uf.filename, force_flag,
                )
            except Exception:
                log.exception("batch convert failed for %s", uf.filename)
                continue
            stem = os.path.splitext(_safe_filename(uf.filename))[0]
            zf.writestr(f"{stem}.xlsx", result.xlsx_bytes)
            n_ok += 1
            usage_mod.log_usage(
                db, user=current_user, session_id=x_session_id,
                operation="convert", batch_name=uf.filename,
                file_names=[uf.filename], files_processed=1,
                pages_processed=len(result.pages_summary),
                tables_extracted=sum(int(p.get("tables") or 0) for p in result.pages_summary),
                ocr_used=bool(result.ocr_used), ocr_engine=result.ocr_engine or None,
            )
    if n_ok == 0:
        raise HTTPException(status_code=400, detail="No PDFs could be converted.")
    zip_buf.seek(0)
    return StreamingResponse(
        zip_buf, media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="boss-pdf-batch.zip"'},
    )


@app.post("/api/ar/convert")
async def ar_convert(file: UploadFile = File(...)):
    """Upload an A/R Aging PDF, get the template-formatted .xlsx. No auth — this
    converter doesn't touch the user database."""
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")
    pdf_bytes = await file.read()
    _validate_pdf(pdf_bytes)
    try:
        buf = await run_in_threadpool(_ar_convert_sync, pdf_bytes)
    except Exception as e:
        log.exception("ar convert failed")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}") from e
    out_name = os.path.splitext(_safe_filename(file.filename))[0] + ".xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


@app.post("/api/pdf-info")
async def pdf_info(file: UploadFile = File(...),
                   _: CurrentUser = Depends(get_current_user)) -> dict:
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
                  force_ocr: Optional[str] = Form(None),
                  x_session_id: Optional[str] = Header(None),
                  current_user: CurrentUser = Depends(get_current_user),
                  db: Session = Depends(get_db)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")

    uploaded_at = datetime.now(timezone.utc)

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

    log.info("converting %s (%d bytes) pages=%s%s user=%s",
             file.filename, len(pdf_bytes), pages or "all",
             " FORCE_OCR" if force_flag else "", current_user.username)
    try:
        # Conversion is CPU-bound and synchronous; run it in a worker thread so
        # it doesn't block the event loop (which would freeze every other
        # request — health checks, auth, other uploads — for the whole job).
        result = await run_in_threadpool(
            convert_pdf_bytes,
            pdf_bytes, pages=page_set,
            filename=file.filename, force_ocr=force_flag,
        )
    except Exception as e:
        log.exception("conversion failed")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}") from e

    finished_at = datetime.now(timezone.utc)

    pages_processed = len(result.pages_summary)
    tables_extracted = sum(int(p.get("tables") or 0) for p in result.pages_summary)

    usage_id = usage_mod.log_usage(
        db,
        user=current_user,
        session_id=x_session_id,
        operation="convert",
        batch_name=file.filename,
        file_names=[file.filename],
        files_processed=1,
        pages_processed=pages_processed,
        tables_extracted=tables_extracted,
        ocr_used=bool(result.ocr_used),
        ocr_engine=result.ocr_engine or None,
        uploaded_at=uploaded_at,
        finished_at=finished_at,
    )

    # The xlsx is auto-downloaded by the browser on response — count it as 1.
    # Route the bump to ocr_downloads vs downloads so the dashboard can split
    # "native-text load" from "OCR load" without re-deriving from ocr_used.
    if usage_id:
        usage_mod.bump_download(
            db, usage_id, current_user.id,
            kind="ocr" if result.ocr_used else "download",
        )

    # Stash log + preview server-side; client fetches them from /api/log/{id}
    # and /api/preview/{id} when the user clicks the corresponding button.
    log_id = _stash_result(result.log_text, result.preview_html, usage_id)

    stem = os.path.splitext(_safe_filename(file.filename))[0]
    out_name = f"{stem}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{out_name}"',
        "X-Boss-OCR-Used": "true" if result.ocr_used else "false",
        "X-Boss-OCR-Engine": result.ocr_engine or "",
        "X-Boss-Log-Id": log_id,
        "X-Boss-Pages": str(pages_processed),
        "X-Boss-Usage-Id": str(usage_id) if usage_id else "",
    }
    return StreamingResponse(
        io.BytesIO(result.xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/log/{log_id}")
def get_log(log_id: str, _: CurrentUser = Depends(get_current_user)):
    data = _get_result(log_id)
    if not data:
        raise HTTPException(status_code=404, detail="Log expired or not found.")
    return StreamingResponse(
        io.BytesIO(data["log_text"].encode("utf-8")),
        media_type="text/plain; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="boss-pdf-log.txt"'},
    )


@app.get("/api/preview/{log_id}")
def get_preview(log_id: str, _: CurrentUser = Depends(get_current_user)):
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
                ranges: str = Form(...),
                x_session_id: Optional[str] = Header(None),
                current_user: CurrentUser = Depends(get_current_user),
                db: Session = Depends(get_db)):
    """Split a PDF into one or more output PDFs.

    `ranges` is a JSON array of strings, each a page-spec like "1-3" or "5".
    Returns a single PDF if one range, otherwise a ZIP of named PDFs.
    """
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")

    uploaded_at = datetime.now(timezone.utc)

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
    pages_processed = sum(len(idx) for _, idx in parsed)

    if len(parsed) == 1:
        spec, idx = parsed[0]
        out = _extract_pages_to_pdf(reader, idx)
        finished_at = datetime.now(timezone.utc)
        usage_id = usage_mod.log_usage(
            db,
            user=current_user,
            session_id=x_session_id,
            operation="split",
            batch_name=file.filename,
            file_names=[file.filename],
            files_processed=1,
            pages_processed=pages_processed,
            split_parts=1,
            uploaded_at=uploaded_at,
            finished_at=finished_at,
        )
        out_name = f"{stem}_p{spec.replace(',', '_')}.pdf"
        headers = {
            "Content-Disposition": f'attachment; filename="{out_name}"',
            "X-Boss-Usage-Id": str(usage_id) if usage_id else "",
        }
        return StreamingResponse(io.BytesIO(out), media_type="application/pdf", headers=headers)

    # Multiple ranges -> zip them up
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (spec, idx) in enumerate(parsed, start=1):
            part = _extract_pages_to_pdf(reader, idx)
            label = spec.replace(",", "_")
            zf.writestr(f"{stem}_part{i}_p{label}.pdf", part)
    zip_buf.seek(0)
    finished_at = datetime.now(timezone.utc)
    usage_id = usage_mod.log_usage(
        db,
        user=current_user,
        session_id=x_session_id,
        operation="split",
        batch_name=file.filename,
        file_names=[file.filename],
        files_processed=1,
        pages_processed=pages_processed,
        split_parts=len(parsed),
        uploaded_at=uploaded_at,
        finished_at=finished_at,
    )
    out_name = f"{stem}_split.zip"
    headers = {
        "Content-Disposition": f'attachment; filename="{out_name}"',
        "X-Boss-Usage-Id": str(usage_id) if usage_id else "",
    }
    return StreamingResponse(zip_buf, media_type="application/zip", headers=headers)


# Serve the static frontend at /
if WEB.exists():
    app.mount("/static", StaticFiles(directory=str(WEB)), name="static")

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(str(WEB / "index.html"))

    @app.get("/login")
    def login_page() -> FileResponse:
        return FileResponse(str(WEB / "login.html"))

    @app.get("/ar")
    def ar_page() -> FileResponse:
        return FileResponse(str(WEB / "ar.html"))

    @app.get("/datapuller")
    def datapuller_page() -> FileResponse:
        return FileResponse(str(WEB / "datapuller.html"))

    @app.get("/tablescrape")
    def tablescrape_page() -> FileResponse:
        return FileResponse(str(WEB / "tablescrape.html"))

    @app.get("/usage")
    def usage_page() -> FileResponse:
        return FileResponse(str(WEB / "usage.html"))
