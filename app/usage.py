"""Usage tracking for boss-pdf.

Schema: boss_pdf_usage (lives in the shared odin_ems database) — one row per
upload/job, with denormalized username so other apps can read it cheaply.

Fields tracked (per row):
  - operation        : 'convert' | 'split'
  - batch_name       : the uploaded filename
  - files_processed  : 1 per upload (one PDF in = one row)
  - pages_processed  : pages actually worked on (selected pages for convert,
                       pages covered for split)
  - tables_extracted : tables found across processed pages (convert only)
  - ocr_used         : 1 if any OCR ran on this job
  - ocr_engine       : engine name (paddleocr-v3, rapidocr, tesseract, ...)
  - split_parts      : number of output PDFs produced by a split
  - downloads        : Excel downloads from a NATIVE-text convert
  - ocr_downloads    : Excel downloads from an OCR'd convert (so you can tell
                       "native-text vs OCR" load apart in the dashboard)
  - split_downloads  : per-file PDF downloads from a split job
  - uploaded_at      : when the request started server-side
  - finished_at      : when the work finished server-side
  - duration_ms      : finished_at - uploaded_at, in ms (for time-saved math)
  - session_id       : opaque id passed by the frontend, lets us count sessions
"""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone


def _iso_utc(v):
    """Naive datetimes coming back from MySQL TIMESTAMP columns are UTC by
    convention (we always insert UTC). Emit them as ISO with a trailing 'Z'
    so JS's new Date(...) treats them as UTC instead of local time."""
    if not isinstance(v, datetime):
        return v
    if v.tzinfo is None:
        v = v.replace(tzinfo=timezone.utc)
    return v.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
from typing import Any, Optional

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import text
from sqlalchemy.orm import Session

from .auth import CurrentUser, get_current_user
from .db import engine, get_db

log = logging.getLogger("boss-pdf.usage")

TABLE = "boss_pdf_usage"


_DOWNLOAD_COLUMNS = {
    # kind -> physical column name
    "download": "downloads",
    "ocr":      "ocr_downloads",
    "split":    "split_downloads",
}


def ensure_table() -> None:
    """Create boss_pdf_usage if it doesn't exist. Safe to call repeatedly.

    Also runs forward-only column additions for the per-kind download counters
    so we don't strand older deployments on the previous schema.
    """
    ddl = f"""
    CREATE TABLE IF NOT EXISTS {TABLE} (
      id                INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
      user_id           INT UNSIGNED NOT NULL,
      username          VARCHAR(150)  DEFAULT NULL,
      session_id        VARCHAR(64)   DEFAULT NULL,
      operation         VARCHAR(32)   NOT NULL,
      batch_name        VARCHAR(255)  DEFAULT NULL,
      files_processed   INT NOT NULL  DEFAULT 0,
      pages_processed   INT NOT NULL  DEFAULT 0,
      tables_extracted  INT NOT NULL  DEFAULT 0,
      ocr_used          TINYINT(1)    NOT NULL DEFAULT 0,
      ocr_engine        VARCHAR(64)   DEFAULT NULL,
      split_parts       INT NOT NULL  DEFAULT 0,
      downloads         INT NOT NULL  DEFAULT 0,
      ocr_downloads     INT NOT NULL  DEFAULT 0,
      split_downloads   INT NOT NULL  DEFAULT 0,
      file_names        JSON          DEFAULT NULL,
      uploaded_at       TIMESTAMP     NULL DEFAULT NULL,
      finished_at       TIMESTAMP     NULL DEFAULT NULL,
      duration_ms       INT UNSIGNED  DEFAULT NULL,
      used_at           TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      KEY idx_user_id   (user_id),
      KEY idx_username  (username),
      KEY idx_used_at   (used_at),
      KEY idx_operation (operation),
      KEY idx_session   (session_id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
    """
    try:
        with engine.begin() as conn:
            conn.execute(text(ddl))
            # Forward-only: tables created before the per-kind split lacked
            # ocr_downloads / split_downloads. IF NOT EXISTS on ADD COLUMN is
            # MySQL 8.0.29+; we guard with a probe so older servers still work.
            for col in ("ocr_downloads", "split_downloads"):
                exists = conn.execute(
                    text(
                        "SELECT 1 FROM information_schema.columns "
                        "WHERE table_schema = DATABASE() "
                        "  AND table_name = :t AND column_name = :c"
                    ),
                    {"t": TABLE, "c": col},
                ).first()
                if not exists:
                    conn.execute(
                        text(
                            f"ALTER TABLE {TABLE} "
                            f"ADD COLUMN {col} INT NOT NULL DEFAULT 0 AFTER downloads"
                        )
                    )
                    log.info("added column %s.%s", TABLE, col)
        log.info("boss_pdf_usage table ready.")
    except Exception as e:
        log.error("Could not ensure boss_pdf_usage table: %s", e)


def log_usage(
    db: Session,
    *,
    user: CurrentUser,
    session_id: Optional[str],
    operation: str,
    batch_name: str,
    file_names: Optional[list[str]] = None,
    files_processed: int = 1,
    pages_processed: int = 0,
    tables_extracted: int = 0,
    ocr_used: bool = False,
    ocr_engine: Optional[str] = None,
    split_parts: int = 0,
    uploaded_at: Optional[datetime] = None,
    finished_at: Optional[datetime] = None,
) -> Optional[int]:
    """Insert a usage row. Returns the new id, or None on failure."""
    fin = finished_at or datetime.now(timezone.utc)
    up = uploaded_at or fin
    duration_ms = max(0, int((fin - up).total_seconds() * 1000))
    try:
        result = db.execute(
            text(
                f"""
                INSERT INTO {TABLE}
                  (user_id, username, session_id, operation, batch_name,
                   files_processed, pages_processed, tables_extracted,
                   ocr_used, ocr_engine, split_parts,
                   file_names, uploaded_at, finished_at, duration_ms)
                VALUES
                  (:uid, :uname, :sid, :op, :name,
                   :fp, :pp, :te,
                   :ou, :oe, :sp,
                   :fn, :up, :fin, :dur)
                """
            ),
            {
                "uid": user.id,
                "uname": (user.username or "")[:150] or None,
                "sid": (session_id or "")[:64] or None,
                "op": operation[:32],
                "name": (batch_name or "")[:255] or None,
                "fp": files_processed,
                "pp": pages_processed,
                "te": tables_extracted,
                "ou": 1 if ocr_used else 0,
                "oe": (ocr_engine or "")[:64] or None,
                "sp": split_parts,
                "fn": json.dumps(file_names[:200]) if file_names else None,
                "up": up.replace(tzinfo=None),
                "fin": fin.replace(tzinfo=None),
                "dur": duration_ms,
            },
        )
        db.commit()
        return int(result.lastrowid) if result.lastrowid else None
    except Exception as e:
        log.warning("usage log failed: %s", e)
        db.rollback()
        return None


def bump_download(db: Session, usage_id: int, user_id: int,
                  kind: str = "download") -> bool:
    """Increment the appropriate download counter on a usage row.

    kind:
      'download' -> downloads        (native-text convert artifact)
      'ocr'      -> ocr_downloads    (OCR'd convert artifact)
      'split'    -> split_downloads  (per-part PDF from a split job)
    """
    col = _DOWNLOAD_COLUMNS.get(kind, "downloads")
    try:
        res = db.execute(
            text(
                f"UPDATE {TABLE} SET {col} = {col} + 1 "
                f"WHERE id = :id AND user_id = :uid"
            ),
            {"id": usage_id, "uid": user_id},
        )
        db.commit()
        return res.rowcount > 0
    except Exception as e:
        log.warning("download bump (%s) failed: %s", kind, e)
        db.rollback()
        return False


# ---------- API ----------

router = APIRouter(prefix="/usage", tags=["usage"])


def _row_to_dict(row: Any) -> dict:
    d = dict(row)
    fn = d.get("file_names")
    if isinstance(fn, (bytes, bytearray)):
        fn = fn.decode("utf-8", errors="replace")
    if isinstance(fn, str):
        try:
            d["file_names"] = json.loads(fn)
        except Exception:
            d["file_names"] = None
    # JSON-serialize datetimes as ISO-UTC with trailing 'Z'
    for k in ("uploaded_at", "finished_at", "used_at"):
        if isinstance(d.get(k), datetime):
            d[k] = _iso_utc(d[k])
    return d


@router.get("")
def list_usage(
    _: CurrentUser = Depends(get_current_user),  # everyone logged-in can view
    db: Session = Depends(get_db),
) -> dict:
    """All rows + per-row enrichment via cross-DB join."""
    rows = db.execute(
        text(
            f"""
            SELECT
              pu.id, pu.user_id, pu.username, pu.session_id, pu.operation,
              pu.batch_name, pu.files_processed, pu.pages_processed,
              pu.tables_extracted, pu.ocr_used, pu.ocr_engine,
              pu.split_parts,
              pu.downloads, pu.ocr_downloads, pu.split_downloads,
              pu.file_names,
              pu.uploaded_at, pu.finished_at, pu.duration_ms, pu.used_at,
              u.email                AS email,
              NULLIF(TRIM(CONCAT_WS(' ', e.first_name, e.last_name)), '') AS full_name,
              e.designation          AS designation,
              e.profile_photo_url    AS profile_photo_url
            FROM {TABLE} pu
            LEFT JOIN users     u ON u.username = pu.username
            LEFT JOIN employees e ON e.user_id  = u.id
            ORDER BY pu.used_at DESC
            LIMIT 5000
            """
        )
    ).mappings().all()
    return {"rows": [_row_to_dict(r) for r in rows]}


@router.get("/me")
def my_usage(
    current: CurrentUser = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    totals = db.execute(
        text(
            f"""
            SELECT
              COUNT(*)                                  AS jobs,
              COALESCE(SUM(files_processed), 0)         AS files_processed,
              COALESCE(SUM(pages_processed), 0)         AS pages_processed,
              COALESCE(SUM(tables_extracted), 0)        AS tables_extracted,
              COALESCE(SUM(ocr_used), 0)                AS ocr_jobs,
              COALESCE(SUM(split_parts), 0)             AS split_parts,
              COALESCE(SUM(downloads), 0)               AS downloads,
              COALESCE(SUM(ocr_downloads), 0)           AS ocr_downloads,
              COALESCE(SUM(split_downloads), 0)         AS split_downloads,
              COALESCE(SUM(downloads + ocr_downloads + split_downloads), 0) AS total_downloads,
              COALESCE(SUM(duration_ms), 0)             AS total_duration_ms,
              COUNT(DISTINCT session_id)                AS sessions,
              MAX(used_at)                              AS last_used_at
            FROM {TABLE} WHERE user_id = :uid
            """
        ),
        {"uid": current.id},
    ).mappings().first()
    out = dict(totals) if totals else {}
    if isinstance(out.get("last_used_at"), datetime):
        out["last_used_at"] = _iso_utc(out["last_used_at"])
    return out


@router.post("/{usage_id}/download")
def mark_download(
    usage_id: int,
    kind: str = "download",
    current: CurrentUser = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """Bump a download counter. `kind` is one of: download | ocr | split."""
    if kind not in _DOWNLOAD_COLUMNS:
        raise HTTPException(status_code=400,
                            detail=f"Unknown kind '{kind}'. Use one of: "
                                   f"{', '.join(_DOWNLOAD_COLUMNS)}.")
    ok = bump_download(db, usage_id, current.id, kind=kind)
    if not ok:
        raise HTTPException(status_code=404, detail="Usage row not found for this user.")
    return {"ok": True, "kind": kind}
