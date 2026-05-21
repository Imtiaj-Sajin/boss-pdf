"""SQLAlchemy engine + session for boss-pdf's MySQL DB.

boss-pdf shares the ODIN EMS database — no separate schema. The
boss_pdf_usage table is created inside odin_ems alongside users / employees /
roles / permissions, so every query runs against the default schema and
there are no cross-DB JOINs / COLLATE workarounds.
"""
from __future__ import annotations

import logging
import os
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

log = logging.getLogger("boss-pdf.db")

DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_NAME = os.getenv("DB_NAME", "odin_ems")


def _build_url() -> str:
    # mysql-connector-python driver
    return (
        f"mysql+mysqlconnector://{DB_USER}:{DB_PASSWORD}"
        f"@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset=utf8mb4"
    )


engine: Engine = create_engine(
    _build_url(),
    pool_pre_ping=True,
    pool_recycle=1800,
    future=True,
)

SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)


def get_db():
    """FastAPI dependency yielding a request-scoped Session."""
    db: Session = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def ping() -> bool:
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        return True
    except Exception as e:
        log.error("MySQL ping failed: %s", e)
        return False
