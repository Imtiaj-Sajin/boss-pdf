"""ODIN EMS auth for boss-pdf.

This app does NOT own a users table. Identity lives in odin_ems.users
on the same MySQL server. We verify bcrypt against odin_ems.users.password_hash
and issue a JWT that's interoperable with the ODIN EMS Node backend (same secret).
"""
from __future__ import annotations

import logging
import os
import re
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import APIRouter, Depends, Header, HTTPException, status
from jose import JWTError, jwt
from passlib.context import CryptContext
from pydantic import BaseModel
from sqlalchemy import text
from sqlalchemy.orm import Session

from .db import get_db

log = logging.getLogger("boss-pdf.auth")

JWT_SECRET = os.getenv("JWT_SECRET", "change-me")
JWT_ALGORITHM = "HS256"


def _expires_seconds() -> int:
    """Parse JWT_EXPIRES_IN ('24h', '60m', '3600', '7d') -> seconds. Default 24h."""
    raw = os.getenv("JWT_EXPIRES_IN", "24h").strip().lower()
    m = re.match(r"^(\d+)\s*([smhd])?$", raw)
    if not m:
        return 24 * 3600
    n = int(m.group(1))
    unit = m.group(2) or "s"
    return n * {"s": 1, "m": 60, "h": 3600, "d": 86400}[unit]


_pwd_ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")


def verify_password(plain: str, hashed: str) -> bool:
    try:
        return _pwd_ctx.verify(plain, hashed)
    except Exception as e:  # truncated hashes, bad format, etc.
        log.debug("bcrypt verify failed: %s", e)
        return False


# ---------- ODIN EMS cross-DB lookups ----------

def _lookup_ems_user(db: Session, identifier: str) -> Optional[dict]:
    """Find an active user by username OR email."""
    row = db.execute(
        text(
            "SELECT id, username, email, password_hash, is_active "
            "FROM users "
            "WHERE username = :u OR email = :u LIMIT 1"
        ),
        {"u": identifier},
    ).mappings().first()
    return dict(row) if row else None


def _lookup_ems_user_by_id(db: Session, user_id: int) -> Optional[dict]:
    row = db.execute(
        text(
            "SELECT id, username, email, is_active "
            "FROM users WHERE id = :id LIMIT 1"
        ),
        {"id": user_id},
    ).mappings().first()
    return dict(row) if row else None


def _fetch_me(db: Session, user_id: int) -> Optional[dict]:
    """Return user + employee profile + roles + permissions."""
    base = db.execute(
        text(
            """
            SELECT
              u.id, u.username, u.email,
              e.first_name, e.last_name,
              NULLIF(TRIM(CONCAT_WS(' ', e.first_name, e.last_name)), '') AS full_name,
              e.designation, e.profile_photo_url
            FROM users u
            LEFT JOIN employees e ON e.user_id = u.id
            WHERE u.id = :id LIMIT 1
            """
        ),
        {"id": user_id},
    ).mappings().first()
    if not base:
        return None

    roles = [
        r["name"]
        for r in db.execute(
            text(
                """
                SELECT r.name FROM user_roles ur
                JOIN roles r ON r.id = ur.role_id
                WHERE ur.user_id = :id
                """
            ),
            {"id": user_id},
        ).mappings()
    ]

    perms = [
        p["name"]
        for p in db.execute(
            text(
                """
                SELECT DISTINCT p.name
                FROM user_roles ur
                JOIN role_permissions rp ON rp.role_id = ur.role_id
                JOIN permissions p ON p.id = rp.permission_id
                WHERE ur.user_id = :id
                """
            ),
            {"id": user_id},
        ).mappings()
    ]

    out = dict(base)
    out["roles"] = roles
    out["permissions"] = perms
    return out


# ---------- JWT ----------

def _sign(user_id: int, username: str) -> str:
    now = datetime.now(timezone.utc)
    payload = {
        "userId": user_id,
        "username": username,
        "iat": int(now.timestamp()),
        "exp": int((now + timedelta(seconds=_expires_seconds())).timestamp()),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)


class CurrentUser:
    __slots__ = ("id", "username")

    def __init__(self, id: int, username: str) -> None:
        self.id = id
        self.username = username


def get_current_user(
    authorization: Optional[str] = Header(None),
    db: Session = Depends(get_db),
) -> CurrentUser:
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED,
                            detail="Missing bearer token.")
    token = authorization.split(" ", 1)[1].strip()
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
    except JWTError as e:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED,
                            detail=f"Invalid token: {e}") from e

    uid = payload.get("userId")
    if not isinstance(uid, int):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED,
                            detail="Token missing userId.")
    user = _lookup_ems_user_by_id(db, uid)
    if not user or not user.get("is_active"):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED,
                            detail="User no longer active.")
    return CurrentUser(id=user["id"], username=user["username"])


# ---------- Router ----------

router = APIRouter(prefix="/auth", tags=["auth"])


class LoginRequest(BaseModel):
    username: str  # username OR email
    password: str


class LoginResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: dict


@router.post("/login", response_model=LoginResponse)
def login(body: LoginRequest, db: Session = Depends(get_db)) -> LoginResponse:
    ident = (body.username or "").strip()
    if not ident or not body.password:
        raise HTTPException(status_code=400, detail="Username and password are required.")
    user = _lookup_ems_user(db, ident)
    if not user or not verify_password(body.password, user["password_hash"] or ""):
        raise HTTPException(status_code=401, detail="Invalid credentials.")
    if not user.get("is_active"):
        raise HTTPException(status_code=403, detail="Account is disabled.")
    token = _sign(user["id"], user["username"])
    return LoginResponse(
        access_token=token,
        user={"id": user["id"], "username": user["username"], "email": user.get("email")},
    )


@router.get("/me")
def me(current: CurrentUser = Depends(get_current_user),
       db: Session = Depends(get_db)) -> dict:
    data = _fetch_me(db, current.id)
    if not data:
        raise HTTPException(status_code=404, detail="User not found.")
    return data
