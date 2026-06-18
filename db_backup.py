"""Database backup and restore helpers for ProQuote."""
from __future__ import annotations

import os
import re
import shutil
import sqlite3
import tempfile
from datetime import datetime
from pathlib import Path

import db


BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")
REQUIRED_TABLES = {
    "Projects_Master",
    "Project_Sheets",
    "Items_Catalog",
    "Project_BoQ_Lines",
    "Settings",
    "Users",
    "Roles",
    "RolePerms",
}


def _safe_label(label: str) -> str:
    label = re.sub(r"[^A-Za-z0-9_.-]+", "-", (label or "manual").strip())
    return label.strip("-._") or "manual"


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def _sidecar_paths(path: str) -> list[str]:
    return [f"{path}-wal", f"{path}-shm"]


def _checkpoint_current_db() -> None:
    if not os.path.exists(db.DB_PATH):
        return
    conn = sqlite3.connect(db.DB_PATH, timeout=30)
    try:
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
    finally:
        conn.close()


def validate_database(path: str) -> tuple[bool, str]:
    if not os.path.exists(path):
        return False, "Database file was not found."
    try:
        conn = sqlite3.connect(path)
        try:
            row = conn.execute("PRAGMA integrity_check;").fetchone()
            if not row or row[0] != "ok":
                return False, f"SQLite integrity check failed: {row[0] if row else 'no result'}"
            tables = {
                r[0]
                for r in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
        finally:
            conn.close()
    except sqlite3.DatabaseError as exc:
        return False, f"Not a valid SQLite database: {exc}"

    missing = sorted(REQUIRED_TABLES - tables)
    if missing:
        return False, "Backup does not look like a ProQuote database. Missing: " + ", ".join(missing)
    return True, "OK"


def create_backup(label: str = "manual") -> str:
    os.makedirs(BACKUP_DIR, exist_ok=True)
    conn = db.init_db()
    conn.close()
    _checkpoint_current_db()

    label = _safe_label(label)
    dest = os.path.join(BACKUP_DIR, f"proquote-{label}-{_timestamp()}.db")
    src = sqlite3.connect(db.DB_PATH, timeout=30)
    try:
        out = sqlite3.connect(dest)
        try:
            src.backup(out)
        finally:
            out.close()
    finally:
        src.close()

    ok, message = validate_database(dest)
    if not ok:
        try:
            os.remove(dest)
        finally:
            raise RuntimeError(message)
    return dest


def restore_from_bytes(uploaded_bytes: bytes) -> tuple[str, str]:
    if not uploaded_bytes:
        raise ValueError("Uploaded backup is empty.")

    os.makedirs(db.DATA_DIR, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix="proquote-restore-", suffix=".db", dir=db.DATA_DIR)
    try:
        with os.fdopen(fd, "wb") as tmp:
            tmp.write(uploaded_bytes)

        ok, message = validate_database(temp_path)
        if not ok:
            raise RuntimeError(message)

        safety_backup = create_backup("before-restore") if os.path.exists(db.DB_PATH) else ""
        _checkpoint_current_db()
        for sidecar in _sidecar_paths(db.DB_PATH):
            if os.path.exists(sidecar):
                os.remove(sidecar)
        shutil.copy2(temp_path, db.DB_PATH)
        for sidecar in _sidecar_paths(db.DB_PATH):
            if os.path.exists(sidecar):
                os.remove(sidecar)

        conn = db.init_db()
        conn.close()
        return db.DB_PATH, safety_backup
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass


def list_backups(limit: int = 10) -> list[dict]:
    if not os.path.isdir(BACKUP_DIR):
        return []
    files = sorted(Path(BACKUP_DIR).glob("*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    out = []
    for path in files[:limit]:
        stat = path.stat()
        out.append({"path": str(path), "name": path.name, "size": stat.st_size, "mtime": stat.st_mtime})
    return out