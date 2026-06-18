"""Database backup and restore helpers for ProQuote."""
from __future__ import annotations

import os
import re
import shutil
import sqlite3
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import db


BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")
PROFILE_MANIFEST = "proquote-profile-backup.txt"
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


def _asset_files() -> list[Path]:
    assets_dir = Path(db.ASSETS_DIR)
    if not assets_dir.is_dir():
        return []
    return sorted(p for p in assets_dir.rglob("*") if p.is_file())


def _safe_zip_names(zf: zipfile.ZipFile) -> tuple[bool, str]:
    for info in zf.infolist():
        name = info.filename.replace("\\", "/")
        if not name or name.startswith("/") or ".." in Path(name).parts:
            return False, f"Backup contains an unsafe path: {info.filename}"
    return True, "OK"


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


def create_profile_backup(label: str = "manual") -> str:
    """Create a ZIP backup containing the database and branding assets."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    conn = db.init_db()
    conn.close()
    _checkpoint_current_db()

    label = _safe_label(label)
    dest = os.path.join(BACKUP_DIR, f"proquote-profile-{label}-{_timestamp()}.zip")
    with tempfile.TemporaryDirectory(prefix="proquote-profile-", dir=db.DATA_DIR) as tmp_dir:
        db_snapshot = os.path.join(tmp_dir, "proquote.db")
        src = sqlite3.connect(db.DB_PATH, timeout=30)
        try:
            out = sqlite3.connect(db_snapshot)
            try:
                src.backup(out)
            finally:
                out.close()
        finally:
            src.close()

        ok, message = validate_database(db_snapshot)
        if not ok:
            raise RuntimeError(message)

        with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(db_snapshot, "proquote.db")
            zf.writestr(PROFILE_MANIFEST, "ProQuote profile backup\nIncludes proquote.db and assets/.\n")
            for path in _asset_files():
                rel = path.relative_to(db.ASSETS_DIR).as_posix()
                zf.write(path, f"assets/{rel}")

    ok, message = validate_profile_backup(dest)
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


def validate_profile_backup(path: str) -> tuple[bool, str]:
    if not os.path.exists(path):
        return False, "Profile backup file was not found."
    try:
        with zipfile.ZipFile(path) as zf:
            ok, message = _safe_zip_names(zf)
            if not ok:
                return False, message
            if "proquote.db" not in zf.namelist():
                return False, "Profile backup is missing proquote.db."
            with tempfile.TemporaryDirectory(prefix="proquote-validate-", dir=db.DATA_DIR) as tmp_dir:
                db_path = os.path.join(tmp_dir, "proquote.db")
                with zf.open("proquote.db") as src, open(db_path, "wb") as out:
                    shutil.copyfileobj(src, out)
                return validate_database(db_path)
    except zipfile.BadZipFile:
        return False, "Not a valid ZIP profile backup."


def restore_profile_from_bytes(uploaded_bytes: bytes) -> tuple[str, str]:
    if not uploaded_bytes:
        raise ValueError("Uploaded profile backup is empty.")

    os.makedirs(db.DATA_DIR, exist_ok=True)
    os.makedirs(db.ASSETS_DIR, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix="proquote-profile-restore-", suffix=".zip", dir=db.DATA_DIR)
    try:
        with os.fdopen(fd, "wb") as tmp:
            tmp.write(uploaded_bytes)

        ok, message = validate_profile_backup(temp_path)
        if not ok:
            raise RuntimeError(message)

        safety_backup = create_profile_backup("before-profile-restore") if os.path.exists(db.DB_PATH) else ""
        with zipfile.ZipFile(temp_path) as zf, tempfile.TemporaryDirectory(
            prefix="proquote-profile-extract-", dir=db.DATA_DIR
        ) as tmp_dir:
            zf.extract("proquote.db", tmp_dir)
            restored_db = os.path.join(tmp_dir, "proquote.db")

            _checkpoint_current_db()
            for sidecar in _sidecar_paths(db.DB_PATH):
                if os.path.exists(sidecar):
                    os.remove(sidecar)
            shutil.copy2(restored_db, db.DB_PATH)
            for sidecar in _sidecar_paths(db.DB_PATH):
                if os.path.exists(sidecar):
                    os.remove(sidecar)

            shutil.rmtree(db.ASSETS_DIR, ignore_errors=True)
            os.makedirs(db.ASSETS_DIR, exist_ok=True)
            for info in zf.infolist():
                name = info.filename.replace("\\", "/")
                if info.is_dir() or not name.startswith("assets/"):
                    continue
                target = os.path.join(db.ASSETS_DIR, name[len("assets/"):])
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with zf.open(info) as src, open(target, "wb") as out:
                    shutil.copyfileobj(src, out)

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


def list_profile_backups(limit: int = 10) -> list[dict]:
    if not os.path.isdir(BACKUP_DIR):
        return []
    files = sorted(Path(BACKUP_DIR).glob("*.zip"), key=lambda p: p.stat().st_mtime, reverse=True)
    out = []
    for path in files[:limit]:
        stat = path.stat()
        out.append({"path": str(path), "name": path.name, "size": stat.st_size, "mtime": stat.st_mtime})
    return out
