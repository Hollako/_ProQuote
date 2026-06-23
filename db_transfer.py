"""Portable PostgreSQL backups and one-time SQLite-to-PostgreSQL migration."""
from __future__ import annotations

import io
import json
import os
import sqlite3
import zipfile
import base64
import datetime as dt
import mimetypes
from pathlib import Path

import db_postgres


PORTABLE_MANIFEST = "proquote-portable-manifest.json"
PORTABLE_FORMAT = 1
REQUIRED_TABLES = {
    "Projects_Master", "Project_Sheets", "Items_Catalog", "Project_BoQ_Lines",
    "Settings", "Users", "Roles", "RolePerms",
}


def _sqlite_table_names(conn: sqlite3.Connection) -> set[str]:
    return {
        row[0]
        for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }


def _sqlite_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    return [row[1] for row in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]


def _disable_postgres_triggers(conn) -> None:
    for table in db_postgres.TABLES_IN_LOAD_ORDER:
        if table != "Audit_Log":
            conn.execute(f"ALTER TABLE {table} DISABLE TRIGGER USER")


def _enable_postgres_triggers(conn) -> None:
    for table in db_postgres.TABLES_IN_LOAD_ORDER:
        if table != "Audit_Log":
            conn.execute(f"ALTER TABLE {table} ENABLE TRIGGER USER")


def _target_has_data(conn) -> bool:
    for table in ("Projects_Master", "Items_Catalog", "Users", "Audit_Log"):
        if int(conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] or 0):
            return True
    return False


def _clear_postgres(conn) -> None:
    conn.execute(
        "TRUNCATE TABLE " + ",".join(db_postgres.TABLES_IN_DELETE_ORDER)
        + " RESTART IDENTITY CASCADE"
    )


def _insert_batch(conn, table: str, columns: list[str], records: list[tuple]) -> None:
    if not records:
        return
    placeholders = ",".join("?" for _ in columns)
    conn.executemany(
        f"INSERT INTO {table} ({','.join(columns)}) VALUES ({placeholders})",
        records,
    )


def _source_asset_files(sqlite_path: str) -> list[Path]:
    asset_dir = Path(sqlite_path).resolve().parent / "assets"
    return sorted(path for path in asset_dir.rglob("*") if path.is_file()) if asset_dir.is_dir() else []


def _copy_source_assets(target, sqlite_path: str) -> int:
    asset_dir = Path(sqlite_path).resolve().parent / "assets"
    files = _source_asset_files(sqlite_path)
    for path in files:
        key = path.relative_to(asset_dir).as_posix()
        target.execute(
            "INSERT INTO App_Assets(AssetKey,FileName,MimeType,Content,UpdatedAt,Deleted) "
            "VALUES(?,?,?,?,?,0) ON CONFLICT(AssetKey) DO UPDATE SET "
            "FileName=excluded.FileName,MimeType=excluded.MimeType,Content=excluded.Content," 
            "UpdatedAt=excluded.UpdatedAt,Deleted=0",
            (key, path.name, mimetypes.guess_type(path)[0] or "application/octet-stream",
             path.read_bytes(), dt.datetime.fromtimestamp(path.stat().st_mtime).isoformat(
                 timespec="seconds")),
        )
    return len(files)


def _source_valid_ids(source: sqlite3.Connection) -> dict[str, set[int]]:
    projects = {int(row[0]) for row in source.execute("SELECT ProjectID FROM Projects_Master")}
    sheets = {
        int(row[0]) for row in source.execute("SELECT SheetID,ProjectID FROM Project_Sheets")
        if int(row[1]) in projects
    }
    items = {int(row[0]) for row in source.execute("SELECT ItemID FROM Items_Catalog")}
    return {"projects": projects, "sheets": sheets, "items": items}


def _clean_source_record(table: str, record: dict, valid: dict[str, set[int]]) -> tuple[dict | None, int]:
    """Return (clean record, repair count); None means an unreachable orphan is skipped."""
    repaired = 0
    for field, value in record.items():
        if isinstance(value, str) and "\x00" in value:
            record[field] = value.replace("\x00", "")
            repaired += 1
    if table in {"Project_Sheets", "Project_BoQ_Lines", "Finance_Payments", "Finance_Purchases"}:
        if int(record.get("ProjectID") or 0) not in valid["projects"]:
            return None, 0
    if table == "Project_BoQ_Lines":
        if record.get("SheetID") is not None and int(record["SheetID"]) not in valid["sheets"]:
            record["SheetID"] = None
            repaired += 1
        if record.get("ItemID") is not None and int(record["ItemID"]) not in valid["items"]:
            record["ItemID"] = None
            repaired += 1
    return record, repaired


def _copy_sqlite_table(source: sqlite3.Connection, target, table: str, valid,
                       batch_size: int = 1000) -> tuple[int, int, int]:
    columns = _sqlite_columns(source, table)
    if not columns:
        return 0, 0, 0
    cursor = source.execute(f'SELECT {",".join(columns)} FROM "{table}"')
    count, skipped, repaired, batch = 0, 0, 0, []
    for row in cursor:
        record, repaired_links = _clean_source_record(
            table, {column: row[column] for column in columns}, valid
        )
        if record is None:
            skipped += 1
            continue
        repaired += repaired_links
        batch.append(tuple(record[column] for column in columns))
        if len(batch) >= batch_size:
            _insert_batch(target, table, columns, batch)
            count += len(batch)
            batch.clear()
    if batch:
        _insert_batch(target, table, columns, batch)
        count += len(batch)
    return count, skipped, repaired


def migrate_sqlite_to_postgres(sqlite_path: str, database_url: str, *, replace=False,
                               progress=None) -> dict:
    """Copy a complete ProQuote SQLite database into PostgreSQL."""
    if not os.path.isfile(sqlite_path):
        raise FileNotFoundError(sqlite_path)
    source = sqlite3.connect(sqlite_path)
    source.row_factory = sqlite3.Row
    source_tables = _sqlite_table_names(source)
    missing = sorted(REQUIRED_TABLES - source_tables)
    if missing:
        source.close()
        raise ValueError("SQLite source is missing ProQuote tables: " + ", ".join(missing))

    target = db_postgres.init_db(database_url, {
        "username": "migration", "display_name": "SQLite migration"
    })
    copied, skipped, repaired = {}, {}, {}
    try:
        if _target_has_data(target) and not replace:
            raise RuntimeError(
                "PostgreSQL already contains ProQuote data. Re-run with replace=True only "
                "after confirming the target database."
            )
        _disable_postgres_triggers(target)
        if replace:
            _clear_postgres(target)
        valid = _source_valid_ids(source)
        for table in db_postgres.TABLES_IN_LOAD_ORDER:
            if table not in source_tables:
                copied[table], skipped[table], repaired[table] = 0, 0, 0
                continue
            copied[table], skipped[table], repaired[table] = _copy_sqlite_table(
                source, target, table, valid
            )
            if progress:
                progress(table, copied[table])
        if "App_Assets" not in source_tables:
            copied["App_Assets"] = _copy_source_assets(target, sqlite_path)
            if progress:
                progress("App_Assets", copied["App_Assets"])
        db_postgres.reset_identity_sequences(target)
        _enable_postgres_triggers(target)
        target.commit()
    except Exception:
        target.rollback()
        raise
    finally:
        source.close()
        target.close()

    verification = verify_sqlite_against_postgres(sqlite_path, database_url)
    if not verification["ok"]:
        raise RuntimeError("Migration count verification failed: " + repr(verification["differences"]))
    return {
        "copied": copied,
        "skipped_orphans": {key: value for key, value in skipped.items() if value},
        "repaired_values": {key: value for key, value in repaired.items() if value},
        "verification": verification,
    }


def verify_sqlite_against_postgres(sqlite_path: str, database_url: str) -> dict:
    source = sqlite3.connect(sqlite_path)
    target = db_postgres.connect(database_url, {
        "username": "migration", "display_name": "Migration verification"
    })
    differences = {}
    try:
        source_tables = _sqlite_table_names(source)
        valid = _source_valid_ids(source)
        for table in db_postgres.TABLES_IN_LOAD_ORDER:
            if table == "App_Assets" and table not in source_tables:
                source_count = len(_source_asset_files(sqlite_path))
            elif table not in source_tables:
                source_count = 0
            elif table in {"Project_Sheets", "Project_BoQ_Lines", "Finance_Payments", "Finance_Purchases"}:
                source_count = sum(
                    1 for row in source.execute(f'SELECT ProjectID FROM "{table}"')
                    if int(row[0] or 0) in valid["projects"]
                )
            else:
                source_count = int(source.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
            target_count = int(target.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
            if source_count != target_count:
                differences[table] = {"sqlite": source_count, "postgres": target_count}
        source_total = source.execute(
            "SELECT COALESCE(SUM(l.TPriceSAR),0) FROM Project_BoQ_Lines l "
            "JOIN Projects_Master p ON p.ProjectID=l.ProjectID"
        ).fetchone()[0]
        target_total = target.execute(
            "SELECT COALESCE(SUM(TPriceSAR),0) FROM Project_BoQ_Lines"
        ).fetchone()[0]
        if abs(float(source_total or 0) - float(target_total or 0)) > 0.01:
            differences["TPriceSAR"] = {"sqlite": source_total, "postgres": target_total}
    finally:
        source.close()
        target.close()
    return {"ok": not differences, "differences": differences}


def write_portable_backup(zip_path: str, conn, asset_paths=()) -> dict:
    """Write database rows and assets to a backend-neutral ZIP profile backup."""
    manifest = {"format": PORTABLE_FORMAT, "backend": "postgresql", "tables": {}}
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for table in db_postgres.TABLES_IN_LOAD_ORDER:
            rows = conn.execute(f"SELECT * FROM {table}")
            entry = f"database/{table}.jsonl"
            count = 0
            with zf.open(entry, "w") as raw:
                text = io.TextIOWrapper(raw, encoding="utf-8", newline="\n")
                for row in rows:
                    record = {
                        key: ({"__proquote_bytes__": base64.b64encode(value).decode("ascii")}
                              if isinstance(value, (bytes, bytearray, memoryview)) else value)
                        for key, value in dict(row).items()
                    }
                    text.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")
                    count += 1
                text.flush()
                text.detach()
            manifest["tables"][table] = {"entry": entry, "rows": count}
        for path, archive_name in asset_paths:
            zf.write(path, archive_name)
        zf.writestr(PORTABLE_MANIFEST, json.dumps(manifest, indent=2))
    return manifest


def read_portable_manifest(zf: zipfile.ZipFile) -> dict:
    if PORTABLE_MANIFEST not in zf.namelist():
        raise ValueError("Backup is not a portable ProQuote PostgreSQL profile.")
    manifest = json.loads(zf.read(PORTABLE_MANIFEST).decode("utf-8"))
    if manifest.get("format") != PORTABLE_FORMAT or not isinstance(manifest.get("tables"), dict):
        raise ValueError("Unsupported ProQuote portable backup format.")
    missing = sorted(REQUIRED_TABLES - set(manifest["tables"]))
    if missing:
        raise ValueError("Portable backup is missing tables: " + ", ".join(missing))
    return manifest


def restore_portable_backup(zip_path: str, database_url: str, *, progress=None) -> dict:
    target = db_postgres.init_db(database_url, {
        "username": "restore", "display_name": "Profile restore"
    })
    restored = {}
    try:
        with zipfile.ZipFile(zip_path) as zf:
            manifest = read_portable_manifest(zf)
            _disable_postgres_triggers(target)
            _clear_postgres(target)
            for table in db_postgres.TABLES_IN_LOAD_ORDER:
                meta = manifest["tables"].get(table)
                if not meta:
                    restored[table] = 0
                    continue
                batch, columns, count = [], None, 0
                with zf.open(meta["entry"]) as raw:
                    for raw_line in raw:
                        record = json.loads(raw_line.decode("utf-8"))
                        record = {
                            key: (base64.b64decode(value["__proquote_bytes__"])
                                  if isinstance(value, dict) and "__proquote_bytes__" in value
                                  else value)
                            for key, value in record.items()
                        }
                        if columns is None:
                            columns = list(record)
                        batch.append(tuple(record.get(column) for column in columns))
                        if len(batch) >= 1000:
                            _insert_batch(target, table, columns, batch)
                            count += len(batch)
                            batch.clear()
                    if batch and columns:
                        _insert_batch(target, table, columns, batch)
                        count += len(batch)
                restored[table] = count
                if count != int(meta.get("rows", count)):
                    raise RuntimeError(f"Row count mismatch while restoring {table}.")
                if progress:
                    progress(table, count)
            db_postgres.reset_identity_sequences(target)
            _enable_postgres_triggers(target)
            target.commit()
    except Exception:
        target.rollback()
        raise
    finally:
        target.close()
    return restored
