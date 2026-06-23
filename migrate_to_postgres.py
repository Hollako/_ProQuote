"""One-time ProQuote SQLite to PostgreSQL migration command."""
from __future__ import annotations

import argparse
import os
import tomllib
from pathlib import Path

import db
import db_transfer


def migration_url() -> str:
    for key in ("POSTGRES_MIGRATION_URL", "DATABASE_URL"):
        value = os.environ.get(key, "").strip()
        if value:
            return value
    secrets_path = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"
    if secrets_path.is_file():
        values = tomllib.loads(secrets_path.read_text(encoding="utf-8"))
        for key in ("POSTGRES_MIGRATION_URL", "DATABASE_URL"):
            value = str(values.get(key, "")).strip()
            if value:
                return value
    return ""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sqlite", default=db.DB_PATH, help="Source ProQuote SQLite file")
    parser.add_argument(
        "--replace", action="store_true",
        help="Replace data already present in the target PostgreSQL database",
    )
    args = parser.parse_args()
    database_url = migration_url()
    if not database_url:
        raise SystemExit(
            "Set POSTGRES_MIGRATION_URL to the direct PostgreSQL URL in the environment "
            "or .streamlit/secrets.toml first."
        )

    def progress(table, count):
        print(f"  {table}: {count:,}")

    print(f"Migrating: {args.sqlite}")
    result = db_transfer.migrate_sqlite_to_postgres(
        args.sqlite, database_url, replace=args.replace, progress=progress
    )
    print("Migration verified successfully.")
    print(result["verification"])


if __name__ == "__main__":
    main()
