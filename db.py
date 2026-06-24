"""ProQuote database layer with SQLite fallback and PostgreSQL production support."""
import os
import math
import sqlite3
import contextvars
import datetime as dt
import mimetypes

# The database + assets live in a per-company DATA DIRECTORY. Set the env var
# BOQ_DATA_DIR to give each company its own DB + logo/banner (Model A: one shared
# codebase, one data profile per company). Defaults to the app folder (backward
# compatible with the original single-company setup).
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("BOQ_DATA_DIR", "").strip() or APP_DIR
ASSETS_DIR = os.path.join(DATA_DIR, "assets")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "proquote.db")


def database_url() -> str:
    return os.environ.get("DATABASE_URL", "").strip()


def is_postgres() -> bool:
    return database_url().lower().startswith(("postgres://", "postgresql://"))

if "_AUDIT_ACTOR" not in globals():
    _AUDIT_ACTOR = contextvars.ContextVar(
        "proquote_audit_actor",
        default={"user_id": None, "username": "system", "display_name": "System"},
    )


def set_audit_actor(user: dict | None) -> None:
    """Attach the authenticated user to subsequent DB writes in this execution context."""
    user = user or {}
    _AUDIT_ACTOR.set({
        "user_id": user.get("UserID"),
        "username": str(user.get("Username") or "system"),
        "display_name": str(user.get("DisplayName") or user.get("Username") or "System"),
    })


def clear_audit_actor() -> None:
    set_audit_actor(None)


def get_audit_actor() -> dict:
    return dict(_AUDIT_ACTOR.get())


def banner_path() -> str:
    """Path to this company's full-width banner image (PNG)."""
    return os.path.join(ASSETS_DIR, "header_banner.png")


def logo_path() -> str:
    """Path to this company's standalone logo mark (PNG)."""
    return os.path.join(ASSETS_DIR, "logo.png")


def header_left_path() -> str:
    """Path to this company's left header section image (PNG)."""
    return os.path.join(ASSETS_DIR, "header_left.png")


def header_middle_path() -> str:
    """Path to this company's middle header section image (PNG)."""
    return os.path.join(ASSETS_DIR, "header_middle.png")


def header_right_path() -> str:
    """Path to this company's right header section image (PNG)."""
    return os.path.join(ASSETS_DIR, "header_right.png")


def footer_full_path() -> str:
    """Path to this company's full-width footer image (PNG)."""
    return os.path.join(ASSETS_DIR, "footer_full.png")


def footer_left_path() -> str:
    """Path to this company's left footer section image (PNG)."""
    return os.path.join(ASSETS_DIR, "footer_left.png")


def footer_middle_path() -> str:
    """Path to this company's middle footer section image (PNG)."""
    return os.path.join(ASSETS_DIR, "footer_middle.png")


def footer_right_path() -> str:
    """Path to this company's right footer section image (PNG)."""
    return os.path.join(ASSETS_DIR, "footer_right.png")


def _asset_key(path: str) -> str:
    return os.path.relpath(os.path.abspath(path), ASSETS_DIR).replace("\\", "/")


def save_asset(path: str, content: bytes, mime_type: str | None = None) -> None:
    """Persist an app branding asset locally and, under PostgreSQL, in the database."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as out:
        out.write(content)
    if is_postgres():
        key = _asset_key(path)
        with connect() as conn:
            conn.execute(
                "INSERT INTO App_Assets(AssetKey,FileName,MimeType,Content,UpdatedAt,Deleted) "
                "VALUES(?,?,?,?,?,0) ON CONFLICT(AssetKey) DO UPDATE SET "
                "FileName=excluded.FileName,MimeType=excluded.MimeType,Content=excluded.Content," 
                "UpdatedAt=excluded.UpdatedAt,Deleted=0",
                (key, os.path.basename(path), mime_type or mimetypes.guess_type(path)[0]
                 or "application/octet-stream", bytes(content),
                 dt.datetime.now().isoformat(timespec="seconds")),
            )
            conn.commit()


def delete_asset(path: str) -> None:
    """Remove an asset and persist the deletion across cloud container restarts."""
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    if is_postgres():
        key = _asset_key(path)
        with connect() as conn:
            conn.execute(
                "INSERT INTO App_Assets(AssetKey,FileName,MimeType,Content,UpdatedAt,Deleted) "
                "VALUES(?,?,?,?,?,1) ON CONFLICT(AssetKey) DO UPDATE SET "
                "Content=excluded.Content,UpdatedAt=excluded.UpdatedAt,Deleted=1",
                (key, os.path.basename(path), mimetypes.guess_type(path)[0]
                 or "application/octet-stream", b"",
                 dt.datetime.now().isoformat(timespec="seconds")),
            )
            conn.commit()


def _sync_postgres_assets(conn) -> None:
    rows = conn.execute(
        "SELECT AssetKey,Content,Deleted FROM App_Assets ORDER BY AssetKey"
    ).fetchall()
    if not rows:
        for root, _dirs, files in os.walk(ASSETS_DIR):
            for filename in files:
                path = os.path.join(root, filename)
                with open(path, "rb") as src:
                    content = src.read()
                key = _asset_key(path)
                conn.execute(
                    "INSERT INTO App_Assets(AssetKey,FileName,MimeType,Content,UpdatedAt,Deleted) "
                    "VALUES(?,?,?,?,?,0) ON CONFLICT(AssetKey) DO NOTHING",
                    (key, filename, mimetypes.guess_type(path)[0] or "application/octet-stream",
                     content, dt.datetime.now().isoformat(timespec="seconds")),
                )
        conn.commit()
        return
    for row in rows:
        path = os.path.join(ASSETS_DIR, str(row["AssetKey"]).replace("/", os.sep))
        if row["Deleted"]:
            try:
                os.remove(path)
            except FileNotFoundError:
                pass
            continue
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as out:
            out.write(bytes(row["Content"]))

SCHEMA = r"""
PRAGMA foreign_keys = ON;

-- One row per source Excel file (a "project / offer" document).
CREATE TABLE IF NOT EXISTS Projects_Master (
    ProjectID        INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectName      TEXT,
    ClientName       TEXT,
    ContactName      TEXT,
    ContactPhone     TEXT,
    Contractor       TEXT,
    Region           TEXT,
    SalesPerson      TEXT,            -- sales person who worked on the offer
    PresalesEngineer TEXT,            -- pre-sales engineer who worked on the offer
    ProjectManager   TEXT,            -- project manager assigned to the offer
    OfferNo          TEXT,
    CreationDate     TEXT,            -- ISO date (from Quotation 'Date:' or file mtime)
    UpdatedDate      TEXT,            -- ISO date; changes whenever the offer is modified
    DiscountAmount   REAL DEFAULT 0,  -- primary discount (first system sheet)
    CommissionAmount REAL DEFAULT 0,  -- internal expense; excluded from client totals and profit
    CommissionPercent REAL DEFAULT 0, -- percentage gross-up applied to item margins
    CommissionMode TEXT DEFAULT 'Deduct from profit',
    ConversionFactor REAL,            -- primary factor (e.g. 1.69)
    SourceFile       TEXT UNIQUE,     -- absolute path; used for idempotent re-ingest
    IngestedAt       TEXT,
    RevisionNo       INTEGER DEFAULT 0,  -- 0 = original; 1,2,... = saved revisions
    BaseName         TEXT,            -- groups an offer + its revisions together
    OfferTerms       TEXT,            -- JSON: subject, greeting, scope, payment, ... (Quotation notes)
    ProjectSheetInfo TEXT,            -- JSON: project-sheet export details
    Approved         INTEGER DEFAULT 0,  -- 1 = this is the approved revision/option
    ApprovedAt       TEXT,            -- ISO timestamp of approval
    OptionLabel      TEXT,            -- alternative within a revision (e.g. Dynalite / KNX)
    Archived         INTEGER DEFAULT 0,  -- 1 = soft-deleted (hidden, restorable)
    ArchivedBy       INTEGER          -- ProjectID whose approval auto-archived this (for auto-restore)
);

-- One row per BOQ <System> sheet inside a file (a file can hold LCS, ELV, BGM ...).
CREATE TABLE IF NOT EXISTS Project_Sheets (
    SheetID          INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectID        INTEGER NOT NULL REFERENCES Projects_Master(ProjectID) ON DELETE CASCADE,
    SheetName        TEXT,            -- e.g. 'BOQ LCS'
    SystemSuffix     TEXT,            -- e.g. 'LCS'
    DiscountAmount   REAL DEFAULT 0,
    Factor1          REAL,            -- the standalone conversion factors found
    Factor2          REAL,            -- under the totals block (e.g. 1.69 / 1.49 / 1.42)
    Factor3          REAL,
    SubtotalSAR      REAL,
    GrandTotalSAR    REAL
);

-- Deduplicated master catalogue of every distinct item ever quoted.
CREATE TABLE IF NOT EXISTS Items_Catalog (
    ItemID           INTEGER PRIMARY KEY AUTOINCREMENT,
    Description      TEXT,
    Brand            TEXT,
    Model            TEXT,
    ListPriceUSD     REAL,
    ExUnitCostUSD    REAL,
    Currency         TEXT DEFAULT 'USD',  -- currency of List Price / Ex Unit Cost
    ShippingPercent  REAL DEFAULT 30,
    UnitCostUSD      REAL,
    DefaultUPriceUSD REAL,
    DefaultUPriceSAR REAL,
    PriceUpdatedAt   TEXT DEFAULT '2025-01-01',
    TimesQuoted      INTEGER DEFAULT 0,
    LastSeenFile     TEXT,
    LastSeenAt       TEXT,
    UNIQUE (Brand, Model, Description)
);

-- Every line item of every project (the full internal BOQ grid).
CREATE TABLE IF NOT EXISTS Project_BoQ_Lines (
    LineID           INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectID        INTEGER NOT NULL REFERENCES Projects_Master(ProjectID) ON DELETE CASCADE,
    SheetID          INTEGER REFERENCES Project_Sheets(SheetID) ON DELETE CASCADE,
    ItemID           INTEGER REFERENCES Items_Catalog(ItemID),
    RowOrder         INTEGER,         -- preserves original sheet order
    Area             TEXT,
    System           TEXT,
    Description      TEXT,
    Brand            TEXT,
    Model            TEXT,
    Qty              REAL,
    ListPriceUSD     REAL,
    ExUnitCostUSD    REAL,
    Currency         TEXT DEFAULT 'USD',  -- currency of List Price / Ex Unit Cost
    ShippingPercent  REAL DEFAULT 30,
    FinalUnitCostUSD REAL,
    TotalCostUSD     REAL,
    FinalUPriceUSD   REAL,
    TPriceUSD        REAL,
    FinalUPriceSAR   REAL,
    TPriceSAR        REAL,
    MarginExtra      REAL,            -- persisted per-line pricing multiplier (Markup x)
    LineType         TEXT DEFAULT 'item',  -- 'item' | 'discount' | 'service'
    Paid             INTEGER DEFAULT 0,   -- supplier paid (procurement tracking)
    Received         INTEGER DEFAULT 0,   -- received from supplier
    ReceivedQty      REAL DEFAULT 0,      -- quantity received from supplier
    Delivered        INTEGER DEFAULT 0,   -- delivered to site
    DeliveredQty     REAL DEFAULT 0,      -- quantity delivered to site
    PONumber         TEXT,            -- purchase-order number (from Accounting)
    DeliveryNote     TEXT,            -- delivery note reference / free text
    PaidAt           TEXT,            -- ISO timestamp when Paid was first checked
    ReceivedAt       TEXT,            -- ISO timestamp when Received was first checked
    DeliveredAt      TEXT             -- ISO timestamp when Delivered was first checked
);

-- Application settings (key/value): offer prefix, number padding, type codes, ...
CREATE TABLE IF NOT EXISTS Settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);

-- Application users (login + role-based access).
CREATE TABLE IF NOT EXISTS Users (
    UserID       INTEGER PRIMARY KEY AUTOINCREMENT,
    Username     TEXT UNIQUE NOT NULL,
    DisplayName  TEXT,
    PasswordHash TEXT NOT NULL,
    Role         TEXT NOT NULL DEFAULT 'viewer',
    Active       INTEGER DEFAULT 1,
    CreatedAt    TEXT
);

-- Configurable roles + their granted permissions (editable matrix).
CREATE TABLE IF NOT EXISTS Roles ( Role TEXT PRIMARY KEY );
CREATE TABLE IF NOT EXISTS RolePerms (
    Role       TEXT NOT NULL,
    Permission TEXT NOT NULL,
    PRIMARY KEY (Role, Permission)
);

-- Finance: client payments/invoices billed against an approved offer.
CREATE TABLE IF NOT EXISTS Finance_Payments (
    PayID       INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectID   INTEGER NOT NULL REFERENCES Projects_Master(ProjectID) ON DELETE CASCADE,
    RowOrder    INTEGER,
    Description TEXT,
    AmountSAR   REAL DEFAULT 0,
    InvoiceNo   TEXT             -- invoice number (plain text)
);

-- Finance: purchases/costs (materials, suppliers) spent on the project.
CREATE TABLE IF NOT EXISTS Finance_Purchases (
    PurID       INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectID   INTEGER NOT NULL REFERENCES Projects_Master(ProjectID) ON DELETE CASCADE,
    RowOrder    INTEGER,
    Description TEXT,
    AmountSAR   REAL DEFAULT 0,
    PORef       TEXT             -- purchase-order reference (plain text)
);

-- Immutable application audit trail. Triggers are generated after migrations so
-- their before/after JSON always follows the current table columns.
CREATE TABLE IF NOT EXISTS Audit_Log (
    AuditID      INTEGER PRIMARY KEY AUTOINCREMENT,
    EventAt      TEXT NOT NULL,
    UserID       INTEGER,
    Username     TEXT NOT NULL,
    DisplayName  TEXT,
    Action       TEXT NOT NULL,
    EntityType   TEXT NOT NULL,
    EntityID     TEXT,
    Summary      TEXT,
    OldValues    TEXT,
    NewValues    TEXT
);

CREATE INDEX IF NOT EXISTS idx_catalog_model ON Items_Catalog(Model);
CREATE INDEX IF NOT EXISTS idx_catalog_desc  ON Items_Catalog(Description);
CREATE INDEX IF NOT EXISTS idx_projects_offer_no ON Projects_Master(OfferNo);
CREATE INDEX IF NOT EXISTS idx_projects_archived ON Projects_Master(Archived);
CREATE INDEX IF NOT EXISTS idx_sheets_project ON Project_Sheets(ProjectID);
CREATE INDEX IF NOT EXISTS idx_lines_project ON Project_BoQ_Lines(ProjectID);
CREATE INDEX IF NOT EXISTS idx_lines_item    ON Project_BoQ_Lines(ItemID);
CREATE INDEX IF NOT EXISTS idx_lines_type_project ON Project_BoQ_Lines(LineType, ProjectID);
CREATE INDEX IF NOT EXISTS idx_fin_pay_project ON Finance_Payments(ProjectID);
CREATE INDEX IF NOT EXISTS idx_fin_pur_project ON Finance_Purchases(ProjectID);
CREATE INDEX IF NOT EXISTS idx_audit_event_at ON Audit_Log(EventAt DESC);
CREATE INDEX IF NOT EXISTS idx_audit_user ON Audit_Log(Username, EventAt DESC);
CREATE INDEX IF NOT EXISTS idx_audit_entity ON Audit_Log(EntityType, EntityID, EventAt DESC);
"""


def connect_sqlite(db_path: str = DB_PATH) -> sqlite3.Connection:
    # timeout = how long a blocked writer waits for the lock before erroring.
    conn = sqlite3.connect(db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")       # concurrent readers + 1 writer
    conn.execute("PRAGMA busy_timeout = 30000;")     # 30s wait-and-retry on lock
    conn.execute("PRAGMA synchronous = NORMAL;")     # safe with WAL, much faster
    conn.create_function("audit_user_id", 0, lambda: _AUDIT_ACTOR.get().get("user_id"))
    conn.create_function("audit_username", 0, lambda: _AUDIT_ACTOR.get().get("username"))
    conn.create_function(
        "audit_display_name", 0, lambda: _AUDIT_ACTOR.get().get("display_name")
    )
    return conn


def connect(db_path: str = DB_PATH):
    if is_postgres():
        import db_postgres

        return db_postgres.connect(database_url(), get_audit_actor())
    return connect_sqlite(db_path)


# Columns added after the first release - applied to pre-existing databases.
MIGRATIONS = {
    "Projects_Master": {
        "SalesPerson": "TEXT",
        "ContactPhone": "TEXT",
        "Contractor": "TEXT",
        "Region": "TEXT",
        "PresalesEngineer": "TEXT",
        "ProjectManager": "TEXT",
        "CommissionAmount": "REAL DEFAULT 0",
        "CommissionPercent": "REAL DEFAULT 0",
        "CommissionMode": "TEXT DEFAULT 'Deduct from profit'",
        "RevisionNo": "INTEGER DEFAULT 0",
        "BaseName": "TEXT",
        "OfferTerms": "TEXT",
        "ProjectSheetInfo": "TEXT",
        "Approved": "INTEGER DEFAULT 0",
        "ApprovedAt": "TEXT",
        "OptionLabel": "TEXT",
        "Archived": "INTEGER DEFAULT 0",
        "ArchivedBy": "INTEGER",
        "UpdatedDate": "TEXT",
        "InclusionMode": "TEXT",
        "InclusionMarkup": "REAL",
    },
    "Project_BoQ_Lines": {
        "ShippingPercent": "REAL",
        "Currency": "TEXT DEFAULT 'USD'",
        "Paid": "INTEGER DEFAULT 0",
        "Received": "INTEGER DEFAULT 0",
        "ReceivedQty": "REAL DEFAULT 0",
        "Delivered": "INTEGER DEFAULT 0",
        "DeliveredQty": "REAL DEFAULT 0",
        "PONumber": "TEXT",
        "DeliveryNote": "TEXT",
        "PaidAt": "TEXT",
        "ReceivedAt": "TEXT",
        "DeliveredAt": "TEXT",
        "ReceivedRegion": "TEXT",
        "IncludedInItems": "INTEGER DEFAULT 0",
    },
    "Items_Catalog": {
        "ShippingPercent": "REAL",
        "Currency": "TEXT DEFAULT 'USD'",
        "PriceUpdatedAt": "TEXT DEFAULT '2025-01-01'",
        "Discontinued": "INTEGER DEFAULT 0",
    },
}


def _migrate(conn: sqlite3.Connection) -> None:
    for table, cols in MIGRATIONS.items():
        existing = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
        for col, decl in cols.items():
            if col not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")
    conn.commit()


def _backfill_tracking_quantities(conn: sqlite3.Connection) -> None:
    existing = {r["name"] for r in conn.execute("PRAGMA table_info(Project_BoQ_Lines)")}
    if not {"ReceivedQty", "DeliveredQty"}.issubset(existing):
        return
    conn.execute(
        """UPDATE Project_BoQ_Lines
              SET ReceivedQty=COALESCE(Qty, 0)
            WHERE IFNULL(Received,0)=1
              AND IFNULL(ReceivedQty,0)=0"""
    )
    conn.execute(
        """UPDATE Project_BoQ_Lines
              SET DeliveredQty=COALESCE(Qty, 0)
            WHERE IFNULL(Delivered,0)=1
              AND IFNULL(DeliveredQty,0)=0"""
    )
    conn.execute(
        """UPDATE Project_BoQ_Lines
              SET Received=1
            WHERE IFNULL(ReceivedQty,0)>0
              AND IFNULL(Received,0)=0"""
    )
    conn.execute(
        """UPDATE Project_BoQ_Lines
              SET Delivered=1
            WHERE IFNULL(DeliveredQty,0)>0
              AND IFNULL(Delivered,0)=0"""
    )
    conn.commit()


def _backfill_catalog_price_dates(conn: sqlite3.Connection) -> None:
    existing = {r["name"] for r in conn.execute("PRAGMA table_info(Items_Catalog)")}
    if "PriceUpdatedAt" not in existing:
        return
    conn.execute(
        "UPDATE Items_Catalog SET PriceUpdatedAt='2025-01-01' "
        "WHERE PriceUpdatedAt IS NULL OR TRIM(PriceUpdatedAt)=''"
    )
    conn.commit()


def _backfill_project_updated_dates(conn: sqlite3.Connection) -> None:
    """Existing offer dates become both the original creation and initial update date."""
    existing = {r["name"] for r in conn.execute("PRAGMA table_info(Projects_Master)")}
    if not {"CreationDate", "UpdatedDate"}.issubset(existing):
        return
    conn.execute(
        "UPDATE Projects_Master SET UpdatedDate=CreationDate "
        "WHERE UpdatedDate IS NULL OR TRIM(UpdatedDate)=''"
    )
    conn.commit()


def _ensure_project_date_triggers(conn: sqlite3.Connection) -> None:
    """Keep UpdatedDate current for offer-header, sheet, and line modifications."""
    conn.executescript(
        """
        CREATE TRIGGER IF NOT EXISTS trg_projects_master_updated_date
        AFTER UPDATE ON Projects_Master
        FOR EACH ROW
        WHEN COALESCE(NEW.UpdatedDate,'') = COALESCE(OLD.UpdatedDate,'')
        BEGIN
            UPDATE Projects_Master
               SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=NEW.ProjectID;
        END;

        CREATE TRIGGER IF NOT EXISTS trg_project_sheets_insert_updated_date
        AFTER INSERT ON Project_Sheets
        BEGIN
            UPDATE Projects_Master SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=NEW.ProjectID;
        END;
        CREATE TRIGGER IF NOT EXISTS trg_project_sheets_update_updated_date
        AFTER UPDATE ON Project_Sheets
        BEGIN
            UPDATE Projects_Master SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=NEW.ProjectID;
        END;
        CREATE TRIGGER IF NOT EXISTS trg_project_sheets_delete_updated_date
        AFTER DELETE ON Project_Sheets
        BEGIN
            UPDATE Projects_Master SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=OLD.ProjectID;
        END;

        CREATE TRIGGER IF NOT EXISTS trg_project_lines_insert_updated_date
        AFTER INSERT ON Project_BoQ_Lines
        BEGIN
            UPDATE Projects_Master SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=NEW.ProjectID;
        END;
        CREATE TRIGGER IF NOT EXISTS trg_project_lines_update_updated_date
        AFTER UPDATE ON Project_BoQ_Lines
        BEGIN
            UPDATE Projects_Master SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=NEW.ProjectID;
        END;
        CREATE TRIGGER IF NOT EXISTS trg_project_lines_delete_updated_date
        AFTER DELETE ON Project_BoQ_Lines
        BEGIN
            UPDATE Projects_Master SET UpdatedDate=date('now','localtime')
             WHERE ProjectID=OLD.ProjectID;
        END;
        """
    )
    conn.commit()


_AUDITED_TABLES = (
    "Projects_Master", "Project_Sheets", "Project_BoQ_Lines", "Items_Catalog",
    "Finance_Payments", "Finance_Purchases", "Settings", "Users", "Roles", "RolePerms",
)
_AUDIT_REDACTED_COLUMNS = {"Users": {"PasswordHash"}}


def _sql_ident(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _sql_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _audit_json_expr(table: str, prefix: str, columns: list[str]) -> str:
    parts = []
    redacted = _AUDIT_REDACTED_COLUMNS.get(table, set())
    for column in columns:
        parts.append(_sql_literal(column))
        parts.append(
            _sql_literal("[REDACTED]")
            if column in redacted else f"{prefix}.{_sql_ident(column)}"
        )
    return f"json_object({','.join(parts)})"


def _audit_entity_id_expr(prefix: str, pk_columns: list[str]) -> str:
    if not pk_columns:
        return f"CAST({prefix}.rowid AS TEXT)"
    pieces = [
        f"{_sql_literal(column + '=')} || COALESCE(CAST({prefix}.{_sql_ident(column)} AS TEXT),'')"
        for column in pk_columns
    ]
    return " || '; ' || ".join(pieces)


def _ensure_audit_triggers(conn: sqlite3.Connection) -> None:
    """Generate immutable row-level auditing for every mutable business table."""
    existing_tables = {
        r["name"] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    for table in _AUDITED_TABLES:
        if table not in existing_tables:
            continue
        info = conn.execute(f"PRAGMA table_info({_sql_ident(table)})").fetchall()
        columns = [r["name"] for r in info]
        pk_columns = [r["name"] for r in sorted(info, key=lambda r: r["pk"]) if r["pk"]]
        table_ident = _sql_ident(table)
        table_lit = _sql_literal(table)
        old_json = _audit_json_expr(table, "OLD", columns)
        new_json = _audit_json_expr(table, "NEW", columns)
        old_id = _audit_entity_id_expr("OLD", pk_columns)
        new_id = _audit_entity_id_expr("NEW", pk_columns)
        compared = [c for c in columns if not (table == "Projects_Master" and c == "UpdatedDate")]
        update_when = " OR ".join(
            f"OLD.{_sql_ident(c)} IS NOT NEW.{_sql_ident(c)}" for c in compared
        ) or "0"

        for action in ("insert", "update", "delete"):
            conn.execute(f"DROP TRIGGER IF EXISTS audit_{table}_{action}")

        common_cols = (
            "EventAt,UserID,Username,DisplayName,Action,EntityType,EntityID,Summary,"
            "OldValues,NewValues"
        )
        actor_sql = "audit_user_id(),audit_username(),audit_display_name()"
        event_sql = "strftime('%Y-%m-%dT%H:%M:%S','now','localtime')"
        conn.executescript(
            f"""
            CREATE TRIGGER audit_{table}_insert AFTER INSERT ON {table_ident}
            BEGIN
                INSERT INTO Audit_Log({common_cols})
                VALUES ({event_sql},{actor_sql},'INSERT',{table_lit},{new_id},
                        'Created ' || {table_lit},NULL,{new_json});
            END;
            CREATE TRIGGER audit_{table}_update AFTER UPDATE ON {table_ident}
            WHEN {update_when}
            BEGIN
                INSERT INTO Audit_Log({common_cols})
                VALUES ({event_sql},{actor_sql},'UPDATE',{table_lit},{new_id},
                        'Updated ' || {table_lit},{old_json},{new_json});
            END;
            CREATE TRIGGER audit_{table}_delete AFTER DELETE ON {table_ident}
            BEGIN
                INSERT INTO Audit_Log({common_cols})
                VALUES ({event_sql},{actor_sql},'DELETE',{table_lit},{old_id},
                        'Deleted ' || {table_lit},{old_json},NULL);
            END;
            """
        )
    conn.commit()


def _backfill_imported_margins(conn: sqlite3.Connection) -> None:
    """One-time repair: replace imported trailing-cell values with price/cost margin."""
    migration_key = "migration_imported_margin_v2"
    done = conn.execute("SELECT value FROM Settings WHERE key=?", (migration_key,)).fetchone()
    if done and done["value"] == "1":
        return
    rows = conn.execute(
        """SELECT l.LineID,l.FinalUnitCostUSD,l.TotalCostUSD,l.FinalUPriceUSD,l.TPriceUSD
             FROM Project_BoQ_Lines l
             JOIN Projects_Master p ON p.ProjectID=l.ProjectID
            WHERE l.LineType NOT IN ('discount','spare')
              AND IFNULL(p.SourceFile,'') NOT LIKE 'app://%'"""
    ).fetchall()
    updates = []
    for row in rows:
        unit_cost = row["FinalUnitCostUSD"]
        unit_price = row["FinalUPriceUSD"]
        total_cost = row["TotalCostUSD"]
        total_price = row["TPriceUSD"]
        editor_cost = math.ceil(float(unit_cost)) if unit_cost is not None and unit_cost > 0 else 0
        if editor_cost > 0 and unit_price is not None:
            margin = round(max(float(unit_price) / editor_cost, 0.0), 4)
        elif total_cost is not None and total_cost > 0 and total_price is not None:
            margin = round(max(float(total_price) / float(total_cost), 0.0), 4)
        else:
            margin = 0.0
        updates.append((margin, row["LineID"]))
    conn.executemany(
        "UPDATE Project_BoQ_Lines SET MarginExtra=? WHERE LineID=?",
        updates,
    )
    conn.execute(
        "INSERT OR REPLACE INTO Settings(key,value) VALUES (?, '1')",
        (migration_key,),
    )
    conn.commit()


def init_db(db_path: str = DB_PATH):
    if is_postgres():
        import db_postgres

        conn = db_postgres.init_db(database_url(), get_audit_actor())
        _sync_postgres_assets(conn)
        return conn
    conn = connect(db_path)
    conn.executescript(SCHEMA)
    _migrate(conn)
    _backfill_tracking_quantities(conn)
    _backfill_catalog_price_dates(conn)
    _backfill_project_updated_dates(conn)
    _ensure_project_date_triggers(conn)
    _ensure_audit_triggers(conn)
    _backfill_imported_margins(conn)
    conn.commit()
    return conn


if __name__ == "__main__":
    c = init_db()
    if is_postgres():
        print("Initialized PostgreSQL schema.")
        rows = c.execute(
            "SELECT table_name AS name FROM information_schema.tables "
            "WHERE table_schema='public' ORDER BY table_name"
        )
    else:
        print(f"Initialized schema at: {DB_PATH}")
        rows = c.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    for row in rows:
        name = row["name"]
        print("  table:", name)
    c.close()
