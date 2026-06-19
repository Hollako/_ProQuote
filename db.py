"""
ProQuote System - SQLite schema & connection layer.

Single-file relational database for the centralized Bill of Quantities engine.
All money is stored exactly as found in the source sheets (no rounding on ingest).
"""
import os
import sqlite3

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
    DiscountAmount   REAL DEFAULT 0,  -- primary discount (first system sheet)
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
    MarginExtra      REAL,            -- the trailing analytic column (col O+)
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

CREATE INDEX IF NOT EXISTS idx_catalog_model ON Items_Catalog(Model);
CREATE INDEX IF NOT EXISTS idx_catalog_desc  ON Items_Catalog(Description);
CREATE INDEX IF NOT EXISTS idx_projects_archived ON Projects_Master(Archived);
CREATE INDEX IF NOT EXISTS idx_sheets_project ON Project_Sheets(ProjectID);
CREATE INDEX IF NOT EXISTS idx_lines_project ON Project_BoQ_Lines(ProjectID);
CREATE INDEX IF NOT EXISTS idx_lines_item    ON Project_BoQ_Lines(ItemID);
CREATE INDEX IF NOT EXISTS idx_lines_type_project ON Project_BoQ_Lines(LineType, ProjectID);
CREATE INDEX IF NOT EXISTS idx_fin_pay_project ON Finance_Payments(ProjectID);
CREATE INDEX IF NOT EXISTS idx_fin_pur_project ON Finance_Purchases(ProjectID);
"""


def connect(db_path: str = DB_PATH) -> sqlite3.Connection:
    # timeout = how long a blocked writer waits for the lock before erroring.
    conn = sqlite3.connect(db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")       # concurrent readers + 1 writer
    conn.execute("PRAGMA busy_timeout = 30000;")     # 30s wait-and-retry on lock
    conn.execute("PRAGMA synchronous = NORMAL;")     # safe with WAL, much faster
    return conn


# Columns added after the first release - applied to pre-existing databases.
MIGRATIONS = {
    "Projects_Master": {
        "SalesPerson": "TEXT",
        "ContactPhone": "TEXT",
        "Contractor": "TEXT",
        "Region": "TEXT",
        "PresalesEngineer": "TEXT",
        "ProjectManager": "TEXT",
        "RevisionNo": "INTEGER DEFAULT 0",
        "BaseName": "TEXT",
        "OfferTerms": "TEXT",
        "ProjectSheetInfo": "TEXT",
        "Approved": "INTEGER DEFAULT 0",
        "ApprovedAt": "TEXT",
        "OptionLabel": "TEXT",
        "Archived": "INTEGER DEFAULT 0",
        "ArchivedBy": "INTEGER",
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
    },
    "Items_Catalog": {
        "ShippingPercent": "REAL",
        "Currency": "TEXT DEFAULT 'USD'",
        "PriceUpdatedAt": "TEXT DEFAULT '2025-01-01'",
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


def init_db(db_path: str = DB_PATH) -> sqlite3.Connection:
    conn = connect(db_path)
    conn.executescript(SCHEMA)
    _migrate(conn)
    _backfill_tracking_quantities(conn)
    _backfill_catalog_price_dates(conn)
    conn.commit()
    return conn


if __name__ == "__main__":
    c = init_db()
    print(f"Initialized schema at: {DB_PATH}")
    for (name,) in c.execute(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    ):
        print("  table:", name)
    c.close()
