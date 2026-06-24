"""PostgreSQL backend and SQLite-compatible connection facade for ProQuote."""
from __future__ import annotations

import os
import re
import atexit
import threading
from collections.abc import Iterator


_write_epoch: int = 0
_write_epoch_lock = threading.Lock()


def _bump_write_epoch() -> None:
    global _write_epoch
    with _write_epoch_lock:
        _write_epoch += 1


def write_epoch() -> int:
    return _write_epoch


TABLES_IN_LOAD_ORDER = (
    "Settings",
    "App_Assets",
    "Roles",
    "RolePerms",
    "Users",
    "Projects_Master",
    "Project_Sheets",
    "Items_Catalog",
    "Project_BoQ_Lines",
    "Finance_Payments",
    "Finance_Purchases",
    "Audit_Log",
)
TABLES_IN_DELETE_ORDER = tuple(reversed(TABLES_IN_LOAD_ORDER))

IDENTITY_COLUMNS = {
    "Projects_Master": "ProjectID",
    "Project_Sheets": "SheetID",
    "Items_Catalog": "ItemID",
    "Project_BoQ_Lines": "LineID",
    "Users": "UserID",
    "Finance_Payments": "PayID",
    "Finance_Purchases": "PurID",
    "Audit_Log": "AuditID",
}

_CANONICAL_NAMES = """
ProjectID ProjectName ClientName ContactName ContactPhone Contractor Region SalesPerson
PresalesEngineer ProjectManager OfferNo CreationDate UpdatedDate DiscountAmount
CommissionAmount CommissionPercent CommissionMode ConversionFactor SourceFile IngestedAt
RevisionNo BaseName OfferTerms ProjectSheetInfo Approved ApprovedAt OptionLabel Archived
ArchivedBy SheetID SheetName SystemSuffix Factor1 Factor2 Factor3 SubtotalSAR GrandTotalSAR
ItemID Description Brand Model ListPriceUSD ExUnitCostUSD Currency ShippingPercent UnitCostUSD
DefaultUPriceUSD DefaultUPriceSAR PriceUpdatedAt TimesQuoted LastSeenFile LastSeenAt LineID
RowOrder Area System Qty FinalUnitCostUSD TotalCostUSD FinalUPriceUSD TPriceUSD FinalUPriceSAR
TPriceSAR MarginExtra LineType Paid Received ReceivedQty Delivered DeliveredQty PONumber
DeliveryNote PaidAt ReceivedAt DeliveredAt ReceivedRegion IncludedInItems InclusionMode InclusionMarkup Discontinued key value UserID Username DisplayName PasswordHash
Role Active CreatedAt Permission PayID AmountSAR InvoiceNo PurID PORef AuditID EventAt Action
EntityType EntityID Summary OldValues NewValues Value Collected PO_Spend Count m n
OfferCount CurrentValue Margin
AssetKey FileName MimeType Content UpdatedAt
Deleted
""".split()
_CANONICAL = {name.lower(): name for name in _CANONICAL_NAMES}


class CompatRow:
    """Row supporting both sqlite.Row-style numeric and named access."""

    def __init__(self, names: list[str], values):
        self._names = names
        self._values = tuple(values)
        self._index = {name.lower(): i for i, name in enumerate(names)}

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        return self._values[self._index[str(key).lower()]]

    def __iter__(self) -> Iterator:
        return iter(self._values)

    def __len__(self) -> int:
        return len(self._values)

    def keys(self):
        return list(self._names)


def _compat_names(description) -> list[str]:
    names = []
    for column in description or ():
        raw = getattr(column, "name", column[0])
        names.append(_CANONICAL.get(str(raw).lower(), str(raw)))
    return names


def _replace_qmarks(sql: str) -> str:
    """Convert qmark placeholders outside SQL string/identifier literals."""
    out, quote = [], None
    i = 0
    while i < len(sql):
        char = sql[i]
        if quote:
            out.append(char)
            if char == quote:
                if i + 1 < len(sql) and sql[i + 1] == quote:
                    out.append(sql[i + 1])
                    i += 1
                else:
                    quote = None
        elif char in {"'", '"'}:
            quote = char
            out.append(char)
        elif char == "?":
            out.append("%s")
        else:
            out.append(char)
        i += 1
    return "".join(out)


def postgres_sql(sql: str) -> str:
    sql = _replace_qmarks(sql)
    sql = re.sub(r"\bIFNULL\s*\(", "COALESCE(", sql, flags=re.I)
    sql = re.sub(r"\bNOT\s+LIKE\b", "NOT ILIKE", sql, flags=re.I)
    sql = re.sub(r"(?<!NOT )\bLIKE\b", "ILIKE", sql, flags=re.I)
    return sql


class CompatCursor:
    def __init__(self, cursor):
        self._cursor = cursor
        self._names = _compat_names(cursor.description)

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def _row(self, values):
        return None if values is None else CompatRow(self._names, values)

    def fetchone(self):
        return self._row(self._cursor.fetchone())

    def fetchall(self):
        return [self._row(row) for row in self._cursor.fetchall()]

    def __iter__(self):
        for row in self._cursor:
            yield self._row(row)


_POOL = None
_POOL_URL = None
_SCHEMA_READY = set()
_SCHEMA_LOCK = threading.Lock()


def _configure_connection(conn) -> None:
    """Keep read-only calls out of transactions so pool return is network-free."""
    conn.autocommit = True


def close_pool() -> None:
    global _POOL, _POOL_URL
    if _POOL is not None:
        _POOL.close()
    _POOL = None
    _POOL_URL = None


atexit.register(close_pool)


def _pool(database_url: str):
    global _POOL, _POOL_URL
    if _POOL is not None and _POOL_URL == database_url:
        return _POOL
    if _POOL is not None:
        _POOL.close()
    try:
        from psycopg_pool import ConnectionPool
    except ImportError as exc:
        raise RuntimeError(
            "PostgreSQL is configured but psycopg is not installed. "
            "Run: pip install 'psycopg[binary,pool]'"
        ) from exc
    _POOL = ConnectionPool(
        conninfo=database_url,
        min_size=max(1, int(os.environ.get("PROQUOTE_DB_POOL_MIN_SIZE", "2"))),
        max_size=max(2, int(os.environ.get("PROQUOTE_DB_POOL_SIZE", "10"))),
        timeout=30,
        kwargs={
            "prepare_threshold": None,
            "keepalives": 1,
            "keepalives_idle": 60,
            "keepalives_interval": 10,
            "keepalives_count": 5,
        },
        configure=_configure_connection,
        open=True,
    )
    _POOL_URL = database_url
    return _POOL


class PostgresConnection:
    def __init__(self, database_url: str, actor: dict | None = None):
        self._database_url = database_url
        self._actor = actor or {}
        self._closed = False
        self._dirty = False
        self._actor_ready = False
        self._checkout()

    def _checkout(self) -> None:
        self._context = _pool(self._database_url).connection(timeout=30)
        self._raw = self._context.__enter__()

    @staticmethod
    def _is_write(sql: str) -> bool:
        clean = re.sub(r"^\s*(?:--[^\n]*\n|/\*.*?\*/\s*)*", "", sql, flags=re.S)
        match = re.match(r"([A-Za-z]+)", clean)
        keyword = match.group(1).upper() if match else ""
        if keyword == "WITH":
            return bool(re.search(r"\b(?:INSERT|UPDATE|DELETE|MERGE)\b", clean, flags=re.I))
        return keyword not in {"SELECT", "SHOW", "EXPLAIN", "VALUES"}

    def _begin_write(self) -> None:
        if not self._dirty:
            self._raw.autocommit = False
            self._dirty = True
        if self._actor_ready:
            return
        actor = self._actor
        self._raw.execute(
            "SELECT set_config('proquote.user_id', %s, true), "
            "set_config('proquote.username', %s, true), "
            "set_config('proquote.display_name', %s, true)",
            (
                "" if actor.get("user_id") is None else str(actor.get("user_id")),
                str(actor.get("username") or "system"),
                str(actor.get("display_name") or actor.get("username") or "System"),
            ),
        )
        self._actor_ready = True

    def _replace_failed_read_connection(self, exc) -> None:
        """Discard one stale idle connection; the next attempt gets a fresh one."""
        try:
            self._context.__exit__(type(exc), exc, exc.__traceback__)
        finally:
            self._checkout()

    def execute(self, sql: str, params=None):
        converted = postgres_sql(sql)
        values = tuple(params or ())
        is_write = self._is_write(converted)
        if is_write:
            self._begin_write()
        for attempt in range(2):
            try:
                cur = (self._raw.execute(converted, values)
                       if values else self._raw.execute(converted))
                return CompatCursor(cur)
            except Exception as exc:
                if is_write or attempt or self._dirty:
                    raise
                try:
                    import psycopg.errors as pg_errors
                    retryable = isinstance(
                        exc, (pg_errors.AdminShutdown, pg_errors.OperationalError, OSError)
                    )
                except ImportError:
                    retryable = isinstance(exc, OSError)
                if not retryable:
                    raise
                self._replace_failed_read_connection(exc)

    def query_many(self, statements):
        """Run independent reads in one PostgreSQL pipeline/network round trip."""
        prepared = []
        for sql, params in statements:
            converted = postgres_sql(sql)
            if self._is_write(converted):
                raise ValueError("query_many accepts read-only statements")
            prepared.append((converted, tuple(params or ())))
        cursors = []
        with self._raw.pipeline():
            for converted, values in prepared:
                cursors.append(
                    self._raw.execute(converted, values)
                    if values else self._raw.execute(converted)
                )
        return [CompatCursor(cur) for cur in cursors]

    def begin_transaction(self) -> None:
        """Start an explicit transaction for pooled-connection maintenance work."""
        self._begin_write()

    def begin_bulk_write(self) -> None:
        """Suppress row-by-row child audit/touch work for one offer transaction."""
        self._begin_write()
        self._raw.execute("SELECT set_config('proquote.bulk_write', '1', true)")

    def executemany(self, sql: str, params_seq):
        self._begin_write()
        cur = self._raw.cursor()
        cur.executemany(postgres_sql(sql), params_seq)
        return CompatCursor(cur)

    def executescript(self, sql: str):
        self._begin_write()
        cur = self._raw.execute(sql, prepare=False)
        return CompatCursor(cur)

    def commit(self):
        if not self._dirty:
            return
        self._raw.commit()
        self._raw.autocommit = True
        self._dirty = False
        self._actor_ready = False
        _bump_write_epoch()

    def rollback(self):
        if not self._dirty:
            return
        self._raw.rollback()
        self._raw.autocommit = True
        self._dirty = False
        self._actor_ready = False

    def close(self):
        if self._closed:
            return
        if self._dirty:
            self.commit()
        self._context.__exit__(None, None, None)
        self._closed = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._closed:
            return False
        if self._dirty:
            self.rollback() if exc_type else self.commit()
        self._context.__exit__(None, None, None)
        self._closed = True
        return False


POSTGRES_SCHEMA = r"""
CREATE TABLE IF NOT EXISTS projects_master (
    projectid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    projectname TEXT, clientname TEXT, contactname TEXT, contactphone TEXT,
    contractor TEXT, region TEXT, salesperson TEXT, presalesengineer TEXT,
    projectmanager TEXT, offerno TEXT, creationdate TEXT, updateddate TEXT,
    discountamount DOUBLE PRECISION DEFAULT 0,
    commissionamount DOUBLE PRECISION DEFAULT 0,
    commissionpercent DOUBLE PRECISION DEFAULT 0,
    commissionmode TEXT DEFAULT 'Deduct from profit',
    conversionfactor DOUBLE PRECISION, sourcefile TEXT UNIQUE, ingestedat TEXT,
    revisionno INTEGER DEFAULT 0, basename TEXT, offerterms TEXT,
    projectsheetinfo TEXT, approved INTEGER DEFAULT 0, approvedat TEXT,
    optionlabel TEXT, archived INTEGER DEFAULT 0, archivedby BIGINT
);

CREATE TABLE IF NOT EXISTS project_sheets (
    sheetid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    projectid BIGINT NOT NULL REFERENCES projects_master(projectid) ON DELETE CASCADE,
    sheetname TEXT, systemsuffix TEXT, discountamount DOUBLE PRECISION DEFAULT 0,
    factor1 DOUBLE PRECISION, factor2 DOUBLE PRECISION, factor3 DOUBLE PRECISION,
    subtotalsar DOUBLE PRECISION, grandtotalsar DOUBLE PRECISION
);

CREATE TABLE IF NOT EXISTS items_catalog (
    itemid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    description TEXT, brand TEXT, model TEXT,
    listpriceusd DOUBLE PRECISION, exunitcostusd DOUBLE PRECISION,
    currency TEXT DEFAULT 'USD', shippingpercent DOUBLE PRECISION DEFAULT 30,
    unitcostusd DOUBLE PRECISION, defaultupriceusd DOUBLE PRECISION,
    defaultupricesar DOUBLE PRECISION, priceupdatedat TEXT DEFAULT '2025-01-01',
    timesquoted INTEGER DEFAULT 0, lastseenfile TEXT, lastseenat TEXT,
    discontinued INTEGER DEFAULT 0,
    UNIQUE (brand, model, description)
);

CREATE TABLE IF NOT EXISTS project_boq_lines (
    lineid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    projectid BIGINT NOT NULL REFERENCES projects_master(projectid) ON DELETE CASCADE,
    sheetid BIGINT REFERENCES project_sheets(sheetid) ON DELETE CASCADE,
    itemid BIGINT REFERENCES items_catalog(itemid), roworder INTEGER,
    area TEXT, system TEXT, description TEXT, brand TEXT, model TEXT,
    qty DOUBLE PRECISION, listpriceusd DOUBLE PRECISION,
    exunitcostusd DOUBLE PRECISION, currency TEXT DEFAULT 'USD',
    shippingpercent DOUBLE PRECISION DEFAULT 30,
    finalunitcostusd DOUBLE PRECISION, totalcostusd DOUBLE PRECISION,
    finalupriceusd DOUBLE PRECISION, tpriceusd DOUBLE PRECISION,
    finalupricesar DOUBLE PRECISION, tpricesar DOUBLE PRECISION,
    marginextra DOUBLE PRECISION, linetype TEXT DEFAULT 'item',
    paid INTEGER DEFAULT 0, received INTEGER DEFAULT 0,
    receivedqty DOUBLE PRECISION DEFAULT 0, delivered INTEGER DEFAULT 0,
    deliveredqty DOUBLE PRECISION DEFAULT 0, ponumber TEXT, deliverynote TEXT,
    paidat TEXT, receivedat TEXT, deliveredat TEXT
);

CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS app_assets (
    assetkey TEXT PRIMARY KEY, filename TEXT, mimetype TEXT,
    content BYTEA NOT NULL, updatedat TEXT, deleted INTEGER DEFAULT 0
);
ALTER TABLE app_assets ADD COLUMN IF NOT EXISTS deleted INTEGER DEFAULT 0;
CREATE TABLE IF NOT EXISTS users (
    userid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    username TEXT UNIQUE NOT NULL, displayname TEXT, passwordhash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'viewer', active INTEGER DEFAULT 1, createdat TEXT
);
CREATE TABLE IF NOT EXISTS roles (role TEXT PRIMARY KEY);
CREATE TABLE IF NOT EXISTS roleperms (
    role TEXT NOT NULL, permission TEXT NOT NULL, PRIMARY KEY (role, permission)
);
CREATE TABLE IF NOT EXISTS finance_payments (
    payid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    projectid BIGINT NOT NULL REFERENCES projects_master(projectid) ON DELETE CASCADE,
    roworder INTEGER, description TEXT, amountsar DOUBLE PRECISION DEFAULT 0,
    invoiceno TEXT
);
CREATE TABLE IF NOT EXISTS finance_purchases (
    purid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    projectid BIGINT NOT NULL REFERENCES projects_master(projectid) ON DELETE CASCADE,
    roworder INTEGER, description TEXT, amountsar DOUBLE PRECISION DEFAULT 0,
    poref TEXT
);
CREATE TABLE IF NOT EXISTS audit_log (
    auditid BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    eventat TEXT NOT NULL, userid BIGINT, username TEXT NOT NULL,
    displayname TEXT, action TEXT NOT NULL, entitytype TEXT NOT NULL,
    entityid TEXT, summary TEXT, oldvalues TEXT, newvalues TEXT
);

CREATE INDEX IF NOT EXISTS idx_catalog_model ON items_catalog(model);
CREATE INDEX IF NOT EXISTS idx_catalog_desc ON items_catalog(description);
CREATE INDEX IF NOT EXISTS idx_projects_offer_no ON projects_master(offerno);
CREATE INDEX IF NOT EXISTS idx_projects_archived ON projects_master(archived);
CREATE INDEX IF NOT EXISTS idx_sheets_project ON project_sheets(projectid);
CREATE INDEX IF NOT EXISTS idx_lines_project ON project_boq_lines(projectid);
CREATE INDEX IF NOT EXISTS idx_lines_item ON project_boq_lines(itemid);
CREATE INDEX IF NOT EXISTS idx_lines_type_project ON project_boq_lines(linetype, projectid);
CREATE INDEX IF NOT EXISTS idx_fin_pay_project ON finance_payments(projectid);
CREATE INDEX IF NOT EXISTS idx_fin_pur_project ON finance_purchases(projectid);
CREATE INDEX IF NOT EXISTS idx_audit_event_at ON audit_log(eventat DESC);
CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_log(username, eventat DESC);
CREATE INDEX IF NOT EXISTS idx_audit_entity ON audit_log(entitytype, entityid, eventat DESC);

CREATE OR REPLACE FUNCTION proquote_touch_project() RETURNS trigger AS $$
DECLARE target_id BIGINT;
BEGIN
    target_id := CASE WHEN TG_OP = 'DELETE' THEN OLD.projectid ELSE NEW.projectid END;
    UPDATE projects_master
       SET updateddate=to_char(clock_timestamp() AT TIME ZONE 'Asia/Riyadh', 'YYYY-MM-DD')
     WHERE projectid=target_id;
    RETURN CASE WHEN TG_OP = 'DELETE' THEN OLD ELSE NEW END;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION proquote_master_updated_date() RETURNS trigger AS $$
BEGIN
    IF NEW.updateddate IS NOT DISTINCT FROM OLD.updateddate THEN
        NEW.updateddate := to_char(clock_timestamp() AT TIME ZONE 'Asia/Riyadh', 'YYYY-MM-DD');
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION proquote_audit_row() RETURNS trigger AS $$
DECLARE
    old_data JSONB;
    new_data JSONB;
    row_data JSONB;
    entity_id TEXT;
    entity_type TEXT;
    actor_id BIGINT;
BEGIN
    old_data := CASE WHEN TG_OP = 'INSERT' THEN NULL ELSE to_jsonb(OLD) END;
    new_data := CASE WHEN TG_OP = 'DELETE' THEN NULL ELSE to_jsonb(NEW) END;
    IF TG_TABLE_NAME = 'users' THEN
        IF old_data IS NOT NULL THEN old_data := jsonb_set(old_data, '{passwordhash}', '"[REDACTED]"'); END IF;
        IF new_data IS NOT NULL THEN new_data := jsonb_set(new_data, '{passwordhash}', '"[REDACTED]"'); END IF;
    END IF;
    IF TG_OP = 'UPDATE' AND TG_TABLE_NAME = 'projects_master'
       AND (old_data - 'updateddate') = (new_data - 'updateddate') THEN
        RETURN NEW;
    END IF;
    row_data := COALESCE(new_data, old_data);
    entity_type := CASE TG_TABLE_NAME
        WHEN 'projects_master' THEN 'Projects_Master'
        WHEN 'project_sheets' THEN 'Project_Sheets'
        WHEN 'project_boq_lines' THEN 'Project_BoQ_Lines'
        WHEN 'items_catalog' THEN 'Items_Catalog'
        WHEN 'finance_payments' THEN 'Finance_Payments'
        WHEN 'finance_purchases' THEN 'Finance_Purchases'
        WHEN 'settings' THEN 'Settings'
        WHEN 'users' THEN 'Users'
        WHEN 'roles' THEN 'Roles'
        WHEN 'roleperms' THEN 'RolePerms'
        ELSE TG_TABLE_NAME END;
    entity_id := CASE TG_TABLE_NAME
        WHEN 'projects_master' THEN row_data->>'projectid'
        WHEN 'project_sheets' THEN row_data->>'sheetid'
        WHEN 'project_boq_lines' THEN row_data->>'lineid'
        WHEN 'items_catalog' THEN row_data->>'itemid'
        WHEN 'finance_payments' THEN row_data->>'payid'
        WHEN 'finance_purchases' THEN row_data->>'purid'
        WHEN 'users' THEN row_data->>'userid'
        WHEN 'settings' THEN row_data->>'key'
        WHEN 'roles' THEN row_data->>'role'
        WHEN 'roleperms' THEN COALESCE(row_data->>'role','') || '; ' || COALESCE(row_data->>'permission','')
        ELSE '' END;
    actor_id := NULLIF(current_setting('proquote.user_id', true), '')::BIGINT;
    INSERT INTO audit_log(
        eventat,userid,username,displayname,action,entitytype,entityid,
        summary,oldvalues,newvalues
    ) VALUES (
        to_char(clock_timestamp() AT TIME ZONE 'Asia/Riyadh', 'YYYY-MM-DD"T"HH24:MI:SS'),
        actor_id,
        COALESCE(NULLIF(current_setting('proquote.username', true), ''), 'system'),
        COALESCE(NULLIF(current_setting('proquote.display_name', true), ''), 'System'),
        TG_OP, entity_type, entity_id,
        CASE TG_OP WHEN 'INSERT' THEN 'Created ' WHEN 'UPDATE' THEN 'Updated ' ELSE 'Deleted ' END || entity_type,
        CASE WHEN old_data IS NULL THEN NULL ELSE old_data::TEXT END,
        CASE WHEN new_data IS NULL THEN NULL ELSE new_data::TEXT END
    );
    RETURN CASE WHEN TG_OP = 'DELETE' THEN OLD ELSE NEW END;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_projects_master_updated_date ON projects_master;
CREATE TRIGGER trg_projects_master_updated_date
BEFORE UPDATE ON projects_master FOR EACH ROW EXECUTE FUNCTION proquote_master_updated_date();

DROP TRIGGER IF EXISTS trg_project_sheets_touch ON project_sheets;
CREATE TRIGGER trg_project_sheets_touch
AFTER INSERT OR UPDATE OR DELETE ON project_sheets FOR EACH ROW
WHEN (COALESCE(current_setting('proquote.bulk_write', true),'') <> '1')
EXECUTE FUNCTION proquote_touch_project();
DROP TRIGGER IF EXISTS trg_project_lines_touch ON project_boq_lines;
CREATE TRIGGER trg_project_lines_touch
AFTER INSERT OR UPDATE OR DELETE ON project_boq_lines FOR EACH ROW
WHEN (COALESCE(current_setting('proquote.bulk_write', true),'') <> '1')
EXECUTE FUNCTION proquote_touch_project();

DO $$
DECLARE table_name TEXT;
BEGIN
    FOREACH table_name IN ARRAY ARRAY[
        'projects_master','project_sheets','project_boq_lines','items_catalog',
        'finance_payments','finance_purchases','settings','users','roles','roleperms'
    ] LOOP
        EXECUTE format('DROP TRIGGER IF EXISTS audit_%s_row ON %I', table_name, table_name);
        EXECUTE format(
            'CREATE TRIGGER audit_%s_row AFTER INSERT OR UPDATE OR DELETE ON %I '
            'FOR EACH ROW WHEN (COALESCE(current_setting(''proquote.bulk_write'', true),'''') <> ''1'') '
            'EXECUTE FUNCTION proquote_audit_row()', table_name, table_name
        );
    END LOOP;
END $$;
"""


POSTGRES_PERFORMANCE_SCHEMA = r"""
DROP TRIGGER IF EXISTS trg_project_sheets_touch ON project_sheets;
CREATE TRIGGER trg_project_sheets_touch
AFTER INSERT OR UPDATE OR DELETE ON project_sheets FOR EACH ROW
WHEN (COALESCE(current_setting('proquote.bulk_write', true),'') <> '1')
EXECUTE FUNCTION proquote_touch_project();

DROP TRIGGER IF EXISTS trg_project_lines_touch ON project_boq_lines;
CREATE TRIGGER trg_project_lines_touch
AFTER INSERT OR UPDATE OR DELETE ON project_boq_lines FOR EACH ROW
WHEN (COALESCE(current_setting('proquote.bulk_write', true),'') <> '1')
EXECUTE FUNCTION proquote_touch_project();

DO $$
DECLARE table_name TEXT;
BEGIN
    FOREACH table_name IN ARRAY ARRAY[
        'projects_master','project_sheets','project_boq_lines','items_catalog',
        'finance_payments','finance_purchases','settings','users','roles','roleperms'
    ] LOOP
        EXECUTE format('DROP TRIGGER IF EXISTS audit_%s_row ON %I', table_name, table_name);
        EXECUTE format(
            'CREATE TRIGGER audit_%s_row AFTER INSERT OR UPDATE OR DELETE ON %I '
            'FOR EACH ROW WHEN (COALESCE(current_setting(''proquote.bulk_write'', true),'''') <> ''1'') '
            'EXECUTE FUNCTION proquote_audit_row()', table_name, table_name
        );
    END LOOP;
END $$;
"""

def connect(database_url: str, actor: dict | None = None) -> PostgresConnection:
    return PostgresConnection(database_url, actor)


# Additive column migrations for existing databases (never drop or rename).
_COLUMN_MIGRATIONS = [
    "ALTER TABLE items_catalog ADD COLUMN IF NOT EXISTS discontinued INTEGER DEFAULT 0",
    "ALTER TABLE project_boq_lines ADD COLUMN IF NOT EXISTS receivedregion TEXT",
    "ALTER TABLE project_boq_lines ADD COLUMN IF NOT EXISTS includedinItems INTEGER DEFAULT 0",
    "ALTER TABLE projects_master ADD COLUMN IF NOT EXISTS inclusionmode TEXT",
    "ALTER TABLE projects_master ADD COLUMN IF NOT EXISTS inclusionmarkup DOUBLE PRECISION",
]


def _apply_performance_migration(conn) -> None:
    version_key = "schema::postgres_performance"
    expected = "2"
    current = conn.execute(
        "SELECT value FROM Settings WHERE key=?", (version_key,)
    ).fetchone()
    if current and current["value"] == expected:
        return

    conn.begin_transaction()
    try:
        conn.execute("SELECT pg_advisory_xact_lock(hashtext('proquote_schema_migration'))")
        current = conn.execute(
            "SELECT value FROM Settings WHERE key=?", (version_key,)
        ).fetchone()
        if current and current["value"] == expected:
            conn.commit()
            return
        conn.executescript(POSTGRES_PERFORMANCE_SCHEMA)
        conn.execute(
            "INSERT INTO Settings(key,value) VALUES(?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (version_key, expected),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def _apply_column_migrations(conn) -> None:
    conn.execute("SET lock_timeout = '5s'")
    for stmt in _COLUMN_MIGRATIONS:
        try:
            conn.execute(stmt)
            conn.commit()
        except Exception:
            conn.rollback()
    conn.execute("SET lock_timeout = '0'")
    conn.commit()
    _apply_performance_migration(conn)


def init_db(database_url: str, actor: dict | None = None) -> PostgresConnection:
    conn = connect(database_url, actor)
    if database_url in _SCHEMA_READY:
        return conn
    try:
        with _SCHEMA_LOCK:
            if database_url not in _SCHEMA_READY:
                schema_exists = conn.execute(
                    "SELECT 1 FROM information_schema.tables "
                    "WHERE table_schema='public' AND table_name='projects_master'"
                ).fetchone()
                if not schema_exists:
                    conn.executescript(POSTGRES_SCHEMA)
                    conn.commit()
                _apply_column_migrations(conn)
                _SCHEMA_READY.add(database_url)
        return conn
    except Exception:
        conn.rollback()
        conn.close()
        raise


def reset_identity_sequences(conn: PostgresConnection) -> None:
    for table, column in IDENTITY_COLUMNS.items():
        conn.execute(
            "SELECT setval(pg_get_serial_sequence(?, ?), "
            "COALESCE((SELECT MAX(" + column + ") FROM " + table + "), 1), "
            "EXISTS(SELECT 1 FROM " + table + "))",
            (table.lower(), column.lower()),
        )
