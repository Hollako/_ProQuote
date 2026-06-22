"""Read helpers for ProQuote's immutable database audit trail."""
from __future__ import annotations

import json
import pandas as pd

import db


ENTITY_LABELS = {
    "Projects_Master": "Offers / projects",
    "Project_Sheets": "Offer sheets",
    "Project_BoQ_Lines": "BoQ lines / tracking",
    "Items_Catalog": "Products catalogue",
    "Finance_Payments": "Finance payments",
    "Finance_Purchases": "Finance purchases",
    "Settings": "Settings",
    "Users": "Users",
    "Roles": "Roles",
    "RolePerms": "Role permissions",
    "Database": "Database / backups",
}
ACTION_LABELS = {
    "INSERT": "Created", "UPDATE": "Updated", "DELETE": "Deleted", "RESTORE": "Restored"
}


def record_event(action: str, entity_type: str, entity_id="", summary="",
                 old_values=None, new_values=None) -> int:
    """Record a meaningful non-row action such as restoring a database backup."""
    actor = db.get_audit_actor()
    with db.connect() as conn:
        cur = conn.execute(
            """INSERT INTO Audit_Log
                   (EventAt,UserID,Username,DisplayName,Action,EntityType,EntityID,
                    Summary,OldValues,NewValues)
               VALUES (strftime('%Y-%m-%dT%H:%M:%S','now','localtime'),?,?,?,?,?,?,?,?,?)""",
            (
                actor.get("user_id"), actor.get("username"), actor.get("display_name"),
                str(action or "EVENT").upper(), entity_type, str(entity_id or ""), summary,
                json.dumps(old_values, default=str) if old_values is not None else None,
                json.dumps(new_values, default=str) if new_values is not None else None,
            ),
        )
        conn.commit()
        return int(cur.lastrowid)


def filter_options() -> dict:
    with db.connect() as conn:
        users = [r[0] for r in conn.execute(
            "SELECT DISTINCT Username FROM Audit_Log WHERE TRIM(Username)<>'' ORDER BY Username"
        )]
        entities = [r[0] for r in conn.execute(
            "SELECT DISTINCT EntityType FROM Audit_Log ORDER BY EntityType"
        )]
        actions = [r[0] for r in conn.execute(
            "SELECT DISTINCT Action FROM Audit_Log ORDER BY Action"
        )]
    return {"users": users, "entities": entities, "actions": actions}


def query_events(*, username="", action="", entity_type="", search="",
                 date_from=None, date_to=None, limit=250) -> tuple[pd.DataFrame, int]:
    where, params = [], []
    if username:
        where.append("Username=?")
        params.append(username)
    if action:
        where.append("Action=?")
        params.append(action)
    if entity_type:
        where.append("EntityType=?")
        params.append(entity_type)
    if date_from:
        where.append("substr(EventAt,1,10)>=?")
        params.append(str(date_from))
    if date_to:
        where.append("substr(EventAt,1,10)<=?")
        params.append(str(date_to))
    if search:
        where.append(
            "(Username LIKE ? OR DisplayName LIKE ? OR EntityID LIKE ? OR Summary LIKE ? "
            "OR OldValues LIKE ? OR NewValues LIKE ?)"
        )
        term = f"%{search.strip()}%"
        params.extend([term] * 6)
    clause = " WHERE " + " AND ".join(where) if where else ""
    limit = max(1, min(int(limit or 250), 2000))
    with db.connect() as conn:
        total = conn.execute(f"SELECT COUNT(*) FROM Audit_Log{clause}", params).fetchone()[0]
        rows = conn.execute(
            f"""SELECT AuditID,EventAt,UserID,Username,DisplayName,Action,EntityType,
                       EntityID,Summary,OldValues,NewValues
                  FROM Audit_Log{clause}
                 ORDER BY AuditID DESC LIMIT ?""",
            [*params, limit],
        ).fetchall()
    return pd.DataFrame([dict(r) for r in rows]), int(total)


def parse_snapshot(raw) -> dict:
    if not raw:
        return {}
    try:
        value = json.loads(raw)
        return value if isinstance(value, dict) else {"value": value}
    except (TypeError, ValueError, json.JSONDecodeError):
        return {"value": str(raw)}


def changes_frame(event: dict) -> pd.DataFrame:
    old = parse_snapshot(event.get("OldValues"))
    new = parse_snapshot(event.get("NewValues"))
    fields = list(dict.fromkeys([*old, *new]))
    rows = []
    for field in fields:
        before, after = old.get(field), new.get(field)
        if event.get("Action") == "UPDATE" and before == after:
            continue
        rows.append({"Field": field, "Before": before, "After": after})
    return pd.DataFrame(rows, columns=["Field", "Before", "After"])


def describe_event(event: dict) -> str:
    action = event.get("Action")
    if action == "INSERT":
        return "New record"
    if action == "DELETE":
        return "Record removed"
    if action == "RESTORE":
        return event.get("Summary") or "Database restored"
    changes = changes_frame(event)
    if changes.empty:
        return "Record updated"
    fields = changes["Field"].astype(str).tolist()
    shown = ", ".join(fields[:5])
    return shown + (f" +{len(fields) - 5} more" if len(fields) > 5 else "")
