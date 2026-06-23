"""User accounts + role-based access control (login, roles, user management).

Passwords are hashed with PBKDF2-HMAC-SHA256 (salted) - no external deps.
"""
from __future__ import annotations
import hashlib
import hmac
import secrets
import datetime as dt

import db as dbmod

# Every capability the app gates on, in matrix-column order: (key, friendly label).
PERMISSIONS = [
    ("new_offer",      "New Offer"),
    ("load",           "View offers"),
    ("edit",           "Edit / revise"),
    ("approve",        "Approve"),
    ("archive",        "Archive / restore"),
    ("delete",         "Delete"),
    ("tracking",       "Item tracking"),
    ("finance",        "Finance"),
    ("reports",        "Reports & statistics"),
    ("audit",          "Audit trail"),
    ("view_costs",     "See costs"),
    ("catalogue",      "Catalogue (view)"),
    ("catalogue_edit", "Catalogue (edit)"),
    ("settings",       "Settings"),
    ("users",          "Manage users & roles"),
]
ALL_PERMS = [k for k, _ in PERMISSIONS]
PERM_LABELS = dict(PERMISSIONS)
LABEL_TO_PERM = {v: k for k, v in PERMISSIONS}

# "owner" is special: always has every permission and can never be deleted.
PROTECTED_ROLE = "owner"

# Seeded once into the DB on first run; afterwards the matrix is edited in the UI.
_DEFAULT_ROLE_ORDER = ["owner", "admin", "Top Management", "Project Manager", "Pre-Sales",
                       "sales", "procurement", "viewer"]
# Roles introduced after the first release - added to existing DBs by a one-time,
# flag-guarded top-up in ensure_roles_seeded (so deleting them later doesn't resurrect them).
_LATER_ROLES = ("Top Management",)
_LATER_PERMISSION_GRANTS = {
    "audit": ("admin", "Top Management"),
}
DEFAULT_ROLE_PERMS = {
    "owner": set(ALL_PERMS),
    "admin": set(ALL_PERMS) - {"users"},
    "Top Management": set(ALL_PERMS) - {"users", "delete"},
    "Project Manager": {"new_offer", "load", "catalogue", "edit", "approve", "archive",
                        "tracking", "finance", "reports", "view_costs"},
    "Pre-Sales": {"new_offer", "load", "catalogue", "catalogue_edit", "edit", "view_costs"},
    "sales": {"new_offer", "load", "catalogue", "edit", "view_costs"},
    "procurement": {"load", "tracking"},
    "viewer": {"load"},
}


# ---------- configurable roles (DB-backed matrix) ----------

def _seed_role(c, role) -> None:
    c.execute("INSERT INTO Roles(Role) VALUES(?) ON CONFLICT DO NOTHING", (role,))
    for p in DEFAULT_ROLE_PERMS.get(role, set()):
        c.execute("INSERT INTO RolePerms(Role,Permission) VALUES(?,?) ON CONFLICT DO NOTHING",
                  (role, p))


def ensure_roles_seeded() -> None:
    """Seed the default roles on first run; afterwards apply one-time, flag-guarded
    top-ups so roles added in later versions appear on existing databases too -
    without resurrecting a role the user has since deleted."""
    with dbmod.connect() as c:
        try:
            first_run = c.execute("SELECT COUNT(*) FROM Roles").fetchone()[0] == 0
        except Exception:
            return
        if first_run:
            for role in _DEFAULT_ROLE_ORDER:
                _seed_role(c, role)
        # One-time top-ups (also marked done on first run, so they never re-add).
        for role in _LATER_ROLES:
            flag = f"seeded_role::{role}"
            if c.execute("SELECT 1 FROM Settings WHERE key=?", (flag,)).fetchone():
                continue
            if not first_run:
                _seed_role(c, role)
            c.execute(
                "INSERT INTO Settings(key,value) VALUES(?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (flag, "1"))
        for permission, roles in _LATER_PERMISSION_GRANTS.items():
            flag = f"seeded_permission::{permission}"
            if c.execute("SELECT 1 FROM Settings WHERE key=?", (flag,)).fetchone():
                continue
            for role in roles:
                if c.execute("SELECT 1 FROM Roles WHERE Role=?", (role,)).fetchone():
                    c.execute(
                        "INSERT INTO RolePerms(Role,Permission) VALUES(?,?) "
                        "ON CONFLICT DO NOTHING",
                        (role, permission),
                    )
            c.execute(
                "INSERT INTO Settings(key,value) VALUES(?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (flag, "1"))
        c.commit()


def list_roles() -> list[str]:
    with dbmod.connect() as c:
        rows = [r["Role"] for r in c.execute("SELECT Role FROM Roles")]
    order = {r: i for i, r in enumerate(_DEFAULT_ROLE_ORDER)}
    return sorted(rows, key=lambda r: (order.get(r, 999), r.lower()))


def role_perms(role: str) -> set:
    if role == PROTECTED_ROLE:
        return set(ALL_PERMS)
    with dbmod.connect() as c:
        return {r["Permission"] for r in
                c.execute("SELECT Permission FROM RolePerms WHERE Role=?", (role,))}


def has_perm(role: str, perm: str) -> bool:
    return perm in role_perms(role)


def set_role_perms(role: str, perms) -> None:
    """Replace a role's granted permissions. The owner is left untouched (always full)."""
    if role == PROTECTED_ROLE:
        return
    keep = [p for p in perms if p in ALL_PERMS]
    with dbmod.connect() as c:
        c.execute("DELETE FROM RolePerms WHERE Role=?", (role,))
        for p in keep:
            c.execute("INSERT INTO RolePerms(Role,Permission) VALUES(?,?) "
                      "ON CONFLICT DO NOTHING", (role, p))
        c.commit()


def add_role(name: str, perms=None) -> bool:
    name = (name or "").strip()
    if not name:
        return False
    with dbmod.connect() as c:
        if c.execute("SELECT 1 FROM Roles WHERE lower(Role)=lower(?)", (name,)).fetchone():
            return False
        c.execute("INSERT INTO Roles(Role) VALUES(?)", (name,))
        for p in (perms if perms is not None else {"load"}):
            if p in ALL_PERMS:
                c.execute("INSERT INTO RolePerms(Role,Permission) VALUES(?,?) "
                          "ON CONFLICT DO NOTHING", (name, p))
        c.commit()
    return True


def delete_role(name: str) -> bool:
    if name == PROTECTED_ROLE:
        return False
    with dbmod.connect() as c:
        c.execute("DELETE FROM RolePerms WHERE Role=?", (name,))
        c.execute("DELETE FROM Roles WHERE Role=?", (name,))
        c.commit()
    return True


def role_user_count(role: str) -> int:
    with dbmod.connect() as c:
        return c.execute("SELECT COUNT(*) FROM Users WHERE Role=?", (role,)).fetchone()[0]


def users_in_role(role: str) -> list[str]:
    """Display names of active users holding `role` - used for offer people-pickers."""
    with dbmod.connect() as c:
        rows = c.execute(
            "SELECT DisplayName, Username FROM Users WHERE Role=? AND Active=1 "
            "ORDER BY DisplayName, Username", (role,)).fetchall()
    return [((r["DisplayName"] or "").strip() or r["Username"]) for r in rows]


def users_in_roles(roles) -> list[str]:
    """Display names of active users holding ANY of `roles` (case-insensitive, deduped).

    Used for the Sales Person picker, which spans several roles.
    """
    wanted = {str(r).strip().lower() for r in (roles or []) if str(r).strip()}
    if not wanted:
        return []
    with dbmod.connect() as c:
        rows = c.execute(
            "SELECT DisplayName, Username, Role FROM Users WHERE Active=1 "
            "ORDER BY DisplayName, Username").fetchall()
    seen, out = set(), []
    for r in rows:
        if str(r["Role"] or "").strip().lower() in wanted:
            name = (r["DisplayName"] or "").strip() or r["Username"]
            if name not in seen:
                seen.add(name)
                out.append(name)
    return out


# ---------- password hashing ----------

def _hash(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), 200_000).hex()
    return f"{salt}${h}"


def _verify(password: str, stored: str) -> bool:
    try:
        salt, h = stored.split("$", 1)
    except (ValueError, AttributeError):
        return False
    return hmac.compare_digest(_hash(password, salt).split("$", 1)[1], h)


# ---------- user store ----------

def user_count() -> int:
    with dbmod.connect() as c:
        try:
            return c.execute("SELECT COUNT(*) FROM Users").fetchone()[0]
        except Exception:
            return 0


def create_user(username: str, password: str, display_name: str = "", role: str = "viewer"):
    """Returns new UserID, or None if the username already exists."""
    username = (username or "").strip()
    if not username or not password:
        return None
    now = dt.datetime.now().isoformat(timespec="seconds")
    with dbmod.connect() as c:
        if c.execute("SELECT 1 FROM Users WHERE lower(Username)=lower(?)", (username,)).fetchone():
            return None
        cur = c.execute(
            "INSERT INTO Users(Username,DisplayName,PasswordHash,Role,Active,CreatedAt) "
            "VALUES(?,?,?,?,1,?) RETURNING UserID",
            (username, (display_name or "").strip() or username, _hash(password), role, now))
        user_id = cur.fetchone()["UserID"]
        c.commit()
        return user_id


def verify_login(username: str, password: str):
    with dbmod.connect() as c:
        r = c.execute("SELECT * FROM Users WHERE lower(Username)=lower(?) AND Active=1",
                      (username or "",)).fetchone()
    if r and _verify(password, r["PasswordHash"]):
        return dict(r)
    return None


def list_users() -> list[dict]:
    with dbmod.connect() as c:
        return [dict(r) for r in c.execute(
            "SELECT UserID,Username,DisplayName,Role,Active,CreatedAt FROM Users ORDER BY UserID")]


def update_user(user_id: int, display_name=None, role=None, active=None) -> None:
    sets, vals = [], []
    if display_name is not None:
        sets.append("DisplayName=?"); vals.append(display_name)
    if role is not None:
        sets.append("Role=?"); vals.append(role)
    if active is not None:
        sets.append("Active=?"); vals.append(int(bool(active)))
    if not sets:
        return
    with dbmod.connect() as c:
        c.execute(f"UPDATE Users SET {','.join(sets)} WHERE UserID=?", (*vals, user_id))
        c.commit()


def set_password(user_id: int, password: str) -> None:
    with dbmod.connect() as c:
        c.execute("UPDATE Users SET PasswordHash=? WHERE UserID=?", (_hash(password), user_id))
        c.commit()


def delete_user(user_id: int) -> None:
    with dbmod.connect() as c:
        c.execute("DELETE FROM Users WHERE UserID=?", (user_id,))
        c.commit()
