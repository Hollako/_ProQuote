"""
Data-access layer: catalogue search + project load/save for the UI and PDF.
"""
from __future__ import annotations
import re
import json
import datetime as dt
from functools import lru_cache
import pandas as pd

import db as dbmod
import calc

# Offer-terms keys persisted as JSON in Projects_Master.OfferTerms.
TERMS_KEYS = ["subject", "greeting", "system_note", "scope", "exclusions",
              "prerequisites", "delivery", "payment", "validity", "notes"]

# Project sheet keys persisted as JSON in Projects_Master.ProjectSheetInfo.
PROJECT_SHEET_KEYS = [
    "job_reference", "sheet_date", "lead_source", "commission", "shipment_by",
    "downpayment_date", "invoice_to", "delivery_instructions",
    "salesman_signature", "gm_signature",
]

CATALOG_INITIAL_PRICE_DATE = "2025-01-01"

# Safe allow-list for the Settings -> Data Tools bulk project cleanup.  Values are
# database column names, so callers never provide SQL identifiers directly.
PROJECT_CLEANUP_FIELDS = {
    "Sales Person": "SalesPerson",
    "Pre-sales Engineer": "PresalesEngineer",
    "Project Manager": "ProjectManager",
    "Client": "ClientName",
    "Contact": "ContactName",
    "Contractor": "Contractor",
    "Region": "Region",
}


def load_terms(meta: dict) -> dict:
    """Parse the OfferTerms JSON blob from a project_meta row (empty if none)."""
    raw = meta.get("OfferTerms") if meta else None
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return {}


def load_project_sheet_info(meta: dict) -> dict:
    """Parse ProjectSheetInfo JSON from a project_meta row (empty if none)."""
    raw = meta.get("ProjectSheetInfo") if meta else None
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return {}


def _str(x) -> str:
    """Coerce to a clean string, treating None / NaN as empty."""
    if x is None or (isinstance(x, float) and x != x):   # NaN != NaN
        return ""
    return str(x)


def _catalog_price_date(value=None) -> str:
    """Normalize catalogue price-date values to ISO date text."""
    if value is None or _str(value).strip() == "":
        return CATALOG_INITIAL_PRICE_DATE
    try:
        if pd.isna(value):
            return CATALOG_INITIAL_PRICE_DATE
    except (TypeError, ValueError):
        pass
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()

    text = _str(value).strip()
    for fmt in ("%Y-%m-%d", "%m-%Y", "%m/%Y", "%Y-%m", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(text[:10], fmt).date().replace(day=1).isoformat()
        except ValueError:
            continue
    return CATALOG_INITIAL_PRICE_DATE


def revision_format() -> str:
    """Template for a revision label; a run of x's = number (length = padding).
    e.g. 'Rev.x' -> Rev.1 / Rev.10 ; 'Rxx' -> R01 / R10."""
    return get_setting("revision_format") or "Rev.x"


def revision_separator() -> str:
    """What goes between the offer number and the revision token (e.g. '-' or ' ')."""
    return get_setting("revision_separator") or "-"


def revision_token(rev) -> str:
    fmt = revision_format()
    m = re.search(r"x+", fmt, flags=re.I)
    if m:
        return fmt[:m.start()] + str(int(rev)).zfill(len(m.group(0))) + fmt[m.end():]
    return f"{fmt}{int(rev)}"


@lru_cache(maxsize=1)
def _revision_strip_pattern() -> str:
    fmt = revision_format()
    m = re.search(r"x+", fmt, flags=re.I)
    if m and (fmt[:m.start()] or fmt[m.end():]):     # need a literal so we don't eat plain digits
        tok = re.escape(fmt[:m.start()]) + r"\d+" + re.escape(fmt[m.end():])
    elif not m and fmt.strip():
        tok = re.escape(fmt) + r"\d+"
    else:
        tok = r"rev\.?\s*\d+"
    return r"\s*[-/_.: ]?\s*" + tok + r"\s*$"


def base_name(name) -> str:
    """Strip a trailing revision suffix (configured format, plus legacy 'Rev.N'
    and migrated 'R01/R02' numbering) so an offer groups with its revisions."""
    s = _str(name)
    s = re.sub(_revision_strip_pattern(), "", s, flags=re.I)
    s = re.sub(r"\s*[-/_.: ]?\s*rev\.?\s*\d+\s*$", "", s, flags=re.I)   # legacy / mixed data
    s = re.sub(r"\s*[-/_.: ]\s*r\d+\s*$", "", s, flags=re.I)           # migrated 'R01/R02'
    return s.strip() or "Offer"


# ---------- Settings (key/value) ----------

DEFAULT_SETTINGS = {
    # Template variables: *TYPE* (System Offer), *YY*/*YYYY* (year). A run of x's
    # marks the auto-number slot; its length is the zero-padding.
    "offer_template": "OFR-*TYPE*-*YY*-xxx",
    "offer_number_pad": "3",
    "offer_types": "AV, LCS, ELV, CCTV, PAVA, ACC, BGM, Smart Home, Networking",
    "default_margin": "1.6",
    "revision_format": "Rev.x",     # run of x's = number (length = zero-padding)
    "revision_separator": "-",      # between offer # and revision token
    "eur_to_usd": "1.08",           # USD value of 1 EUR (SAR uses the 3.75 peg)
    "vat_percent": "15",            # VAT rate (%) applied to offers & finance
    "project_sheet_enabled": "1",   # "0" hides the Project Sheet info form + export
    # Per-company branding (banner image lives in the data dir's assets/).
    "company_name": "Company Name",
    "company_tagline": "Smart & Low-Current Systems",
    "company_contact": "Riyadh, Kingdom of Saudi Arabia",
    "company_vat_number": "",
    "company_cr_number": "",
    "company_brand_color": "#002060",
    # PDF header. The full-width banner image overrides these section settings.
    # Placeholders: {company}, {project}, {offer}, {page}, {vat_number}, {cr_number}.
    "header_left_text": "{company}",
    "header_middle_text": "",
    "header_right_text": "{offer}",
    # PDF footer. Placeholders: {company}, {project}, {offer}, {page}, {vat_number}, {cr_number}.
    "footer_left_text": "{company} - {project}",
    "footer_middle_text": "",
    "footer_right_text": "Page {page}",
    "pdf_body_template": "template1",
    # Software updates.
    "github_owner": "Hollako",
    "github_repo": "_ProQuote",
}


def get_setting(key: str, default=None):
    with _conn() as c:
        try:
            r = c.execute("SELECT value FROM Settings WHERE key=?", (key,)).fetchone()
        except Exception:
            r = None
    if r and r["value"] is not None:
        return r["value"]
    return default if default is not None else DEFAULT_SETTINGS.get(key)


def set_setting(key: str, value) -> None:
    with _conn() as c:
        c.execute("INSERT INTO Settings(key,value) VALUES(?,?) "
                  "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))
        c.commit()
    if key == "revision_format":
        _revision_strip_pattern.cache_clear()


def delete_setting(key: str) -> None:
    with _conn() as c:
        c.execute("DELETE FROM Settings WHERE key=?", (key,))
        c.commit()
    if key == "revision_format":
        _revision_strip_pattern.cache_clear()


def offer_types() -> list[str]:
    raw = get_setting("offer_types")
    return [t.strip() for t in (raw or "").split(",") if t.strip()]


# ---------- Offer reference numbers (template-based) ----------
# Template = a pattern with variables:
#   *TYPE*  -> the System Offer (AV, LCS, ...); omit it for a fixed prefix
#   *YY*    -> 2-digit year      *YYYY* -> 4-digit year
#   xxxx    -> the auto-number slot; the run length is the zero-padding
# e.g. "LG-*TYPE*-*YY*/xxxx" -> "LG-AV-26/0053".  When *TYPE* is empty, an
# adjacent separator is dropped to avoid doubles. If there's no x-run, the
# number is appended at the end (padded to offer_number_pad).

def offer_template() -> str:
    return get_setting("offer_template") or "OFR-*TYPE*-*YY*-xxx"


def _year_strs(when=None):
    d = when or dt.date.today()
    return f"{d.year % 100:02d}", str(d.year)


def _fill_template(otype, when=None) -> str:
    """Substitute *TYPE* / *YY* / *YYYY*; leaves the x-run number slot in place."""
    yy, yyyy = _year_strs(when)
    t = offer_template()
    if (otype or "").strip():
        t = re.sub(r"\*type\*", otype.strip(), t, flags=re.I)
    else:
        t = re.sub(r"\*type\*[-/_.:\s]?", "", t, flags=re.I)   # drop type + a trailing sep
    t = re.sub(r"\*yyyy\*", yyyy, t, flags=re.I)
    t = re.sub(r"\*yy\*", yy, t, flags=re.I)
    t = re.sub(r"([-/_.:])\1+", r"\1", t)                      # collapse doubled separators
    return t


def _number_slot(filled):
    """Split a filled template into (before, after, pad) around the x-run."""
    m = re.search(r"x{2,}", filled, flags=re.I)
    if m:
        return filled[:m.start()], filled[m.end():], len(m.group(0))
    pad = int(get_setting("offer_number_pad") or 3)
    return filled, "", pad


def build_offer_no(otype, number, when=None) -> str:
    before, after, pad = _number_slot(_fill_template(otype, when))
    return f"{before}{str(int(number)).zfill(pad)}{after}"


def series_key(otype, when=None) -> str:
    """Identifier for an offer-number series (per type + year); own counter."""
    before, after, _ = _number_slot(_fill_template(otype, when))
    return f"{before}\x00{after}".lower()


def latest_offer_number(otype="", when=None) -> int:
    """Highest number used within this series (ignoring -Rev suffixes)."""
    before, after, _ = _number_slot(_fill_template(otype, when))
    rx = re.compile("^" + re.escape(before) + r"(\d+)" + re.escape(after) + "$")
    where = ["OfferNo IS NOT NULL"]
    params = []
    if before:
        # Use the OfferNo index to narrow to this series before parsing revisions.
        upper = before[:-1] + chr(ord(before[-1]) + 1)
        where.append("OfferNo >= ? AND OfferNo < ?")
        params.extend([before, upper])
    elif after:
        where.append("OfferNo LIKE ?")
        params.append(f"%{after}%")
    with _conn() as c:
        rows = c.execute(
            "SELECT OfferNo FROM Projects_Master WHERE " + " AND ".join(where),
            params,
        ).fetchall()
    nums = [int(m.group(1)) for r in rows if (m := rx.match(base_name(r["OfferNo"])))]
    return max(nums) if nums else 0


def get_series_start(key: str) -> int:
    try:
        return int(get_setting(f"offer_start::{key}", "0"))
    except (TypeError, ValueError):
        return 0


def set_series_start(otype, number, when=None) -> None:
    set_setting(f"offer_start::{series_key(otype, when)}", int(number))


def clear_series_start(otype, when=None) -> None:
    delete_setting(f"offer_start::{series_key(otype, when)}")


def next_offer_number(otype="", when=None) -> int:
    """Next number for this series, honoring a forced start floor."""
    nxt = latest_offer_number(otype, when) + 1
    start = get_series_start(series_key(otype, when))
    return max(nxt, start) if start else nxt


def make_offer_no(otype="", when=None) -> str:
    """Full next offer reference for a System Offer type (template rendered)."""
    return build_offer_no(otype, next_offer_number(otype, when), when)


def next_revision(base: str) -> int:
    """Next revision number for an offer family (max existing + 1, min 1)."""
    with _conn() as c:
        r = c.execute("SELECT MAX(RevisionNo) m FROM Projects_Master WHERE BaseName=?",
                      (base,)).fetchone()
    return (r["m"] or 0) + 1


# ---------- Clean database (migrated-data fix-ups) ----------

# Offer reference at the start, e.g. "LG-AV-25-2098" / "LK-LC-24-2027" /
# "OFR-SWS-RUH-26-024": <code>-<YY>-<number>. Group 1 = full ref, group 2 = YY.
_REF_RE = re.compile(r"^\s*(\S+?-(\d{2})-[A-Za-z0-9]+)")
# Revision marker right after the ref, e.g. "... -2027 R02 ...".
_REV_RE = re.compile(r"-\d{2}-[A-Za-z0-9]+\s+R(\d+)\b", re.I)
# A full date embedded as DD.MM.YYYY (the offer ref uses dashes, so this only
# matches the trailing real date).
_DATE_RE = re.compile(r"(\d{1,2})\.(\d{1,2})\.(20\d{2})")


def parse_offer_ref(text):
    """Return (ref_code, year) from an offer reference like LK-LC-24-2027 -> ('LK-LC-24-2027', 2024)."""
    m = _REF_RE.match(_str(text))
    if not m:
        return None, None
    yy = int(m.group(2))
    return m.group(1), (2000 + yy if 20 <= yy <= 39 else None)


def parse_revision_no(text) -> int:
    """Revision number from an 'R<n>' right after the offer ref (0 = original)."""
    m = _REV_RE.search(_str(text))
    return int(m.group(1)) if m else 0


def parse_full_date(text) -> str | None:
    """Full date from a trailing DD.MM.YYYY -> 'YYYY-MM-DD' (last valid one wins)."""
    best = None
    for m in _DATE_RE.finditer(_str(text)):
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12:
            best = f"{y:04d}-{mo:02d}-{d:02d}"
    return best


def cleanup_stamp_years(apply: bool = False) -> dict:
    """Set each offer's CreationDate from its project name / offer #: the full
    DD.MM.YYYY date if present, else the year from the offer ref (keep month/day)."""
    changes, none_found, already = [], 0, 0
    with _conn() as c:
        rows = c.execute(
            "SELECT ProjectID,OfferNo,ProjectName,CreationDate FROM Projects_Master"
        ).fetchall()
    for r in rows:
        name, off = _str(r["ProjectName"]), _str(r["OfferNo"])
        new = parse_full_date(name) or parse_full_date(off)
        if not new:                                  # no full date -> try the ref year
            _ref, yr = parse_offer_ref(off)
            if not yr:
                _ref, yr = parse_offer_ref(name)
            if not yr:
                none_found += 1
                continue
            cd = _str(r["CreationDate"])
            mo = re.match(r"(\d{4})-(\d{2})-(\d{2})", cd)
            mm_dd = f"{mo.group(2)}-{mo.group(3)}" if mo else "01-01"
            new = f"{yr}-{mm_dd}"
        if _str(r["CreationDate"]) == new:
            already += 1
            continue
        label = (off or name)[:48]
        changes.append((r["ProjectID"], label, _str(r["CreationDate"]) or "(none)", new))
    if apply and changes:
        with _conn() as c:
            for pid, _l, _old, new in changes:
                c.execute("UPDATE Projects_Master SET CreationDate=? WHERE ProjectID=?", (new, pid))
            c.commit()
    return {"to_update": len(changes), "no_year": none_found, "already_ok": already,
            "sample": changes[:50]}


def cleanup_merge_revisions(apply: bool = False) -> dict:
    """Link '<ref> R01/R02' offers into one family: write the clean offer ref into
    OfferNo (so they group), set RevisionNo from the Rxx, and BaseName to the ref."""
    updates, changes = [], []
    with _conn() as c:
        rows = c.execute(
            "SELECT ProjectID,OfferNo,ProjectName,RevisionNo,BaseName FROM Projects_Master"
        ).fetchall()
    for r in rows:
        ref, _y = parse_offer_ref(r["OfferNo"])
        text = _str(r["OfferNo"])
        if not ref:
            ref, _y = parse_offer_ref(r["ProjectName"])
            text = _str(r["ProjectName"])
        if not ref:
            continue                                 # no recognizable ref -> skip
        revno = parse_revision_no(text)
        # Write the revision suffix in the company's configured format (Settings →
        # Revision format / separator), e.g. " R01" (Rxx) or "-Rev.1" (Rev.x).
        new_off = f"{ref}{revision_separator()}{revision_token(revno)}" if revno else ref
        if (int(r["RevisionNo"] or 0) == revno and _str(r["BaseName"]) == ref
                and _str(r["OfferNo"]) == new_off):
            continue
        updates.append((r["ProjectID"], new_off, revno, ref))
        changes.append((r["ProjectID"], new_off, revno, ref))
    if apply and updates:
        with _conn() as c:
            for pid, new_off, revno, ref in updates:
                c.execute("UPDATE Projects_Master SET OfferNo=?, RevisionNo=?, BaseName=? "
                          "WHERE ProjectID=?", (new_off, revno, ref, pid))
            c.commit()
    return {"to_update": len(changes), "families": len({ch[3] for ch in changes}),
            "sample": changes[:50]}


def parse_client_from_name(name) -> str | None:
    """Best-effort client from a migrated project name: drop the offer ref, revision
    and trailing date, then take the part before the scope (' - ' else last '-')."""
    s = _str(name)
    if not s:
        return None
    s = _DATE_RE.sub("", s)                                  # remove DD.MM.YYYY
    s = re.sub(r"^\s*\S+?-\d{2}-[A-Za-z0-9]+\s*", "", s)     # strip leading offer ref
    s = re.sub(r"^\s*R\d+\s+", "", s)                        # strip leading revision
    s = s.strip(" -")
    if not s:
        return None
    if " - " in s:                                           # explicit client / scope split
        s = s.split(" - ", 1)[0]
    elif "-" in s:                                           # scope is the last '-segment'
        s = s.rsplit("-", 1)[0]
    return s.strip() or None


def cleanup_parse_clients(apply: bool = False) -> dict:
    """Fill ClientName from the project name for offers where it's currently blank."""
    changes = []
    with _conn() as c:
        rows = c.execute("SELECT ProjectID,ProjectName,ClientName FROM Projects_Master").fetchall()
    for r in rows:
        if _str(r["ClientName"]):
            continue                                          # keep existing clients
        cl = parse_client_from_name(r["ProjectName"])
        if not cl:
            continue
        changes.append((r["ProjectID"], _str(r["ProjectName"])[:48], cl))
    if apply and changes:
        with _conn() as c:
            for pid, _n, cl in changes:
                c.execute("UPDATE Projects_Master SET ClientName=? WHERE ProjectID=?", (cl, pid))
            c.commit()
    return {"to_update": len(changes), "sample": changes[:50]}


def project_cleanup_values(field_label: str) -> list[dict]:
    """Distinct non-blank stored values and their project-record counts."""
    column = PROJECT_CLEANUP_FIELDS.get(field_label)
    if not column:
        raise ValueError(f"Unsupported project cleanup field: {field_label}")
    with _conn() as c:
        rows = c.execute(
            f"SELECT TRIM(COALESCE({column},'')) AS Value, COUNT(*) AS OfferCount "
            f"FROM Projects_Master WHERE TRIM(COALESCE({column},'')) <> '' "
            f"GROUP BY TRIM(COALESCE({column},'')) "
            f"ORDER BY Value COLLATE NOCASE"
        ).fetchall()
    return [dict(r) for r in rows]


def bulk_replace_project_field(field_label: str, old_values, replacement,
                               apply: bool = False) -> dict:
    """Preview or replace exact (trimmed) values in a project header field.

    Multiple imported typo variants can be merged into one canonical value.  The
    replacement itself is excluded from the source set, avoiding no-op updates when
    it was selected accidentally.
    """
    column = PROJECT_CLEANUP_FIELDS.get(field_label)
    if not column:
        raise ValueError(f"Unsupported project cleanup field: {field_label}")
    new_value = _str(replacement).strip()
    if not new_value:
        raise ValueError("Replacement value cannot be blank.")
    sources = sorted({
        _str(value).strip() for value in (old_values or [])
        if _str(value).strip() and _str(value).strip() != new_value
    }, key=str.casefold)
    if not sources:
        return {"to_update": 0, "sample": [], "sources": [], "replacement": new_value}

    marks = ",".join("?" for _ in sources)
    with _conn() as c:
        rows = c.execute(
            f"SELECT ProjectID, OfferNo, ProjectName, CreationDate, "
            f"TRIM(COALESCE({column},'')) AS CurrentValue "
            f"FROM Projects_Master "
            f"WHERE TRIM(COALESCE({column},'')) IN ({marks}) "
            f"ORDER BY CreationDate DESC, ProjectID DESC",
            sources,
        ).fetchall()
        if apply and rows:
            c.execute(
                f"UPDATE Projects_Master SET {column}=? "
                f"WHERE TRIM(COALESCE({column},'')) IN ({marks})",
                [new_value, *sources],
            )
            c.commit()

    sample = [
        (r["ProjectID"], _str(r["OfferNo"]), _str(r["ProjectName"]),
         _str(r["CreationDate"]), _str(r["CurrentValue"]), new_value)
        for r in rows[:100]
    ]
    return {
        "to_update": len(rows),
        "sample": sample,
        "sources": sources,
        "replacement": new_value,
    }


def clear_imported_data() -> dict:
    """Delete ALL projects (cascades to lines/sheets/finance) and catalogue items,
    for a clean re-import. Keeps Settings, Users, Roles and branding."""
    with _conn() as c:
        np = c.execute("SELECT COUNT(*) FROM Projects_Master").fetchone()[0]
        nc = c.execute("SELECT COUNT(*) FROM Items_Catalog").fetchone()[0]
        c.execute("DELETE FROM Projects_Master")     # FK cascade: lines/sheets/finance
        c.execute("DELETE FROM Items_Catalog")
        c.commit()
    return {"projects": np, "catalogue": nc}


def _conn():
    return dbmod.connect()


# ---------- Catalogue (Workflow 2 type-ahead) ----------

def search_catalog(term: str, limit: int = 25) -> pd.DataFrame:
    """Type-ahead over Model or Description. Empty term -> most-quoted items."""
    with _conn() as c:
        if term and term.strip():
            like = f"%{term.strip()}%"
            rows = c.execute(
                """SELECT ItemID,Brand,Model,Description,ListPriceUSD,ExUnitCostUSD,Currency,
                          ShippingPercent,UnitCostUSD,DefaultUPriceUSD,DefaultUPriceSAR,
                          PriceUpdatedAt,TimesQuoted
                   FROM Items_Catalog
                   WHERE Model LIKE ? OR Description LIKE ? OR Brand LIKE ?
                   ORDER BY TimesQuoted DESC LIMIT ?""",
                (like, like, like, limit),
            ).fetchall()
        else:
            rows = c.execute(
                """SELECT ItemID,Brand,Model,Description,ListPriceUSD,ExUnitCostUSD,Currency,
                          ShippingPercent,UnitCostUSD,DefaultUPriceUSD,DefaultUPriceSAR,
                          PriceUpdatedAt,TimesQuoted
                   FROM Items_Catalog ORDER BY TimesQuoted DESC LIMIT ?""",
                (limit,),
            ).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


CATALOG_EDITABLE = {  # display label -> DB column
    "List Price $": "ListPriceUSD",
    "Ex Unit Cost $": "ExUnitCostUSD",
    "Shipping %": "ShippingPercent",
    "Unit Cost $": "UnitCostUSD",
    "Default U.Price $": "DefaultUPriceUSD",
    "Default U.Price SAR": "DefaultUPriceSAR",
}
# DB columns that update_catalog_item is allowed to write.
_CATALOG_PRICE_WRITABLE = set(CATALOG_EDITABLE.values()) | {"Currency"}
_CATALOG_IDENTITY_WRITABLE = {"Brand", "Model", "Description"}
_CATALOG_WRITABLE = _CATALOG_PRICE_WRITABLE | _CATALOG_IDENTITY_WRITABLE


def update_catalog_item(item_id: int, fields: dict) -> bool:
    """Update identity/cost/default-price/currency fields (keyed by DB column).
    Ex cost is in the item's currency; UnitCostUSD is recomputed in USD when the
    Ex cost, shipping or currency changes. Returns False for a duplicate item."""
    if not any(c in _CATALOG_WRITABLE for c in fields):
        return False
    fields = {key: value for key, value in fields.items() if key in _CATALOG_WRITABLE}
    for field in _CATALOG_IDENTITY_WRITABLE & fields.keys():
        fields[field] = _str(fields[field]).strip()
    with _conn() as c:
        row = c.execute(
            "SELECT Brand,Model,Description,ExUnitCostUSD,ShippingPercent,UnitCostUSD,Currency "
            "FROM Items_Catalog WHERE ItemID=?",
            (int(item_id),),
        ).fetchone()
        if row:
            if _CATALOG_IDENTITY_WRITABLE & fields.keys():
                brand = fields.get("Brand", row["Brand"]) or ""
                model = fields.get("Model", row["Model"]) or ""
                description = fields.get("Description", row["Description"]) or ""
                duplicate = c.execute(
                    "SELECT ItemID FROM Items_Catalog WHERE ItemID<>? "
                    "AND IFNULL(Brand,'')=? AND IFNULL(Model,'')=? "
                    "AND IFNULL(Description,'')=?",
                    (int(item_id), brand, model, description),
                ).fetchone()
                if duplicate:
                    return False
            ex = fields.get("ExUnitCostUSD", row["ExUnitCostUSD"])
            ship = fields.get("ShippingPercent", row["ShippingPercent"])
            unit = fields.get("UnitCostUSD", row["UnitCostUSD"])
            cur = fields.get("Currency", row["Currency"]) or "USD"
            ex_usd = calc.to_usd(ex, cur)
            if (("ExUnitCostUSD" in fields or "ShippingPercent" in fields or "Currency" in fields)
                    and calc._num(ex) > 0):
                fields["UnitCostUSD"] = calc.roundup(ex_usd * (1 + calc.shipping_percent(ship) / 100), 0)
            elif "UnitCostUSD" in fields and calc._num(ex) > 0:
                fields["ShippingPercent"] = calc.infer_shipping_percent(ex_usd, unit)

        cols = [c for c in fields if c in _CATALOG_WRITABLE]
        if not cols:
            return False
        updates_price = any(col in _CATALOG_PRICE_WRITABLE for col in cols)
        sets = [f"{col}=?" for col in cols]
        values = [fields[key] for key in cols]
        if updates_price:
            sets.append("PriceUpdatedAt=?")
            values.append(dt.date.today().isoformat())
        c.execute(f"UPDATE Items_Catalog SET {', '.join(sets)} WHERE ItemID=?",
                  (*values, item_id))
        c.commit()
    return True


def add_catalog_item(brand, model, description, list_price=None, ex_cost=None,
                     shipping_percent=None, unit_cost=None, uprice_usd=None, uprice_sar=None,
                     currency="USD"):
    """Insert a new catalogue item. Returns ItemID, or None if a duplicate
    (same Brand+Model+Description) already exists. The Ex cost may be in any
    currency; UnitCostUSD is derived in USD via the configured exchange rate."""
    now = dt.datetime.now().isoformat(timespec="seconds")
    currency = currency if currency in calc.CURRENCIES else "USD"
    ship = calc.shipping_percent(shipping_percent, calc.to_usd(ex_cost, currency), unit_cost)
    if calc._num(ex_cost) > 0:
        unit_cost = calc.roundup(calc.to_usd(ex_cost, currency) * (1 + ship / 100), 0)
    with _conn() as c:
        dup = c.execute(
            "SELECT ItemID FROM Items_Catalog WHERE IFNULL(Brand,'')=? AND "
            "IFNULL(Model,'')=? AND IFNULL(Description,'')=?",
            (brand or "", model or "", description or "")).fetchone()
        if dup:
            return None
        cur = c.execute(
            """INSERT INTO Items_Catalog
                 (Description,Brand,Model,ListPriceUSD,ExUnitCostUSD,Currency,ShippingPercent,
                  UnitCostUSD,DefaultUPriceUSD,DefaultUPriceSAR,PriceUpdatedAt,TimesQuoted,
                  LastSeenFile,LastSeenAt)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,0,?,?)""",
            (description, brand, model, list_price, ex_cost, currency,
             ship, unit_cost,
             uprice_usd, uprice_sar, dt.date.today().isoformat(), "app://catalog-add", now))
        c.commit()
        return cur.lastrowid


def delete_catalog_items(ids) -> int:
    """Delete catalogue items; unlinks them from any offer lines first (lines keep
    their own data, just lose the catalogue reference)."""
    ids = [int(i) for i in ids]
    if not ids:
        return 0
    qs = ",".join("?" * len(ids))
    with _conn() as c:
        c.execute(f"UPDATE Project_BoQ_Lines SET ItemID=NULL WHERE ItemID IN ({qs})", ids)
        c.execute(f"DELETE FROM Items_Catalog WHERE ItemID IN ({qs})", ids)
        c.commit()
    return len(ids)


# Columns carried in a catalogue backup / restore file (ItemID is regenerated).
CATALOG_BACKUP_COLS = ["Brand", "Model", "Description", "ListPriceUSD", "ExUnitCostUSD",
                       "Currency", "ShippingPercent", "UnitCostUSD", "DefaultUPriceUSD",
                       "DefaultUPriceSAR", "PriceUpdatedAt", "TimesQuoted"]


def catalog_all() -> pd.DataFrame:
    """Every catalogue item (for backup / dedupe), ordered Brand, Model, Description."""
    with _conn() as c:
        rows = c.execute(
            "SELECT ItemID, " + ",".join(CATALOG_BACKUP_COLS) + " FROM Items_Catalog "
            "ORDER BY Brand, Model, Description").fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def replace_catalog(df: pd.DataFrame) -> int:
    """Restore: replace the WHOLE catalogue with df's rows. Offer lines are unlinked
    from the old IDs first (they keep their own stored data). Refuses to clear the
    catalogue if the file has no usable rows. Returns the item count afterwards."""
    now = dt.datetime.now().isoformat(timespec="seconds")
    recs = []
    for _, r in df.iterrows():
        brand, model, desc = _str(r.get("Brand")), _str(r.get("Model")), _str(r.get("Description"))
        if not (brand or model or desc):
            continue
        cur = _str(r.get("Currency")) or "USD"
        recs.append((desc, brand, model, _f(r.get("ListPriceUSD")), _f(r.get("ExUnitCostUSD")),
                     cur if cur in calc.CURRENCIES else "USD", _f(r.get("ShippingPercent")),
                     _f(r.get("UnitCostUSD")), _f(r.get("DefaultUPriceUSD")),
                     _f(r.get("DefaultUPriceSAR")), _catalog_price_date(r.get("PriceUpdatedAt")),
                     int(calc._num(r.get("TimesQuoted"))), now))
    if not recs:
        raise ValueError("No usable rows found - the file needs Brand / Model / Description columns.")
    with _conn() as c:
        c.execute("UPDATE Project_BoQ_Lines SET ItemID=NULL")
        c.execute("DELETE FROM Items_Catalog")
        c.executemany(
            """INSERT OR IGNORE INTO Items_Catalog
                 (Description,Brand,Model,ListPriceUSD,ExUnitCostUSD,Currency,ShippingPercent,
                  UnitCostUSD,DefaultUPriceUSD,DefaultUPriceSAR,PriceUpdatedAt,TimesQuoted,LastSeenAt)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", recs)
        n = c.execute("SELECT COUNT(*) FROM Items_Catalog").fetchone()[0]
        c.commit()
    return n


def _norm_key(x) -> str:
    return " ".join(str(x or "").split()).strip().lower()


def catalog_duplicates() -> list:
    """Groups of catalogue items sharing the same (Model, Description) ignoring case and
    spacing. Returns [{model, description, identical, items:[...]}] for groups of >1 item;
    `identical` means brand/currency/all prices match across the group too."""
    df = catalog_all()
    if df.empty:
        return []
    groups: dict = {}
    for _, r in df.iterrows():
        key = (_norm_key(r.get("Model")), _norm_key(r.get("Description")))
        if not any(key):                       # skip rows with no Model AND no Description
            continue
        groups.setdefault(key, []).append(dict(r))
    cmp_num = ["ListPriceUSD", "ExUnitCostUSD", "ShippingPercent", "UnitCostUSD",
               "DefaultUPriceUSD", "DefaultUPriceSAR"]
    out = []
    for items in groups.values():
        if len(items) < 2:
            continue
        def sig(it):
            return (_norm_key(it.get("Brand")), _norm_key(it.get("Currency")),
                    tuple(round(calc._num(it.get(c)), 2) for c in cmp_num))
        identical = len({sig(it) for it in items}) == 1
        out.append({"model": items[0].get("Model"), "description": items[0].get("Description"),
                    "identical": identical, "items": items})
    out.sort(key=lambda g: (not g["identical"], str(g["description"] or "").lower()))
    return out


def item_to_grid_row(item: dict, area="", system="", qty=1, default_margin=0.0) -> dict:
    """Map a catalogue item onto a fresh grid row, auto-filling cost + margin.

    The historical effective margin (price / landed cost) is pre-filled so the
    Excel formula reproduces the item's usual price; tweak the margin to re-price.
    """
    row = calc.blank_row(area=area, system=system)
    cur = item.get("Currency") or "USD"
    ex_raw = item.get("ExUnitCostUSD") or 0.0          # stored in the item's currency
    ex_usd = calc.to_usd(ex_raw, cur)
    ship = calc.shipping_percent(item.get("ShippingPercent"), ex_usd, item.get("UnitCostUSD"))
    unit = item.get("UnitCostUSD") or calc.roundup(ex_usd * (1 + ship / 100), 0)
    uprice = item.get("DefaultUPriceUSD") or 0.0
    margin = round(uprice / unit, 2) if (unit and uprice) else (default_margin or 0.0)
    row.update({
        "Description": item.get("Description") or "",
        "Brand": item.get("Brand") or "",
        "Model": item.get("Model") or "",
        "Qty": qty,
        "Cur": cur if cur in calc.CURRENCIES else "USD",
        "List Price $": item.get("ListPriceUSD") or 0.0,
        "Ex Unit Cost $": ex_raw,
        "Shipping %": ship,
        "Unit Cost $": item.get("UnitCostUSD") or 0.0,
        "Margin x": margin,
        "U. Price $": uprice,
        "U. Price SAR": item.get("DefaultUPriceSAR") or 0.0,
        "_ItemID": item.get("ItemID"),
    })
    return row


# ---------- Projects ----------

def list_projects() -> pd.DataFrame:
    with _conn() as c:
        rows = c.execute(
            """SELECT ProjectID,ProjectName,ClientName,OfferNo,CreationDate,
                      ConversionFactor,Approved,RevisionNo,BaseName,OptionLabel,Archived
               FROM Projects_Master ORDER BY CreationDate DESC, ProjectID DESC"""
        ).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def family_key(offer_no, project_name) -> str:
    """Groups an offer with its revisions (by offer-number base, else name base)."""
    off = _str(offer_no).strip()
    return base_name(off).lower() if off else base_name(project_name).lower()


def db_counts() -> dict:
    """Lightweight sidebar counts without loading big catalogue/report tables."""
    with _conn() as c:
        rows = c.execute(
            "SELECT OfferNo,ProjectName FROM Projects_Master"
        ).fetchall()
        ncat = c.execute("SELECT COUNT(*) FROM Items_Catalog").fetchone()[0]
    return {
        "project_records": len(rows),
        "project_families": len({family_key(r["OfferNo"], r["ProjectName"]) for r in rows}),
        "catalogue_items": int(ncat or 0),
    }


def approve_offer(project_id: int) -> int:
    """Approve this revision+option. Every OTHER entry in the same offer family
    (losing revisions and options) is auto-archived. Returns # auto-archived."""
    now = dt.datetime.now().isoformat(timespec="seconds")
    meta = project_meta(project_id)
    if not meta:
        return 0
    fam = family_key(meta.get("OfferNo"), meta.get("ProjectName"))
    archived = 0
    with _conn() as c:
        for r in c.execute("SELECT ProjectID,OfferNo,ProjectName,Archived "
                           "FROM Projects_Master").fetchall():
            if r["ProjectID"] != project_id and family_key(r["OfferNo"], r["ProjectName"]) == fam:
                if r["Archived"]:                       # already archived -> just unapprove
                    c.execute("UPDATE Projects_Master SET Approved=0 WHERE ProjectID=?",
                              (r["ProjectID"],))
                else:                                   # auto-archive & remember who did it
                    c.execute("UPDATE Projects_Master SET Approved=0, Archived=1, ArchivedBy=? "
                              "WHERE ProjectID=?", (project_id, r["ProjectID"]))
                    archived += 1
        c.execute("UPDATE Projects_Master SET Approved=1, ApprovedAt=?, Archived=0, ArchivedBy=NULL "
                  "WHERE ProjectID=?", (now, project_id))
        c.commit()
    return archived


def unapprove_offer(project_id: int) -> int:
    """Un-approve, and auto-restore entries this approval had archived. Returns # restored."""
    with _conn() as c:
        restored = c.execute("SELECT COUNT(*) FROM Projects_Master WHERE ArchivedBy=?",
                             (project_id,)).fetchone()[0]
        c.execute("UPDATE Projects_Master SET Archived=0, ArchivedBy=NULL WHERE ArchivedBy=?",
                  (project_id,))
        c.execute("UPDATE Projects_Master SET Approved=0, ApprovedAt=NULL WHERE ProjectID=?",
                  (project_id,))
        c.commit()
    return restored


def archive_project(project_id: int) -> None:
    with _conn() as c:
        c.execute("UPDATE Projects_Master SET Archived=1, ArchivedBy=NULL WHERE ProjectID=?",
                  (project_id,))
        c.commit()


def unarchive_project(project_id: int) -> None:
    with _conn() as c:
        c.execute("UPDATE Projects_Master SET Archived=0, ArchivedBy=NULL WHERE ProjectID=?",
                  (project_id,))
        c.commit()


# ---------- Deletion ----------

def deletion_ids(project_id: int, scope: str) -> list[int]:
    """ProjectIDs affected by deleting at `scope` ('option' | 'revision' | 'offer')."""
    meta = project_meta(project_id)
    if not meta:
        return []
    if scope == "option":
        return [int(project_id)]
    fam = family_key(meta.get("OfferNo"), meta.get("ProjectName"))
    rev = meta.get("RevisionNo") or 0
    ids = []
    with _conn() as c:
        for r in c.execute("SELECT ProjectID,OfferNo,ProjectName,RevisionNo "
                           "FROM Projects_Master").fetchall():
            if family_key(r["OfferNo"], r["ProjectName"]) == fam:
                if scope == "offer" or (r["RevisionNo"] or 0) == rev:
                    ids.append(int(r["ProjectID"]))
    return ids


def delete_projects(ids: list[int]) -> int:
    """Delete the given projects and their lines/sheets (explicit child cleanup)."""
    ids = [int(i) for i in ids]
    if not ids:
        return 0
    qs = ",".join("?" * len(ids))
    with _conn() as c:
        c.execute(f"DELETE FROM Project_BoQ_Lines WHERE ProjectID IN ({qs})", ids)
        c.execute(f"DELETE FROM Project_Sheets   WHERE ProjectID IN ({qs})", ids)
        c.execute(f"DELETE FROM Projects_Master  WHERE ProjectID IN ({qs})", ids)
        c.commit()
    return len(ids)


def list_systems(project_id: int) -> list[str]:
    with _conn() as c:
        rows = c.execute(
            "SELECT DISTINCT SheetName FROM Project_Sheets WHERE ProjectID=? ORDER BY SheetID",
            (project_id,),
        ).fetchall()
    return [r["SheetName"] for r in rows]


def load_project_grid(project_id: int, sheet_name: str | None = None) -> pd.DataFrame:
    """Load a project's lines into a UI/PDF-ready grid DataFrame."""
    sql = """SELECT Area,System,Description,Brand,Model,Qty,
                    ListPriceUSD,ExUnitCostUSD,Currency,ShippingPercent,FinalUnitCostUSD,TotalCostUSD,
                    FinalUPriceUSD,TPriceUSD,FinalUPriceSAR,TPriceSAR,MarginExtra,
                    LineType,ItemID
             FROM Project_BoQ_Lines l
             JOIN Project_Sheets s ON l.SheetID=s.SheetID
             WHERE l.ProjectID=? AND l.LineType NOT IN ('spare','discount')"""
    params = [project_id]
    if sheet_name:
        sql += " AND s.SheetName=?"
        params.append(sheet_name)
    sql += " ORDER BY l.RowOrder"
    with _conn() as c:
        rows = [dict(r) for r in c.execute(sql, params).fetchall()]
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=calc.GRID_COLUMNS + ["LineType", "_ItemID"])
    df = df.rename(columns={
        "ListPriceUSD": "List Price $", "ExUnitCostUSD": "Ex Unit Cost $",
        "Currency": "Cur", "ShippingPercent": "Shipping %",
        "FinalUnitCostUSD": "Unit Cost $", "TotalCostUSD": "Total Cost $",
        "FinalUPriceUSD": "U. Price $", "TPriceUSD": "T. Price $",
        "FinalUPriceSAR": "U. Price SAR", "TPriceSAR": "T. Price SAR",
        "ItemID": "_ItemID",
    })
    df["Cur"] = df["Cur"].apply(lambda v: v if str(v) in calc.CURRENCIES else "USD")
    # effective margin = U. Price $ / Unit Cost $ (read-only, for display)
    uc = df["Unit Cost $"].map(calc._num)
    up = df["U. Price $"].map(calc._num)
    df["Shipping %"] = [
        calc.shipping_percent(ship, ex, unit)
        for ship, ex, unit in zip(df["Shipping %"], df["Ex Unit Cost $"], df["Unit Cost $"])
    ]
    df["Margin x"] = (up / uc.where(uc > 0)).round(2).fillna(0.0)
    return df[[c for c in calc.GRID_COLUMNS] + ["LineType", "_ItemID"]]


# ---------------- Finance (per approved offer) ----------------

def list_approved_offers() -> list[dict]:
    """Approved, non-archived offers - the choices for the Finance tab."""
    with _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT ProjectID,ProjectName,OfferNo,OptionLabel,RevisionNo,ApprovedAt "
            "FROM Projects_Master WHERE Approved=1 AND IFNULL(Archived,0)=0 "
            "ORDER BY ApprovedAt DESC, ProjectID DESC")]


def offer_grand_totals(project_ids: list[int]) -> dict[int, float]:
    """Grand totals for many offers using one aggregate query."""
    ids = list(dict.fromkeys(int(pid) for pid in project_ids if pid is not None))
    if not ids:
        return {}
    qs = ",".join("?" * len(ids))
    with _conn() as c:
        rows = c.execute(
            f"""SELECT p.ProjectID,
                       IFNULL(p.DiscountAmount,0) AS DiscountAmount,
                       IFNULL(SUM(l.TPriceSAR),0) AS SubtotalSAR
                FROM Projects_Master p
                LEFT JOIN Project_BoQ_Lines l
                  ON l.ProjectID = p.ProjectID
                 AND IFNULL(l.LineType,'item') NOT IN ('spare','discount')
                WHERE p.ProjectID IN ({qs})
                GROUP BY p.ProjectID, p.DiscountAmount""",
            ids,
        ).fetchall()
    out = {}
    for r in rows:
        subtotal = float(r["SubtotalSAR"] or 0)
        discount = min(abs(float(r["DiscountAmount"] or 0)), subtotal)
        out[int(r["ProjectID"])] = round((subtotal - discount) * (1 + calc.VAT_RATE), 2)
    return out


def offer_grand_total(project_id: int) -> float:
    """Grand Total (SAR, incl. VAT & discount) of an offer - the finance baseline."""
    return offer_grand_totals([project_id]).get(int(project_id), 0.0)


def get_finance(project_id: int) -> tuple[list[dict], list[dict]]:
    """Return (payments, purchases) rows stored for an offer."""
    with _conn() as c:
        pays = [dict(r) for r in c.execute(
            "SELECT Description,AmountSAR,InvoiceNo FROM Finance_Payments "
            "WHERE ProjectID=? ORDER BY RowOrder", (project_id,))]
        purs = [dict(r) for r in c.execute(
            "SELECT Description,AmountSAR,PORef FROM Finance_Purchases "
            "WHERE ProjectID=? ORDER BY RowOrder", (project_id,))]
    return pays, purs


def save_finance(project_id: int, payments: list[dict], purchases: list[dict]) -> None:
    """Replace an offer's finance rows. Blank rows (no text and zero amount) are dropped."""
    with _conn() as c:
        c.execute("DELETE FROM Finance_Payments WHERE ProjectID=?", (project_id,))
        c.execute("DELETE FROM Finance_Purchases WHERE ProjectID=?", (project_id,))
        order = 0
        for r in payments:
            desc = str(r.get("Description") or "").strip()
            amt = calc._num(r.get("Amount (SAR)"))
            inv = str(r.get("Invoice #") or "").strip()
            if not desc and amt == 0 and not inv:
                continue
            c.execute("INSERT INTO Finance_Payments(ProjectID,RowOrder,Description,AmountSAR,InvoiceNo)"
                      " VALUES(?,?,?,?,?)", (project_id, order, desc, amt, inv))
            order += 1
        order = 0
        for r in purchases:
            desc = str(r.get("Description") or "").strip()
            amt = calc._num(r.get("Cost (SAR)"))
            po = str(r.get("PO #") or "").strip()
            if not desc and amt == 0 and not po:
                continue
            c.execute("INSERT INTO Finance_Purchases(ProjectID,RowOrder,Description,AmountSAR,PORef)"
                      " VALUES(?,?,?,?,?)", (project_id, order, desc, amt, po))
            order += 1
        c.commit()


def load_tracking(project_id: int, sheet_name: str | None = None) -> pd.DataFrame:
    """Line items of an offer with their procurement tracking flags (for approved offers)."""
    sql = """SELECT l.LineID, l.System, l.Description, l.Brand, l.Model, l.Qty,
                    l.FinalUPriceSAR, l.TPriceSAR, IFNULL(l.PONumber,'') PONumber,
                    IFNULL(l.DeliveryNote,'') DeliveryNote,
                    IFNULL(l.Paid,0) Paid, IFNULL(l.PaidAt,'') PaidAt,
                    IFNULL(l.Received,0) Received, IFNULL(l.ReceivedAt,'') ReceivedAt,
                    IFNULL(l.ReceivedQty,0) ReceivedQty,
                    IFNULL(l.Delivered,0) Delivered, IFNULL(l.DeliveredAt,'') DeliveredAt,
                    IFNULL(l.DeliveredQty,0) DeliveredQty
             FROM Project_BoQ_Lines l JOIN Project_Sheets s ON l.SheetID=s.SheetID
             WHERE l.ProjectID=? AND l.LineType IN ('item','service')"""
    params = [project_id]
    if sheet_name:
        sql += " AND s.SheetName=?"
        params.append(sheet_name)
    sql += " ORDER BY l.RowOrder"
    with _conn() as c:
        rows = [dict(r) for r in c.execute(sql, params).fetchall()]
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.rename(columns={"FinalUPriceSAR": "U. Price SAR", "TPriceSAR": "T. Price SAR"})
        for col in ("Paid", "Received", "Delivered"):
            df[col] = df[col].astype(bool)
    return df


def update_tracking(rows) -> int:
    """rows: iterable of
    (line_id, paid, received, delivered[, po_number, delivery_note,
     paid_at, received_at, delivered_at, received_qty, delivered_qty]).
    """
    rows = list(rows)
    with _conn() as c:
        for r in rows:
            lid, paid, rec, deliv = r[0], r[1], r[2], r[3]
            lid = int(lid)
            paid, rec, deliv = bool(paid), bool(rec), bool(deliv)
            old = c.execute(
                "SELECT Paid,Received,Delivered,PaidAt,ReceivedAt,DeliveredAt,"
                "Qty,ReceivedQty,DeliveredQty "
                "FROM Project_BoQ_Lines WHERE LineID=?", (lid,)
            ).fetchone()
            old = dict(old) if old else {}
            line_qty = _bounded_qty(old.get("Qty"), None)
            rec_qty = _bounded_qty(r[9] if len(r) > 9 else (line_qty if rec else 0), line_qty)
            deliv_qty = _bounded_qty(r[10] if len(r) > 10 else (line_qty if deliv else 0), line_qty)
            rec = rec or rec_qty > 0
            deliv = deliv or deliv_qty > 0
            if rec and rec_qty <= 0 and line_qty > 0:
                rec_qty = line_qty
            if deliv and deliv_qty <= 0 and line_qty > 0:
                deliv_qty = line_qty
            if not rec:
                rec_qty = 0.0
            if not deliv:
                deliv_qty = 0.0
            now = dt.datetime.now().isoformat(timespec="minutes")
            new_paid_at = (r[6] if len(r) > 6 else "") or ""
            new_rec_at = (r[7] if len(r) > 7 else "") or ""
            new_deliv_at = (r[8] if len(r) > 8 else "") or ""
            paid_at = old.get("PaidAt") if paid and old.get("Paid") and old.get("PaidAt") else ((new_paid_at or now) if paid else None)
            rec_at = old.get("ReceivedAt") if rec and old.get("Received") and old.get("ReceivedAt") else ((new_rec_at or now) if rec else None)
            deliv_at = old.get("DeliveredAt") if deliv and old.get("Delivered") and old.get("DeliveredAt") else ((new_deliv_at or now) if deliv else None)
            po = (r[4] if len(r) > 4 else "") or ""
            delivery_note = (r[5] if len(r) > 5 else "") or ""
            c.execute(
                """UPDATE Project_BoQ_Lines
                   SET Paid=?, Received=?, Delivered=?,
                       ReceivedQty=?, DeliveredQty=?,
                       PONumber=?, DeliveryNote=?,
                       PaidAt=?, ReceivedAt=?, DeliveredAt=?
                   WHERE LineID=?""",
                (int(paid), int(rec), int(deliv),
                 rec_qty, deliv_qty,
                 po.strip() or None, delivery_note.strip() or None,
                 paid_at, rec_at, deliv_at, lid),
            )
        c.commit()
    return len(rows)


def project_meta(project_id: int) -> dict:
    with _conn() as c:
        r = c.execute(
            "SELECT * FROM Projects_Master WHERE ProjectID=?", (project_id,)
        ).fetchone()
    return dict(r) if r else {}


def save_offer(name, client, contact, offer_no, system_suffix, grid: pd.DataFrame,
               discount_sar=0.0, factors=(None, None, None),
               sales_person=None, presales_engineer=None, project_manager=None,
               revision_no=0, base=None, terms=None, option_label="",
               project_sheet_info=None, phone="", contractor="", region="") -> int:
    """Persist a NEW offer (and its lines) created in the interface."""
    discount_sar = _discount_amount(discount_sar)
    now = dt.datetime.now().isoformat(timespec="seconds")
    today = dt.date.today().isoformat()
    base = base or base_name(name)
    terms_json = json.dumps({k: terms.get(k) for k in TERMS_KEYS}) if terms else None
    ps_json = (json.dumps({k: project_sheet_info.get(k) for k in PROJECT_SHEET_KEYS})
               if project_sheet_info else None)
    with _conn() as c:
        cur = c.execute(
            """INSERT INTO Projects_Master
                 (ProjectName,ClientName,ContactName,ContactPhone,
                  Contractor,Region,SalesPerson,PresalesEngineer,ProjectManager,
                  OfferNo,CreationDate,DiscountAmount,ConversionFactor,SourceFile,IngestedAt,
                  RevisionNo,BaseName,OfferTerms,ProjectSheetInfo,OptionLabel)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (name, client, contact, _str(phone), _str(contractor), _str(region),
             sales_person, presales_engineer, project_manager, offer_no,
             today, discount_sar, factors[0], f"app://offer/{name}/{now}", now, revision_no, base,
             terms_json, ps_json, option_label or ""),
        )
        pid = cur.lastrowid
        _write_sheet_and_lines(c, pid, system_suffix, discount_sar, factors, grid)
        c.commit()
    return pid


def _write_sheet_and_lines(c, pid, system_suffix, discount_sar, factors, grid) -> int:
    """Insert one Project_Sheets row + all Project_BoQ_Lines for `grid`. Returns SheetID."""
    discount_sar = _discount_amount(discount_sar)
    scur = c.execute(
        """INSERT INTO Project_Sheets
             (ProjectID,SheetName,SystemSuffix,DiscountAmount,Factor1,Factor2,Factor3)
           VALUES (?,?,?,?,?,?,?)""",
        (pid, f"BOQ {system_suffix}", system_suffix, discount_sar, *factors),
    )
    sid = scur.lastrowid
    discount_rows = []
    item_rows = []
    for order, (_, r) in enumerate(grid.iterrows(), 1):
        lt = str(r.get("LineType", "item"))
        if lt == "discount":
            discount_rows.append(
                (pid, sid, order, "Discount", "discount", _discount_line_amount(r.get("T. Price SAR")))
            )
            continue
        cur_code = str(r.get("Cur") or "USD")
        if cur_code not in calc.CURRENCIES:
            cur_code = "USD"
        item_rows.append(
            (pid, sid, r.get("_ItemID"), order, r.get("Area"), r.get("System"),
             r.get("Description"), r.get("Brand"), r.get("Model"), _f(r.get("Qty")),
             _f(r.get("List Price $")), _f(r.get("Ex Unit Cost $")), cur_code,
             _f(r.get("Shipping %")), _f(r.get("Unit Cost $")), _f(r.get("Total Cost $")),
             _f(r.get("U. Price $")), _f(r.get("T. Price $")),
             _f(r.get("U. Price SAR")), _f(r.get("T. Price SAR")),
             _f(r.get("Margin x")), lt)
        )
    if discount_rows:
        c.executemany(
            """INSERT INTO Project_BoQ_Lines
                 (ProjectID,SheetID,RowOrder,Description,LineType,TPriceSAR)
               VALUES (?,?,?,?,?,?)""",
            discount_rows,
        )
    if item_rows:
        c.executemany(
            """INSERT INTO Project_BoQ_Lines
                 (ProjectID,SheetID,ItemID,RowOrder,Area,System,Description,Brand,Model,
                  Qty,ListPriceUSD,ExUnitCostUSD,Currency,ShippingPercent,FinalUnitCostUSD,TotalCostUSD,
                  FinalUPriceUSD,TPriceUSD,FinalUPriceSAR,TPriceSAR,MarginExtra,LineType)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            item_rows,
        )
    return sid


def update_offer(base_project_id: int, grid: pd.DataFrame, discount_sar=0.0,
                 factors=(None, None, None), system_suffix="LCS", terms=None,
                 project_sheet_info=None, header=None, option_label=None) -> int:
    """Overwrite an existing revision/option IN PLACE - same ProjectID, Offer #,
    revision, option, approval. Replaces lines, discount, terms; updates the offer
    header (client/project/contact/sales/pre-sales/PM) when `header` is given."""
    discount_sar = _discount_amount(discount_sar)
    terms_json = json.dumps({k: terms.get(k) for k in TERMS_KEYS}) if terms else None
    ps_json = (json.dumps({k: project_sheet_info.get(k) for k in PROJECT_SHEET_KEYS})
               if project_sheet_info else None)
    with _conn() as c:
        c.execute(
            "UPDATE Projects_Master SET DiscountAmount=?, ConversionFactor=?, "
            "OfferTerms=COALESCE(?, OfferTerms), "
            "ProjectSheetInfo=COALESCE(?, ProjectSheetInfo) WHERE ProjectID=?",
            (discount_sar, factors[0], terms_json, ps_json, base_project_id))
        if header is not None:
            c.execute(
                "UPDATE Projects_Master SET ProjectName=COALESCE(?,ProjectName), ClientName=?, "
                "ContactName=?, ContactPhone=?, Contractor=?, Region=?, "
                "SalesPerson=?, PresalesEngineer=?, ProjectManager=? "
                "WHERE ProjectID=?",
                ((header.get("project") or "").strip() or None, _str(header.get("client")),
                 _str(header.get("contact")), _str(header.get("phone")),
                 _str(header.get("contractor")), _str(header.get("region")),
                 _str(header.get("sales")), _str(header.get("presales")),
                 _str(header.get("pm")), base_project_id))
        if option_label is not None:
            c.execute("UPDATE Projects_Master SET OptionLabel=? WHERE ProjectID=?",
                      (_str(option_label), base_project_id))
        c.execute("DELETE FROM Project_BoQ_Lines WHERE ProjectID=?", (base_project_id,))
        c.execute("DELETE FROM Project_Sheets WHERE ProjectID=?", (base_project_id,))
        _write_sheet_and_lines(c, base_project_id, system_suffix, discount_sar, factors, grid)
        c.commit()
    return base_project_id


def _header_fields(meta, header):
    """Resolve edited header fields, falling back to the stored offer when absent.

    Returns client, contact, phone, contractor, region, sales, presales, pm.
    """
    if header is None:
        return (meta.get("ClientName"), meta.get("ContactName"), meta.get("ContactPhone"),
                meta.get("Contractor"), meta.get("Region"), meta.get("SalesPerson"),
                meta.get("PresalesEngineer"), meta.get("ProjectManager"))
    return (_str(header.get("client")), _str(header.get("contact")), _str(header.get("phone")),
            _str(header.get("contractor")), _str(header.get("region")), _str(header.get("sales")),
            _str(header.get("presales")), _str(header.get("pm")))


def save_revision(base_project_id: int, grid: pd.DataFrame, discount_sar=0.0,
                  factors=(None, None, None), system_suffix="LCS",
                  terms=None, option_label=None, project_sheet_info=None,
                  header=None) -> tuple[int, str, int]:
    """Save `grid` as the next revision of an existing offer.

    Returns (new_project_id, new_name, revision_no). The new offer is named
    '<base> -Rev.<n>' and its Offer # gets the same '-Rev.<n>' suffix.
    """
    meta = project_meta(base_project_id)
    base = meta.get("BaseName") or base_name(meta.get("ProjectName") or "Offer")
    rev = next_revision(base)
    opt = option_label if option_label is not None else (meta.get("OptionLabel") or "")
    sep, tok = revision_separator(), revision_token(rev)
    proj_override = (header.get("project") or "").strip() if header is not None else ""
    name = (proj_override or f"{base}{sep}{tok}") + (f" ({opt})" if opt else "")
    offer = meta.get("OfferNo")
    offer_rev = f"{base_name(offer)}{sep}{tok}" if offer else None
    client, contact, phone, contractor, region, sales, presales, pm = _header_fields(meta, header)
    pid = save_offer(
        name=name, client=client, contact=contact,
        offer_no=offer_rev, system_suffix=system_suffix, grid=grid,
        discount_sar=discount_sar, factors=factors,
        sales_person=sales, presales_engineer=presales, project_manager=pm,
        revision_no=rev, base=base, terms=terms if terms is not None else load_terms(meta),
        project_sheet_info=(project_sheet_info if project_sheet_info is not None
                            else load_project_sheet_info(meta)),
        option_label=opt, phone=phone, contractor=contractor, region=region)
    return pid, name, rev


def save_option(base_project_id: int, grid: pd.DataFrame, option_label: str,
                discount_sar=0.0, factors=(None, None, None), system_suffix="LCS",
                terms=None, project_sheet_info=None, header=None) -> tuple[int, str, int]:
    """Save `grid` as another OPTION of the SAME revision (e.g. Dynalite vs KNX).

    Keeps the source revision number and Offer #; only the option label differs.
    Returns (new_project_id, new_name, revision_no).
    """
    meta = project_meta(base_project_id)
    base = meta.get("BaseName") or base_name(meta.get("ProjectName") or "Offer")
    rev = int(meta.get("RevisionNo") or 0)
    opt = (option_label or "").strip() or "Option"
    proj_override = (header.get("project") or "").strip() if header is not None else ""
    stem = proj_override or (f"{base}{revision_separator()}{revision_token(rev)}" if rev else base)
    name = f"{stem} ({opt})"
    client, contact, phone, contractor, region, sales, presales, pm = _header_fields(meta, header)
    pid = save_offer(
        name=name, client=client, contact=contact,
        offer_no=meta.get("OfferNo"), system_suffix=system_suffix, grid=grid,
        discount_sar=discount_sar, factors=factors,
        sales_person=sales, presales_engineer=presales, project_manager=pm,
        revision_no=rev, base=base, terms=terms if terms is not None else load_terms(meta),
        project_sheet_info=(project_sheet_info if project_sheet_info is not None
                            else load_project_sheet_info(meta)),
        option_label=opt, phone=phone, contractor=contractor, region=region)
    return pid, name, rev


def _f(x):
    try:
        if x is None or x == "":
            return None
        return float(x)
    except (TypeError, ValueError):
        return None


def _bounded_qty(value, max_qty=None) -> float:
    qty = _f(value)
    qty = 0.0 if qty is None else max(qty, 0.0)
    if max_qty is not None and max_qty > 0:
        qty = min(qty, max_qty)
    return round(qty, 4)


def _discount_amount(x) -> float:
    v = _f(x)
    return abs(v) if v is not None else 0.0


def _discount_line_amount(x):
    v = _f(x)
    return abs(v) if v is not None else None
