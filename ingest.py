"""
Folder-wide BoQ ingestion engine.

Scans the assigned project directory, finds every `BOQ <System>` sheet (and its
matching `Quotation <System>` sheet for client metadata), parses the canonical
grid, and upserts into the SQLite master database.

Idempotent: re-running re-ingests each file cleanly (old rows for that file are
removed first), so it is safe to run repeatedly as the folder grows.

Usage:
    python ingest.py                 # ingest the default ROOT folder
    python ingest.py "C:\\some\\path" # ingest a different folder
"""
import os
import re
import sys
import glob
import json
import datetime as dt
import warnings

warnings.simplefilter("ignore")
import openpyxl

import calc
import db as dbmod

DEFAULT_ROOT = r"J:\My Drive\1-Projects"

# Canonical BOQ columns -> normalized keys. Detection matches these labels
# (lowercased, punctuation-stripped) against the header row.
HEADER_MAP = {
    "area": "area",
    "system": "system",
    "description": "description",
    "brand": "brand",
    "model": "model",
    "qty": "qty",
    "quantity": "qty",
    "list price": "list_price",
    "list price $": "list_price",
    "ex unit cost": "ex_unit_cost",
    "ex unit cost $": "ex_unit_cost",
    "unit cost": "unit_cost",
    "unit cost $": "unit_cost",
    "total cost": "total_cost",
    "total cost $": "total_cost",
    "u price": "u_price",
    "u price $": "u_price",
    "t price": "t_price",
    "t price $": "t_price",
    "u price sar": "u_price_sar",
    "t price sar": "t_price_sar",
}

REQUIRED_KEYS = {"description", "qty"}  # minimum to trust a header row


def norm(v):
    if v is None:
        return ""
    s = str(v).strip().lower()
    s = s.replace("$", "").replace(".", " ").replace("/", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def find_header(ws, scan_rows=12):
    """Return (header_row_index, {key: col_index}) or (None, None)."""
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        mapping = {}
        for c in range(1, ws.max_column + 1):
            key = HEADER_MAP.get(norm(ws.cell(row=r, column=c).value))
            if key and key not in mapping:
                mapping[key] = c
        if REQUIRED_KEYS.issubset(mapping.keys()) and "model" in mapping:
            return r, mapping
    return None, None


def cell(ws, r, colmap, key):
    c = colmap.get(key)
    return ws.cell(row=r, column=c).value if c else None


def parse_boq_sheet(ws):
    """Parse one BOQ sheet -> (lines, sheet_meta)."""
    hdr, colmap = find_header(ws)
    if hdr is None:
        return None, None

    last_data_col = max(colmap.values())
    lines = []
    factors = []
    discount = 0.0
    order = 0
    in_summary = False   # True once the totals/factor block starts -> rows below = spare
    empty_run = 0        # consecutive blank rows (the 3-4 row gap before spares)
    summary_subtotal = None  # the BOQ's own offer subtotal (first totals row, SAR)

    for r in range(hdr + 1, ws.max_row + 1):
        desc = cell(ws, r, colmap, "description")
        desc_s = norm(desc)
        qty = to_num(cell(ws, r, colmap, "qty"))
        model = cell(ws, r, colmap, "model")
        brand = cell(ws, r, colmap, "brand")

        # Discount row (still part of the offer)
        if desc_s == "discount":
            dval = to_num(cell(ws, r, colmap, "t_price_sar")) \
                or to_num(cell(ws, r, colmap, "total_cost")) \
                or to_num(cell(ws, r, colmap, "t_price"))
            if dval:
                dval = abs(dval)
                discount = dval
            order += 1
            lines.append({"_type": "discount", "_order": order, "description": "Discount",
                          "qty": None, "t_price_sar": dval})
            empty_run = 0
            continue

        has_label = bool((desc and str(desc).strip()) or (model and str(model).strip()))

        # No item label and no qty -> blank / totals / conversion-factor row.
        # The first such non-blank row (or a 3+ row gap) marks the end of the
        # offer; everything below it is a parked "spare" device.
        if not has_label and qty is None:
            row_vals = [to_num(ws.cell(row=r, column=c).value)
                        for c in range(1, last_data_col + 2)]
            nums = [v for v in row_vals if v is not None]
            if not nums:                                # truly empty row
                empty_run += 1
                if empty_run >= 3 and order > 0:
                    in_summary = True
                continue
            empty_run = 0
            if len(nums) == 1 and 1.0 <= nums[0] <= 2.5:
                factors.append(nums[0])                 # 1.69 / 1.49 / 1.42
            elif summary_subtotal is None:              # first totals row = offer subtotal
                summary_subtotal = to_num(cell(ws, r, colmap, "t_price_sar"))
            in_summary = True                           # totals/factor row ends the offer
            continue

        # Item row.
        if has_label and qty is not None:
            empty_run = 0
            if in_summary:
                ltype = "spare"                         # parked below the block
            elif (qty == 1 and not to_num(cell(ws, r, colmap, "ex_unit_cost"))
                  and norm(model) in ("sws", "")):
                ltype = "service"
            else:
                ltype = "item"
            order += 1
            ex_unit_cost = to_num(cell(ws, r, colmap, "ex_unit_cost"))
            unit_cost = to_num(cell(ws, r, colmap, "unit_cost"))
            lines.append({
                "_type": ltype, "_order": order,
                "area": cell(ws, r, colmap, "area"),
                "system": cell(ws, r, colmap, "system"),
                "description": str(desc).strip() if desc else None,
                "brand": str(brand).strip() if brand else None,
                "model": str(model).strip() if model else None,
                "qty": qty,
                "list_price": to_num(cell(ws, r, colmap, "list_price")),
                "ex_unit_cost": ex_unit_cost,
                "shipping_percent": calc.infer_shipping_percent(ex_unit_cost, unit_cost),
                "unit_cost": unit_cost,
                "total_cost": to_num(cell(ws, r, colmap, "total_cost")),
                "u_price": to_num(cell(ws, r, colmap, "u_price")),
                "t_price": to_num(cell(ws, r, colmap, "t_price")),
                "u_price_sar": to_num(cell(ws, r, colmap, "u_price_sar")),
                "t_price_sar": to_num(cell(ws, r, colmap, "t_price_sar")),
                "margin": _trailing_margin(ws, r, last_data_col),
            })
            continue

        # has_label but no qty (e.g. a section sub-header) -> ignore.
        empty_run = 0

    meta = {"discount": discount, "factors": factors[:3], "subtotal_sar": summary_subtotal}
    return lines, meta


def _trailing_margin(ws, r, last_data_col):
    """The analytic margin/extra value that sits just past the named columns."""
    for c in range(last_data_col + 1, last_data_col + 4):
        v = to_num(ws.cell(row=r, column=c).value)
        if v is not None:
            return v
    return None


def _find_quotation_sheet(wb, suffix):
    name = f"Quotation {suffix}".strip()
    if name in wb.sheetnames:
        return wb[name]
    for s in wb.sheetnames:
        if s.lower().startswith("quotation") and suffix.lower() in s.lower():
            return wb[s]
    for s in wb.sheetnames:                      # fall back to any quotation sheet
        if s.lower().startswith("quotation"):
            return wb[s]
    return None


def parse_quotation_meta(wb, suffix):
    """Pull client/project/offer metadata from the matching Quotation sheet."""
    ws = _find_quotation_sheet(wb, suffix)
    meta = {"client": None, "project": None, "contact": None, "offer": None, "date": None}
    if ws is None:
        return meta
    labels = {
        "client name": "client", "project name": "project",
        "contact": "contact", "offer #": "offer", "offer no": "offer",
        "date": "date", "date:": "date", "phone": None, "billed to": None,
    }
    for r in range(1, min(ws.max_row, 20) + 1):
        for c in range(1, min(ws.max_column, 9) + 1):
            key = norm(ws.cell(row=r, column=c).value)
            if key in labels and labels[key] is not None:
                # value is the next non-empty cell to the right, but stop if we
                # hit another label cell (blank value field -> leave as None).
                for cc in range(c + 1, min(ws.max_column, 9) + 1):
                    val = ws.cell(row=r, column=cc).value
                    if val in (None, ""):
                        continue
                    if norm(val) in labels:          # adjacent label, not a value
                        break
                    meta[labels[key]] = val
                    break
    return meta


# Quotation note labels -> offer-terms keys (longest first for greedy matching).
QTERM_LABELS = sorted([
    ("special notes and instructions", "notes"), ("special notes", "notes"),
    ("pre-requirements", "prerequisites"), ("pre requirements", "prerequisites"),
    ("prerequisites", "prerequisites"),
    ("payment terms", "payment"), ("payment", "payment"),
    ("scope of work", "scope"), ("scope", "scope"),
    ("exclusions", "exclusions"), ("exclusion", "exclusions"),
    ("delivery time", "delivery"), ("delivery", "delivery"),
    ("offer validity", "validity"), ("validity", "validity"),
    ("warranty", "notes"), ("remarks", "notes"),
    ("system", "system_note"),
], key=lambda x: len(x[0]), reverse=True)


def parse_quotation_terms(wb, suffix):
    """Extract the offer Subject (title), greeting and note blocks (scope, payment,
    exclusions, delivery, validity, system, ...) from the Quotation sheet."""
    ws = _find_quotation_sheet(wb, suffix)
    terms = {}
    if ws is None:
        return terms
    mc = min(ws.max_column, 10)

    hdr = None
    for r in range(1, min(ws.max_row, 30) + 1):
        rowtext = " | ".join(norm(ws.cell(row=r, column=c).value)
                             for c in range(1, mc + 1) if ws.cell(row=r, column=c).value)
        if "description" in rowtext and ("unit price" in rowtext or "qty" in rowtext):
            hdr = r
            break

    # Subject (title cell ending in 'Offer') + greeting, above the items table.
    for r in range(1, (hdr or 12)):
        for c in range(1, mc + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            vn = norm(v)
            if not terms.get("subject") and vn.endswith("offer") and 3 < len(vn) < 55:
                terms["subject"] = str(v).strip()
            if not terms.get("greeting") and vn.startswith("dear"):
                terms["greeting"] = str(v).strip()

    # Note blocks below the items table, left columns only.
    for r in range((hdr or 14) + 1, ws.max_row + 1):
        for c in range(1, 6):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str) or not v.strip():
                continue
            vn = norm(v)
            for lab, key in QTERM_LABELS:
                if vn.startswith(lab) and not terms.get(key):
                    pat = re.compile(r"^\s*" + r"\s+".join(re.escape(w) for w in lab.split())
                                     + r"\s*[:\-]?\s*", re.I)
                    val = pat.sub("", str(v).strip(), count=1).strip()
                    if val:
                        terms[key] = val
                    break
    return terms


def system_suffix(sheet_name):
    return re.sub(r"^\s*boq\s*", "", sheet_name, flags=re.I).strip()


def upsert_item(conn, ln, src_file, now):
    """Upsert into Items_Catalog by (Brand, Model, Description); refresh defaults."""
    if not (ln.get("description") or ln.get("model")):
        return None
    cur = conn.execute(
        "SELECT ItemID FROM Items_Catalog WHERE IFNULL(Brand,'')=? AND IFNULL(Model,'')=? AND IFNULL(Description,'')=?",
        (ln.get("brand") or "", ln.get("model") or "", ln.get("description") or ""),
    )
    row = cur.fetchone()
    if row:
        iid = row["ItemID"]
        conn.execute(
            """UPDATE Items_Catalog SET
                 ListPriceUSD=COALESCE(?,ListPriceUSD),
                 ExUnitCostUSD=COALESCE(?,ExUnitCostUSD),
                 ShippingPercent=COALESCE(?,ShippingPercent),
                 UnitCostUSD=COALESCE(?,UnitCostUSD),
                 DefaultUPriceUSD=COALESCE(?,DefaultUPriceUSD),
                 DefaultUPriceSAR=COALESCE(?,DefaultUPriceSAR),
                 TimesQuoted=TimesQuoted+1, LastSeenFile=?, LastSeenAt=?
               WHERE ItemID=?""",
            (ln.get("list_price"), ln.get("ex_unit_cost"), ln.get("shipping_percent"), ln.get("unit_cost"),
             ln.get("u_price"), ln.get("u_price_sar"), src_file, now, iid),
        )
        return iid
    cur = conn.execute(
        """INSERT INTO Items_Catalog
             (Description,Brand,Model,ListPriceUSD,ExUnitCostUSD,ShippingPercent,UnitCostUSD,
              DefaultUPriceUSD,DefaultUPriceSAR,TimesQuoted,LastSeenFile,LastSeenAt)
           VALUES (?,?,?,?,?,?,?,?,?,1,?,?)""",
        (ln.get("description"), ln.get("brand"), ln.get("model"), ln.get("list_price"),
         ln.get("ex_unit_cost"), ln.get("shipping_percent"), ln.get("unit_cost"), ln.get("u_price"),
         ln.get("u_price_sar"), src_file, now),
    )
    return cur.lastrowid


def ingest_file(conn, path, stats):
    now = dt.datetime.now().isoformat(timespec="seconds")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        stats["errors"].append((os.path.basename(path), repr(e)[:90]))
        return

    boq_sheets = [s for s in wb.sheetnames if s.lower().strip().startswith("boq")]
    if not boq_sheets:
        wb.close()
        stats["skipped_no_boq"] += 1
        return

    # remove any prior ingest of this exact file (idempotent)
    conn.execute("DELETE FROM Projects_Master WHERE SourceFile=?", (path,))

    # project-level metadata from the first system's Quotation sheet
    first_suffix = system_suffix(boq_sheets[0])
    qmeta = parse_quotation_meta(wb, first_suffix)
    qterms = parse_quotation_terms(wb, first_suffix)
    terms_json = json.dumps(qterms) if qterms else None
    cdate = None
    if isinstance(qmeta.get("date"), (dt.date, dt.datetime)):
        cdate = qmeta["date"].date().isoformat() if isinstance(qmeta["date"], dt.datetime) else qmeta["date"].isoformat()
    if not cdate:
        cdate = dt.datetime.fromtimestamp(os.path.getmtime(path)).date().isoformat()
    project_name = str(qmeta.get("project") or os.path.splitext(os.path.basename(path))[0]).strip()

    cur = conn.execute(
        """INSERT INTO Projects_Master
             (ProjectName,ClientName,ContactName,OfferNo,CreationDate,SourceFile,IngestedAt,OfferTerms)
           VALUES (?,?,?,?,?,?,?,?)""",
        (project_name, _s(qmeta.get("client")), _s(qmeta.get("contact")),
         _s(qmeta.get("offer")), cdate, path, now, terms_json),
    )
    pid = cur.lastrowid
    primary_discount = None
    primary_factor = None

    for sheet in boq_sheets:
        suffix = system_suffix(sheet)
        lines, meta = parse_boq_sheet(wb[sheet])
        if lines is None:
            stats["sheets_unparsed"] += 1
            continue
        f = (meta["factors"] + [None, None, None])[:3]
        scur = conn.execute(
            """INSERT INTO Project_Sheets
                 (ProjectID,SheetName,SystemSuffix,DiscountAmount,Factor1,Factor2,Factor3,SubtotalSAR)
               VALUES (?,?,?,?,?,?,?,?)""",
            (pid, sheet, suffix, meta["discount"], f[0], f[1], f[2], meta.get("subtotal_sar")),
        )
        sid = scur.lastrowid
        if primary_discount is None:
            primary_discount = meta["discount"]
            primary_factor = f[0]

        for ln in lines:
            if ln["_type"] == "discount":
                conn.execute(
                    """INSERT INTO Project_BoQ_Lines
                         (ProjectID,SheetID,RowOrder,Description,LineType,TPriceSAR)
                       VALUES (?,?,?,?,?,?)""",
                    (pid, sid, ln["_order"], "Discount", "discount", ln.get("t_price_sar")),
                )
                continue
            iid = upsert_item(conn, ln, path, now)
            conn.execute(
                """INSERT INTO Project_BoQ_Lines
                     (ProjectID,SheetID,ItemID,RowOrder,Area,System,Description,Brand,Model,
                      Qty,ListPriceUSD,ExUnitCostUSD,ShippingPercent,FinalUnitCostUSD,TotalCostUSD,
                      FinalUPriceUSD,TPriceUSD,FinalUPriceSAR,TPriceSAR,MarginExtra,LineType)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (pid, sid, iid, ln["_order"], _s(ln.get("area")), _s(ln.get("system")),
                 ln.get("description"), ln.get("brand"), ln.get("model"), ln.get("qty"),
                 ln.get("list_price"), ln.get("ex_unit_cost"), ln.get("shipping_percent"), ln.get("unit_cost"),
                 ln.get("total_cost"), ln.get("u_price"), ln.get("t_price"),
                 ln.get("u_price_sar"), ln.get("t_price_sar"), ln.get("margin"),
                 ln["_type"]),
            )
            if ln["_type"] == "spare":
                stats["spares"] += 1
            else:
                stats["lines"] += 1
        stats["sheets"] += 1

    conn.execute(
        "UPDATE Projects_Master SET DiscountAmount=?, ConversionFactor=? WHERE ProjectID=?",
        (primary_discount or 0, primary_factor, pid),
    )
    wb.close()
    stats["files"] += 1


def _s(v):
    if v is None:
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    return str(v).strip()


def main(root):
    conn = dbmod.init_db()
    files = []
    for ext in ("*.xlsx", "*.xlsm"):
        files += glob.glob(os.path.join(root, "**", ext), recursive=True)
    files = sorted(f for f in files if "~$" not in f)

    stats = {"files": 0, "sheets": 0, "lines": 0, "spares": 0, "skipped_no_boq": 0,
             "sheets_unparsed": 0, "errors": []}
    print(f"Scanning {len(files)} workbook(s) under:\n  {root}\n")
    for i, f in enumerate(files, 1):
        ingest_file(conn, f, stats)
        conn.commit()
        if i % 10 == 0:
            print(f"  ...{i}/{len(files)}")
    conn.commit()

    print("\n=== INGEST SUMMARY ===")
    print(f"  Files ingested (had BOQ sheets) : {stats['files']}")
    print(f"  BOQ system-sheets parsed        : {stats['sheets']}")
    print(f"  Offer line items stored         : {stats['lines']}")
    print(f"  Spare (parked) items tagged     : {stats['spares']}")
    print(f"  Files skipped (no BOQ sheet)     : {stats['skipped_no_boq']}")
    print(f"  Sheets with no canonical header  : {stats['sheets_unparsed']}")
    cat = conn.execute("SELECT COUNT(*) FROM Items_Catalog").fetchone()[0]
    print(f"  Distinct catalogue items         : {cat}")
    if stats["errors"]:
        print(f"  Errors ({len(stats['errors'])}):")
        for n, e in stats["errors"][:15]:
            print(f"    - {n}: {e}")
    conn.close()


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_ROOT)
