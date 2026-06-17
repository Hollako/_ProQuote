"""
ProQuote â€” Streamlit interface.

Run:  streamlit run app.py   (from the _ProQuote folder)

Workflow 2 (New Offer): type-ahead catalogue search auto-fills Brand/costs/
prices; live-editable grid recalculates Total Cost / T. Price / T. Price SAR
instantly; discount row + bottom totals block with the markup factor.
Workflow 3 (PDF): one click generates the client-facing Quotation PDF.
"""
from __future__ import annotations
import os
import io
import html
import datetime as dt

import pandas as pd
import streamlit as st

import calc
import repo
import pdf_export
import auth
import db

_LOGO = db.banner_path()                       # per-company banner (follows BOQ_DATA_DIR)
_COMPANY = repo.get_setting("company_name") or "SmartWay Systems"
st.set_page_config(page_title=f"ProQuote â€” {_COMPANY}", layout="wide",
                   initial_sidebar_state="expanded")

# Larger button text + icons; tracking status cells are compact and centered.
st.markdown("""<style>
.stButton button { font-size: 1.05rem; font-weight: 600; min-height: 2.9rem; }
.stButton button p { font-size: 1.05rem; }
[class*="st-key-offer_tab_"] button {
  min-height: 4.15rem !important;
  border-radius: 8px 8px 0 0 !important;
  border: 1px solid #cbd5df !important;
  border-bottom-width: 3px !important;
  font-size: 1.35rem !important;
  font-weight: 800 !important;
}
[class*="st-key-offer_tab_"] button p {
  font-size: 1.35rem !important;
  font-weight: 800 !important;
}
[class*="st-key-offer_tab_active_"] button {
  background: #17324d !important;
  border-color: #17324d !important;
  color: #fff !important;
}
[class*="st-key-offer_tab_active_"] button p { color: #fff !important; }
[class*="st-key-offer_tab_locked_"] button {
  background: #f1f4f7 !important;
  color: #8a96a3 !important;
  border-color: #d5dde5 !important;
  opacity: 0.72 !important;
}
[class*="st-key-offer_tab_locked_"] button p { color: #8a96a3 !important; }
[class*="st-key-trkbtn_"] {
  display: flex !important;
  justify-content: center !important;
}
[class*="st-key-trkbtn_"] button {
  min-height: 1.9rem !important;
  height: 1.9rem !important;
  max-height: 1.9rem !important;
  min-width: 2.15rem !important;
  max-width: 2.15rem !important;
  width: 2.15rem !important;
  padding: 0 !important;
  margin: 0 auto !important;
  border-radius: 7px !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
}
[class*="st-key-trkbtn_on_"] button {
  background: #2ea44f !important;
  border-color: #2ea44f !important;
  color: #fff !important;
}
[class*="st-key-trkbtn_partial_"] button {
  background: #f2b84b !important;
  border-color: #f2b84b !important;
  color: #fff !important;
}
[class*="st-key-trkbtn_off_"] button {
  background: #fff !important;
  border-color: #cfd6df !important;
  color: transparent !important;
}
[class*="st-key-trkbtn_"] button p {
  font-size: 1.3rem !important;
  line-height: 1 !important;
  margin: 0 !important;
}
[class*="st-key-trkbtn_on_"] button p { color: #fff !important; }
[class*="st-key-trkbtn_partial_"] button p { color: #fff !important; }
[class*="st-key-trkbtn_off_"] button p { color: transparent !important; }
[class*="st-key-po_"] input,
[class*="st-key-dn_"] input,
[class*="st-key-trkqty_"] input {
  min-height: 2.15rem !important;
  padding-top: 0.25rem !important;
  padding-bottom: 0.25rem !important;
}
[class*="st-key-trkqty_"] input {
  text-align: center !important;
  padding-left: 0.2rem !important;
  padding-right: 0.2rem !important;
}
.tracking-header {
  width: 100%;
  text-align: center;
  font-size: 0.98rem;
  font-weight: 600;
  color: #333;
}
.tracking-stamp {
  width: 100%;
  text-align: center;
  white-space: nowrap;
  font-size: 0.74rem;
  color: #7b8490;
  line-height: 1.2;
  min-height: 0.95rem;
  margin-top: 0.05rem;
}
.tracking-center-cell {
  width: 100%;
  text-align: center;
  line-height: 2.15rem;
}
.tracking-row-separator {
  width: 100%;
  border-top: 1px solid #e3e8ef;
  margin: 0.28rem 0 0.42rem 0;
}
</style>""", unsafe_allow_html=True)

# Full internal grid (builder always sees costs; the client PDF never shows costs).
BUILDER_COLS = ["Area", "System", "Description", "Brand", "Model", "Qty",
                "Cur", "List Price $", "Ex Unit Cost $", "Shipping %", "Unit Cost $", "Total Cost $",
                "Margin x", "U. Price $", "T. Price $", "U. Price SAR", "T. Price SAR"]
# Pure outputs â€” locked in the editor (everything else is an input/driver).
COMPUTED = ["Total Cost $", "T. Price $", "U. Price SAR", "T. Price SAR"]
MONEY_COLS = ["List Price $", "Ex Unit Cost $", "Unit Cost $", "Total Cost $",
              "U. Price $", "T. Price $", "U. Price SAR", "T. Price SAR"]
# Numeric inputs that affect computed prices â€” a change triggers one auto-rerun
# so the recomputed columns refresh immediately (no st.data_editor 1-step lag).
NUM_DRIVERS = ["Qty", "Ex Unit Cost $", "Shipping %", "Unit Cost $", "Margin x", "U. Price $"]
# Reviewing a loaded offer shows selling prices only â€” all cost columns hidden.
PRICE_VIEW_COLS = ["Area", "System", "Description", "Brand", "Model", "Qty",
                   "U. Price $", "T. Price $", "U. Price SAR", "T. Price SAR"]

# Offer terms/notes â€” keys match repo.TERMS_KEYS; defaults from the historical Quotation sheets.
TERMS_KEYS = repo.TERMS_KEYS
PROJECT_SHEET_KEYS = repo.PROJECT_SHEET_KEYS
DEFAULT_TERMS = {
    "subject": "",
    "greeting": ("Dear Sir,\n\nThank you for the opportunity to quote for the above-mentioned "
                 "project. Kindly find hereinafter our offer for your kind review."),
    "system_note": "",
    "scope": "Supply, Installation, Testing & Commissioning.",
    "exclusions": ("Cables, electrical wiring & conduits, back boxes, cabling and/or civil work. "
                   "Pulling cables."),
    "prerequisites": "Power must be available and fully operational prior to the start of our work.",
    "delivery": "8-10 weeks from date of receiving the down payment.",
    "payment": "70% Down payment, 20% upon delivery, 10% upon installation, testing & commissioning.",
    "validity": "30 Days from its date of issuance.",
    "notes": "",
}

PROJECT_LEAD_SOURCE_OPTIONS = [
    "Self Generated", "International Specs", "Hilights", "Lumiere Studio",
    "Follow-Up", "Selection/Alternative",
]
PROJECT_SHIPMENT_OPTIONS = ["Air", "Sea"]
DEFAULT_PROJECT_SHEET_INFO = {
    "job_reference": "",
    "sheet_date": "",
    "lead_source": "Self Generated",
    "commission": "",
    "shipment_by": "Air",
    "downpayment_date": "",
    "invoice_to": "",
    "delivery_instructions": "",
    "salesman_signature": "Sameera Ibrahim",
    "gm_signature": "",
}

# Which user-role each offer people-field is picked from.
PEOPLE_ROLES = {"sales": "sales", "presales": "Pre-Sales", "pm": "Project Manager"}


def _person_select(col, label, role, current, key):
    """Dropdown of active users holding `role`; keeps any legacy stored value selectable."""
    names = auth.users_in_role(role)
    cur = (current or "").strip()
    opts = ["â€”"] + names
    if cur and cur not in names:
        opts = ["â€”", cur] + names           # preserve a name that isn't a current user
    pick = col.selectbox(label, opts, index=opts.index(cur) if cur in opts else 0, key=key)
    return "" if pick == "â€”" else pick


def _empty_grid() -> pd.DataFrame:
    df = pd.DataFrame([calc.blank_row()])
    df["LineType"] = "item"
    df["_ItemID"] = None
    return df.iloc[0:0]


def _ensure_state():
    if "grid" not in st.session_state:
        st.session_state.grid = _empty_grid()
    if "header" not in st.session_state:
        st.session_state.header = {
            **DEFAULT_TERMS,
            "client": "", "project": "", "contact": "", "phone": "",
            "sales": "", "presales": "", "pm": "",
            "offer": _next_offer_no(), "system": "LCS",
            "date": dt.date.today().isoformat(), "margin": 1.60,
            "project_sheet": dict(DEFAULT_PROJECT_SHEET_INFO),
        }
    if "discount" not in st.session_state:
        st.session_state.discount = 0.0


def _new_offer_header(overrides: dict | None = None) -> dict:
    header = {
        **DEFAULT_TERMS,
        "client": "", "project": "", "contact": "", "phone": "",
        "sales": "", "presales": "", "pm": "",
        "offer": _next_offer_no(), "system": "LCS",
        "date": dt.date.today().isoformat(), "margin": 1.60,
        "project_sheet": dict(DEFAULT_PROJECT_SHEET_INFO),
    }
    if overrides:
        header.update(overrides)
    return header


def _prime_new_offer_form(header: dict | None = None, grid: pd.DataFrame | None = None,
                          discount: float = 0.0):
    """Load data into the New Offer form before its widgets are rendered."""
    h = _new_offer_header(header or {})
    h["project_sheet"] = {**DEFAULT_PROJECT_SHEET_INFO, **(h.get("project_sheet") or {})}
    st.session_state.header = h
    st.session_state.grid = grid.copy() if grid is not None else _empty_grid()
    st.session_state.discount = abs(float(discount or 0.0))
    st.session_state.no_offer_lock = None
    st.session_state.no_saved_options = []
    for key in ("editor", "pdf_bytes", "project_sheet_bytes", "saved_rev"):
        st.session_state.pop(key, None)
    for key in ("no_discount_percent", "no_discount_driver", "no_discount_subtotal"):
        st.session_state.pop(key, None)

    st.session_state["no_client"] = h.get("client", "")
    st.session_state["no_project"] = h.get("project", "")
    st.session_state["no_contact"] = h.get("contact", "")
    st.session_state["no_phone"] = h.get("phone", "")
    st.session_state["no_sales"] = h.get("sales", "")
    st.session_state["no_presales"] = h.get("presales", "")
    st.session_state["no_pm"] = h.get("pm", "")
    st.session_state["no_offer_ov"] = ""
    st.session_state["no_option"] = ""

    offer_types = repo.offer_types()
    system = h.get("system") or ""
    st.session_state["no_offer_type"] = system if system in offer_types else "(none)"

    term_keys = {
        "subject": "no_subject", "greeting": "no_greet", "system_note": "no_sys",
        "scope": "no_scope", "exclusions": "no_excl", "prerequisites": "no_prereq",
        "delivery": "no_deliv", "validity": "no_valid", "payment": "no_pay",
        "notes": "no_notes",
    }
    for src, key in term_keys.items():
        st.session_state[key] = h.get(src, DEFAULT_TERMS.get(src, ""))

    project_sheet_keys = {
        "job_reference": "no_ps_job_reference",
        "sheet_date": "no_ps_sheet_date",
        "lead_source": "no_ps_lead_source",
        "commission": "no_ps_commission",
        "shipment_by": "no_ps_shipment_by",
        "downpayment_date": "no_ps_downpayment_date",
        "invoice_to": "no_ps_invoice_to",
        "delivery_instructions": "no_ps_delivery_instructions",
        "salesman_signature": "no_ps_salesman_signature",
        "gm_signature": "no_ps_gm_signature",
    }
    ps_info = h.get("project_sheet") or {}
    for src, key in project_sheet_keys.items():
        st.session_state[key] = ps_info.get(src, DEFAULT_PROJECT_SHEET_INFO.get(src, ""))


def _next_offer_no() -> str:
    yr = dt.date.today().strftime("%y")
    return f"OFR-SWS-RUH-{yr}-NEW"


def _add_row_to(state_key: str, row: dict):
    g = st.session_state[state_key]
    st.session_state[state_key] = pd.concat([g, pd.DataFrame([row])], ignore_index=True)


def _add_row(row: dict):
    _add_row_to("grid", row)


def _text(value, default: str = "") -> str:
    """Display-safe text for DB/pandas values, treating None/NaN as empty."""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return text if text else default


def render_editable_grid(state_key: str, editor_key: str):
    """Full editable grid (all columns) with live recompute + one auto-rerun on change."""
    grid = calc.recompute(st.session_state[state_key])
    colcfg = {c: st.column_config.NumberColumn(c, format="%.0f") for c in MONEY_COLS}
    colcfg["Qty"] = st.column_config.NumberColumn("Qty", format="%d", min_value=0)
    colcfg["Cur"] = st.column_config.SelectboxColumn(
        "Cur", options=calc.CURRENCIES, required=False, width="small",
        help="Currency of List Price & Ex Unit Cost. The Unit Cost is converted to "
             "USD automatically (EUR rate from Settings; SAR pegged at 3.75).")
    # List Price & Ex Unit Cost are in the row's chosen currency (drop the misleading $).
    colcfg["List Price $"] = st.column_config.NumberColumn("List Price", format="%.0f")
    colcfg["Ex Unit Cost $"] = st.column_config.NumberColumn("Ex Unit Cost", format="%.0f")
    colcfg["Unit Cost $"] = st.column_config.NumberColumn("Unit Cost (USD)", format="%.0f")
    colcfg["Shipping %"] = st.column_config.NumberColumn(
        "Shipping %", format="%.2f", min_value=0.0, step=5.0,
        help="Added to Ex Unit Cost. Unit Cost = Ex Unit Cost x (1 + Shipping % / 100), in USD.")
    colcfg["Margin x"] = st.column_config.NumberColumn(
        "Margin Ã-", format="%.2f", min_value=0.0, step=0.05,
        help="Multiplier on landed Unit Cost. U.Price $ = âŒˆUnit Cost Ã- MarginâŒ‰. "
             "Set 0 to type U.Price $ manually.")
    edited = st.data_editor(
        grid[BUILDER_COLS] if not grid.empty else grid,
        column_config=colcfg, disabled=[c for c in COMPUTED if c in BUILDER_COLS],
        num_rows="dynamic", use_container_width=True, key=editor_key, hide_index=True,
    ).reset_index(drop=True)

    base = st.session_state[state_key].reset_index(drop=True)
    new_grid = edited.copy()
    n = len(new_grid)
    new_grid["LineType"] = ["discount" if str(d).strip().lower() == "discount" else "item"
                            for d in new_grid.get("Description", pd.Series([""] * n))]
    new_grid["_ItemID"] = [base["_ItemID"].iloc[i] if (i < len(base) and "_ItemID" in base.columns)
                           else None for i in range(n)]
    new_grid = calc.recompute(new_grid)
    prev = calc.recompute(base)
    drivers = [c for c in NUM_DRIVERS if c in new_grid.columns and c in prev.columns]
    same_len = len(new_grid) == len(prev)
    cur_changed = same_len and "Cur" in new_grid.columns and "Cur" in prev.columns and not (
        new_grid["Cur"].astype(str).reset_index(drop=True)
        .equals(prev["Cur"].astype(str).reset_index(drop=True)))
    changed = (not same_len) or cur_changed or not (
        new_grid[drivers].fillna(0).round(4).reset_index(drop=True)
        .equals(prev[drivers].fillna(0).round(4).reset_index(drop=True)))
    st.session_state[state_key] = new_grid
    if changed:
        st.rerun()
    _fx_hint(new_grid)                     # show conversions only on the stable pass
    return new_grid


def _fx_hint(grid):
    """Small caption translating any non-USD Ex Unit Cost into USD for the visible lines."""
    if grid is None or grid.empty or "Cur" not in grid.columns:
        return
    parts = []
    for _, r in grid.iterrows():
        cur = str(r.get("Cur") or "USD")
        ex = calc._num(r.get("Ex Unit Cost $"))
        if cur != "USD" and ex > 0:
            desc = (str(r.get("Description") or "").strip() or "item")[:28]
            parts.append(f"{desc}: {ex:,.0f} {cur} â†’ ${calc.to_usd(ex, cur):,.2f}")
    if parts:
        st.caption("ðŸ’± Ex cost â†’ USD (then Ã-(1+Shipping%) = Unit Cost):  "
                   + "   Â·   ".join(parts))


def catalogue_add(state_key: str, default_margin: float, kp: str, default_system: str = "",
                  show_clear: bool = False):
    """Type-ahead catalogue search + add controls writing into st.session_state[state_key]."""
    st.markdown("##### Add item from catalogue")
    term = st.text_input("Search Model / Description / Brand", key=f"{kp}_term",
                         placeholder="e.g. PDEG, keypad, Dynaliteâ€¦")
    results = repo.search_catalog(term, limit=20)
    if not results.empty:
        results = results.assign(_label=results.apply(
            lambda r: f"{r['Model']} â€” {str(r['Description'])[:48]} ({r['Brand']})  Â·x{r['TimesQuoted']}", axis=1))
        a1, a2, a3, a4, a5 = st.columns([4, 1, 1.4, 1.4, 1.3], vertical_alignment="bottom")
        pick = a1.selectbox("Match", results["_label"].tolist(), key=f"{kp}_pick")
        chosen = results[results["_label"] == pick].iloc[0].to_dict()
        qty = a2.number_input("Qty", min_value=1, value=1, step=1, key=f"{kp}_qty")
        area = a3.text_input("Area", value=default_system, key=f"{kp}_area")
        system = a4.text_input("System", value=default_system, key=f"{kp}_system")
        if a5.button("âž• Add", use_container_width=True, key=f"{kp}_add"):
            _add_row_to(state_key, repo.item_to_grid_row(
                chosen, area=area, system=system, qty=int(qty), default_margin=default_margin))
            st.rerun()
    elif term:
        st.info("No catalogue match â€” add a blank row below and type freely.")
    bc1, bc2, _ = st.columns([1, 1, 4])
    if bc1.button("âž• Blank row", key=f"{kp}_blank", use_container_width=True):
        _add_row_to(state_key, {**calc.blank_row(system=default_system),
                                "Margin x": default_margin, "LineType": "item", "_ItemID": None})
        st.rerun()
    if show_clear and bc2.button("ðŸ§¹ Clear grid", key=f"{kp}_clear", use_container_width=True):
        st.session_state[state_key] = _empty_grid()
        st.rerun()


def terms_form(store: dict, kp: str):
    """Editable Quotation terms/notes (subject, greeting, scope, payment, ...)."""
    with st.expander("ðŸ“‹ Terms, scope & notes (appear on the quotation)", expanded=False):
        store["subject"] = st.text_input("Subject (offer title)", store.get("subject", ""),
            key=f"{kp}_subject", placeholder="e.g. Low Current Systems Offer")
        store["greeting"] = st.text_area("Greeting", store.get("greeting", ""),
            key=f"{kp}_greet", height=80)
        c1, c2 = st.columns(2)
        store["system_note"] = c1.text_input("System", store.get("system_note", ""),
            key=f"{kp}_sys", placeholder="e.g. Smart System as per the above detailed BOQ.")
        store["scope"] = c2.text_input("Scope", store.get("scope", ""), key=f"{kp}_scope")
        store["exclusions"] = st.text_area("Exclusions", store.get("exclusions", ""),
            key=f"{kp}_excl", height=70)
        store["prerequisites"] = st.text_area("Pre-requirements", store.get("prerequisites", ""),
            key=f"{kp}_prereq", height=70)
        c3, c4 = st.columns(2)
        store["delivery"] = c3.text_input("Delivery", store.get("delivery", ""), key=f"{kp}_deliv")
        store["validity"] = c4.text_input("Validity", store.get("validity", ""), key=f"{kp}_valid")
        store["payment"] = st.text_area("Payment Terms", store.get("payment", ""),
            key=f"{kp}_pay", height=70)
        store["notes"] = st.text_area("Special notes & instructions", store.get("notes", ""),
            key=f"{kp}_notes", height=70)


def _project_sheet_info(info: dict | None = None, h: dict | None = None) -> dict:
    h = h or {}
    data = {**DEFAULT_PROJECT_SHEET_INFO, **(info or {})}
    if not data.get("sheet_date"):
        data["sheet_date"] = h.get("date") or dt.date.today().isoformat()
    if not data.get("invoice_to"):
        data["invoice_to"] = h.get("client") or ""
    if not data.get("delivery_instructions"):
        contact = " ".join(x for x in [h.get("contact"), h.get("phone")] if x)
        data["delivery_instructions"] = contact
    if data.get("lead_source") not in PROJECT_LEAD_SOURCE_OPTIONS:
        data["lead_source"] = data.get("lead_source") or PROJECT_LEAD_SOURCE_OPTIONS[0]
    if data.get("shipment_by") not in PROJECT_SHIPMENT_OPTIONS:
        data["shipment_by"] = data.get("shipment_by") or PROJECT_SHIPMENT_OPTIONS[0]
    return {k: data.get(k, "") for k in PROJECT_SHEET_KEYS}


def _option_index(options: list[str], value: str) -> tuple[list[str], int]:
    opts = list(options)
    value = _text(value)
    if value and value not in opts:
        opts = [value] + opts
    return opts, opts.index(value) if value in opts else 0


def project_sheet_info_form(store: dict, kp: str):
    """Editable Project Sheet export details."""
    ps = _project_sheet_info(store.get("project_sheet"), store)
    with st.expander("Project Sheet Information", expanded=False):
        c1, c2, c3 = st.columns(3)
        ps["job_reference"] = c1.text_input(
            "Project Job Reference", ps.get("job_reference", ""),
            key=f"{kp}_job_reference")
        ps["sheet_date"] = c2.text_input(
            "Project Sheet Date", ps.get("sheet_date", ""),
            key=f"{kp}_sheet_date")
        lead_opts, lead_idx = _option_index(PROJECT_LEAD_SOURCE_OPTIONS, ps.get("lead_source", ""))
        ps["lead_source"] = c3.selectbox(
            "Project Lead Source", lead_opts, index=lead_idx,
            key=f"{kp}_lead_source")

        c4, c5, c6 = st.columns(3)
        ps["commission"] = c4.text_input(
            "Architect/Contractor Commissions (if any)", ps.get("commission", ""),
            key=f"{kp}_commission")
        ship_opts, ship_idx = _option_index(PROJECT_SHIPMENT_OPTIONS, ps.get("shipment_by", ""))
        ps["shipment_by"] = c5.selectbox(
            "Based on Shipments by", ship_opts, index=ship_idx,
            key=f"{kp}_shipment_by")
        ps["downpayment_date"] = c6.text_input(
            "Downpayment Date", ps.get("downpayment_date", ""),
            key=f"{kp}_downpayment_date")

        ps["invoice_to"] = st.text_input("Invoice to", ps.get("invoice_to", ""),
                                         key=f"{kp}_invoice_to")
        ps["delivery_instructions"] = st.text_area(
            "Delivery Instructions / Contact person & details",
            ps.get("delivery_instructions", ""), key=f"{kp}_delivery_instructions", height=70)
        c9, c10 = st.columns(2)
        ps["salesman_signature"] = c9.text_input(
            "Salesman Signature Name", ps.get("salesman_signature", ""),
            key=f"{kp}_salesman_signature")
        ps["gm_signature"] = c10.text_input(
            "GM Signature Name", ps.get("gm_signature", ""),
            key=f"{kp}_gm_signature")

    store["project_sheet"] = ps
    return ps


def _make_pdf_download(h, grid, summary, options=None):
    notes = {
        "System": h.get("system_note"), "Scope": h.get("scope"),
        "Exclusions": h.get("exclusions"), "Pre-requirements": h.get("prerequisites"),
        "Delivery": h.get("delivery"), "Payment Terms": h.get("payment"),
        "Validity": h.get("validity"), "Notes": h.get("notes"),
    }
    header = {"title": h.get("subject") or "Quotation",
              "client": h.get("client"), "project": h.get("project"),
              "contact": h.get("contact"), "phone": h.get("phone"),
              "sales": h.get("sales"), "presales": h.get("presales"), "pm": h.get("pm"),
              "offer": h.get("offer"), "date": h.get("date"),
              "greeting": h.get("greeting") or DEFAULT_TERMS["greeting"]}
    company = {
        "name": repo.get_setting("company_name") or "SmartWay Systems",
        "tagline": repo.get_setting("company_tagline") or "",
        "contact": repo.get_setting("company_contact") or "",
        "color": repo.get_setting("company_brand_color") or "#002060",
    }
    tmp = os.path.join(db.DATA_DIR, "_last_quotation.pdf")
    if options:                       # one document, a section per option
        pdf_export.generate_options_pdf(tmp, header, options, notes=notes,
                                        company=company, show_costs=False)
    else:
        pdf_export.generate_quotation_pdf(tmp, header, grid, summary, notes=notes,
                                          company=company, show_costs=False)
    with open(tmp, "rb") as f:
        st.session_state.pdf_bytes = f.read()
    n = len(options) if options else 1
    st.toast(f"PDF ready ({n} option{'s' if n > 1 else ''}) â€” use the download button.", icon="ðŸ“„")


def _project_sheet_bytes(h: dict, s: dict) -> bytes:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "PROJECT SHEET"
    ws.sheet_view.showGridLines = False

    ps = _project_sheet_info(h.get("project_sheet"), h)
    net_sales = s.get("discounted_subtotal_sar") or s.get("subtotal_sar") or 0.0
    landed_cost = s.get("cost_sar") or 0.0

    def _display_date(value):
        raw = _text(value)
        if not raw:
            return dt.date.today().strftime("%B %d,%Y")
        try:
            return dt.date.fromisoformat(raw[:10]).strftime("%B %d,%Y")
        except ValueError:
            return raw

    def _shipment_text(value):
        ship = (_text(value) or "Air").lower()
        air = "X" if ship == "air" else " "
        sea = "X" if ship == "sea" else " "
        return f"(  {air}  ) Air                   (  {sea}  ) Sea"

    header_fill = PatternFill("solid", fgColor="FFFFFF")
    fallback_header_fill = PatternFill("solid", fgColor="002060")
    sidebar_fill = PatternFill("solid", fgColor="002060")
    side = Side(style="thin", color="1F1F1F")
    border = Border(left=side, right=side, top=side, bottom=side)
    label_font = Font(name="Calibri", size=11, bold=True)
    value_font = Font(name="Calibri", size=11)

    ws.merge_cells("A1:E4")
    has_banner = os.path.exists(_LOGO)
    for row in range(1, 5):
        ws.row_dimensions[row].height = 20
        for col in range(1, 6):
            ws.cell(row, col).fill = header_fill if has_banner else fallback_header_fill
    if has_banner:
        banner = XLImage(_LOGO)
        banner.width = 695
        banner.height = 77
        ws.add_image(banner, "A1")
    else:
        ws["A1"] = repo.get_setting("company_name") or "SmartWay Systems"
        ws["A1"].font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[5].height = 10

    ws.merge_cells("A6:A24")
    ws["A6"] = "Project Sheet"
    ws["A6"].font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
    ws["A6"].fill = sidebar_fill
    ws["A6"].alignment = Alignment(horizontal="center", vertical="center",
                                   text_rotation=90, wrap_text=True)
    ws["A6"].border = border

    rows = [
        (6, "Project Job Reference:", ps.get("job_reference"), False),
        (7, "Date:", _display_date(ps.get("sheet_date")), False),
        (8, "Project Name:", h.get("project") or h.get("subject") or "", False),
        (9, "Client/Contractor Name:", h.get("client") or "", False),
        (10, "Confirmed Offer Reference:", h.get("offer") or "", False),
        (11, "Salesman:", h.get("sales") or "", False),
        (12, "Project Lead Source:", ps.get("lead_source"), False),
        (13, "Net Projects Sales Amount without VAT:", net_sales, True),
        (14, "Architect/Contractor Commissions (If any):", ps.get("commission"), False),
        (15, "Projected Total Project Landed Cost:", landed_cost, True),
        (16, "Projected Margin:", "=IF(D13>0,(D13-D15)/D13,0)", False, "percent"),
        (17, "Projected Profit:", "=D13-D15", True),
        (18, "Based on Shipments by:", _shipment_text(ps.get("shipment_by")), False),
        (19, "Payment Terrms:", h.get("payment") or "", False),
        (20, "Downpayment Date:", ps.get("downpayment_date"), False),
        (21, "Invoice to:", ps.get("invoice_to"), False),
        (22, "Contractual Project Delivery Date:", h.get("delivery") or "", False),
        (23, "Delivery Instructions / Contact person & details:", ps.get("delivery_instructions"), False),
        (24, "Notes: ", h.get("notes") or "", False),
    ]
    for item in rows:
        row, label, value, money_row = item[:4]
        value_kind = item[4] if len(item) > 4 else ""
        ws.cell(row, 2, label)
        if money_row:
            ws.cell(row, 3, "SAR")
            ws.cell(row, 4, value)
            ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
            ws.cell(row, 3).alignment = Alignment(horizontal="center", vertical="center")
        else:
            ws.cell(row, 3, value)
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
        ws.cell(row, 2).font = label_font
        ws.cell(row, 3).font = value_font
        ws.cell(row, 4).font = value_font
        ws.cell(row, 2).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row, 3).alignment = Alignment(horizontal="center",
                                              vertical="center", wrap_text=True)
        ws.cell(row, 4).alignment = Alignment(horizontal="right" if money_row else "center",
                                              vertical="center", wrap_text=True)
        if value_kind == "percent":
            ws.cell(row, 3).number_format = "0%"
        for col in range(2, 6):
            ws.cell(row, col).border = border

    ws["D13"].number_format = "#,##0.00"
    ws["D15"].number_format = "#,##0.00"
    ws["C16"].number_format = "0%"
    ws["D17"].number_format = "#,##0.00"

    ws.merge_cells("A25:B27")
    ws["A25"] = "Salesman Signature\n\n" + (ps.get("salesman_signature") or "")
    ws.merge_cells("C25:E27")
    ws["C25"] = "GM Signature" + (("\n\n" + ps.get("gm_signature")) if ps.get("gm_signature") else "")
    for cell in (ws["A25"], ws["C25"]):
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for row in range(25, 28):
        for col in range(1, 6):
            ws.cell(row, col).border = border

    for row in range(6, 25):
        ws.row_dimensions[row].height = 27
    ws.row_dimensions[13].height = 31
    ws.row_dimensions[14].height = 31
    ws.row_dimensions[15].height = 31
    ws.row_dimensions[17].height = 31
    ws.row_dimensions[19].height = 35
    ws.row_dimensions[23].height = 33
    ws.row_dimensions[24].height = 33
    ws.row_dimensions[25].height = 28
    ws.row_dimensions[26].height = 28
    ws.row_dimensions[27].height = 28
    for col, width in {
        "A": 9, "B": 39, "C": 9, "D": 22, "E": 20,
    }.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A6"
    ws.print_area = "A1:E27"
    ws.print_title_rows = "1:5"

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _make_project_sheet_download(h: dict, summary: dict):
    st.session_state.project_sheet_bytes = _project_sheet_bytes(h, summary)
    st.toast("Project Sheet ready - use the download button.", icon="ðŸ“Š")


def _safe_filename(value, fallback="Project"):
    name = _text(value, fallback)
    return "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in name).strip("_") or fallback


def revision_options(base_pid):
    """All active options of the same revision as `base_pid`, as PDF sections."""
    meta = repo.project_meta(base_pid)
    fam = repo.family_key(meta.get("OfferNo"), meta.get("ProjectName"))
    rev = int(meta.get("RevisionNo") or 0)
    projs = repo.list_projects()
    same = projs[projs.apply(
        lambda r: repo.family_key(r.get("OfferNo"), r.get("ProjectName")) == fam
        and int(r.get("RevisionNo") or 0) == rev, axis=1)]
    active = same[same["Archived"].fillna(0) == 0]
    rows = active if not active.empty else same[same["ProjectID"] == base_pid]
    rows = rows.sort_values("OptionLabel", na_position="first")
    out = []
    for _, r in rows.iterrows():
        pid = int(r["ProjectID"])
        m = repo.project_meta(pid)
        sh = (repo.list_systems(pid) or [None])[0]
        g = repo.load_project_grid(pid, sh).copy()
        for col in MONEY_COLS:
            if col in g.columns:
                g[col] = g[col].map(lambda v: calc.roundup(v, 0))
        out.append({"label": m.get("OptionLabel") or "",
                    "grid": g, "summary": calc.summarize(g, m.get("DiscountAmount") or 0)})
    return out


def _profit_banner(s: dict):
    """Profit bubble: profit big (left) with Margin/markup beneath, cost big (right)."""
    profit = s.get("gross_margin_sar") or 0.0
    profit_usd = s.get("gross_margin_usd") or 0.0
    sub = s.get("discounted_subtotal_sar") or 0.0
    margin_pct = (profit / sub * 100) if sub else 0.0
    factor = s.get("markup_factor")
    cost_sar = s.get("cost_sar", 0) or 0.0
    cost_usd = s.get("total_cost_usd", 0) or 0.0
    markup_txt = f"Markup Ã-{factor:.2f}" if factor else "Markup â€”"
    if profit >= 0:
        bg, fg, sub_fg = "rgba(33,195,84,0.12)", "#0b6b34", "#3f7d59"
    else:
        bg, fg, sub_fg = "rgba(255,43,43,0.10)", "#b02a37", "#a05560"
    lbl = f"font-size:0.8rem;font-weight:700;color:{sub_fg};text-transform:uppercase;letter-spacing:0.04em"
    big = f"font-size:1.5rem;font-weight:800;color:{fg};line-height:1.2"
    small = f"font-size:0.9rem;font-weight:700;color:{sub_fg}"
    mid = f"font-size:1.0rem;font-weight:800;color:{fg}"

    def _block(label, sar, usd, align):
        return (f"<div style='flex:1;min-width:130px;text-align:{align}'>"
                f"<div style='{lbl}'>{label}</div>"
                f"<div style='{big}'>SAR {sar:,.0f}</div>"
                f"<div style='{small}'>$ {usd:,.0f}</div></div>")

    html = (
        f"<div style='background:{bg};border-radius:8px;padding:14px 24px;margin:2px 0 10px;"
        f"display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap'>"
        + _block("ðŸ§¾ Cost", cost_sar, cost_usd, "left")
        + f"<div style='flex:1;min-width:120px;text-align:center'>"
          f"<div style='{mid}'>{markup_txt}</div>"
          f"<div style='{mid}'>Margin {margin_pct:.1f}%</div></div>"
        + _block("ðŸ’° Gross Profit", profit, profit_usd, "right")
        + "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def _fin_bubble(left_label, left_val, middle_lines, right_label, right_val, positive=True):
    """Green (or red) summary bubble used under each Finance table."""
    bg, fg, sub_fg = (("rgba(33,195,84,0.12)", "#0b6b34", "#3f7d59") if positive
                      else ("rgba(255,43,43,0.10)", "#b02a37", "#a05560"))
    lbl = f"font-size:0.78rem;font-weight:700;color:{sub_fg};text-transform:uppercase;letter-spacing:0.04em"
    big = f"font-size:1.3rem;font-weight:800;color:{fg};line-height:1.25"
    mid = f"font-size:0.92rem;font-weight:800;color:{fg}"
    mid_html = "".join(f"<div style='{mid}'>{m}</div>" for m in middle_lines)
    st.markdown(
        f"<div style='background:{bg};border-radius:8px;padding:12px 20px;margin:8px 0 2px;"
        f"display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap'>"
        f"<div style='flex:1;min-width:120px;text-align:left'><div style='{lbl}'>{left_label}</div>"
        f"<div style='{big}'>{left_val}</div></div>"
        f"<div style='flex:1;min-width:110px;text-align:center'>{mid_html}</div>"
        f"<div style='flex:1;min-width:120px;text-align:right'><div style='{lbl}'>{right_label}</div>"
        f"<div style='{big}'>{right_val}</div></div></div>", unsafe_allow_html=True)


def _subtotal_metric(col, s: dict):
    discount = s.get("discount_sar") or 0.0
    if discount:
        col.metric("Subtotal after discount (SAR)",
                   f"{s['discounted_subtotal_sar']:,.0f}",
                   delta=f"-{discount:,.0f} discount", delta_color="inverse")
    else:
        col.metric("Subtotal (SAR)", f"{s['subtotal_sar']:,.0f}")


def _summary_metrics(s: dict):
    m1, m2, m3 = st.columns(3)
    _subtotal_metric(m1, s)
    m2.metric(f"VAT {calc.VAT_RATE * 100:g}% (SAR)", f"{s['vat_amount_sar']:,.0f}")
    m3.metric("Grand Total (SAR)", f"{s['grand_total_sar']:,.0f}")


def _safe_float(value, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        if pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _discount_percent(amount: float, subtotal: float) -> float:
    return round((amount / subtotal * 100), 4) if subtotal else 0.0


def _sync_discount_from_amount(amount_key: str, percent_key: str,
                               subtotal_key: str, driver_key: str):
    subtotal = max(_safe_float(st.session_state.get(subtotal_key)), 0.0)
    amount = min(abs(_safe_float(st.session_state.get(amount_key))), subtotal)
    st.session_state[amount_key] = amount
    st.session_state[percent_key] = _discount_percent(amount, subtotal)
    st.session_state[driver_key] = "amount"


def _sync_discount_from_percent(amount_key: str, percent_key: str,
                                subtotal_key: str, driver_key: str):
    subtotal = max(_safe_float(st.session_state.get(subtotal_key)), 0.0)
    percent = min(abs(_safe_float(st.session_state.get(percent_key))), 100.0)
    st.session_state[percent_key] = percent
    st.session_state[amount_key] = round(subtotal * percent / 100, 2)
    st.session_state[driver_key] = "percent"


def _discount_inputs(prefix: str, amount_key: str, subtotal: float,
                     amount_col=None, percent_col=None) -> float:
    subtotal = max(_safe_float(subtotal), 0.0)
    percent_key = f"{prefix}_discount_percent"
    subtotal_key = f"{prefix}_discount_subtotal"
    driver_key = f"{prefix}_discount_driver"
    st.session_state[subtotal_key] = subtotal

    amount = min(abs(_safe_float(st.session_state.get(amount_key))), subtotal)
    st.session_state[amount_key] = amount
    if st.session_state.get(driver_key) not in ("amount", "percent"):
        st.session_state[driver_key] = "amount"

    if st.session_state[driver_key] == "percent":
        percent = min(abs(_safe_float(st.session_state.get(percent_key))), 100.0)
        st.session_state[percent_key] = percent
        st.session_state[amount_key] = round(subtotal * percent / 100, 2)
    else:
        st.session_state[percent_key] = _discount_percent(st.session_state[amount_key], subtotal)

    if amount_col is None or percent_col is None:
        amount_col, percent_col = st.columns(2)
    if st.session_state[amount_key] > subtotal:
        st.session_state[amount_key] = subtotal
        st.session_state[percent_key] = _discount_percent(subtotal, subtotal)

    amount_col.number_input(
        "Discount (SAR)", min_value=0.0, max_value=subtotal, step=100.0, format="%.2f", key=amount_key,
        help="Fixed discount amount to subtract from the subtotal.",
        on_change=_sync_discount_from_amount,
        args=(amount_key, percent_key, subtotal_key, driver_key))
    percent_col.number_input(
        "Discount %", min_value=0.0, max_value=100.0, step=1.0, format="%.2f", key=percent_key,
        help="Percentage of the current subtotal. Editing this updates Discount (SAR).",
        on_change=_sync_discount_from_percent,
        args=(amount_key, percent_key, subtotal_key, driver_key))
    return abs(_safe_float(st.session_state.get(amount_key)))


def _set_offer_active_tab(state_key: str, tab: str):
    st.session_state[state_key] = tab


def _offer_tab_selector(project_id: int, approved: bool) -> str:
    tabs = ("BoQ", "Tracking", "Finance")
    state_key = f"offer_active_tab_{project_id}"
    active = st.session_state.get(state_key, "BoQ")
    if active not in tabs or (not approved and active != "BoQ"):
        active = "BoQ"
        st.session_state[state_key] = active

    tab_css = ["<style>"]
    for label in tabs:
        slug = label.lower()
        selector = f'[class*="st-key-offer_tab_{slug}_{project_id}"] button'
        text_selector = f'[class*="st-key-offer_tab_{slug}_{project_id}"] button p'
        locked = label != "BoQ" and not approved
        if label == active:
            tab_css.append(
                f"{selector} {{ background: #17324d !important; border-color: #17324d !important; "
                "color: #fff !important; }}"
            )
            tab_css.append(f"{text_selector} {{ color: #fff !important; }}")
        elif locked:
            tab_css.append(
                f"{selector} {{ background: #f1f4f7 !important; color: #8a96a3 !important; "
                "border-color: #d5dde5 !important; opacity: 0.72 !important; }}"
            )
            tab_css.append(f"{text_selector} {{ color: #8a96a3 !important; }}")
    tab_css.append("</style>")
    st.markdown("\n".join(tab_css), unsafe_allow_html=True)

    cols = st.columns(3, gap="small")
    for label, col in zip(tabs, cols):
        locked = label != "BoQ" and not approved
        key = f"offer_tab_{label.lower()}_{project_id}"
        col.button(label, key=key, disabled=locked, use_container_width=True,
                   type="primary" if active == label else "secondary",
                   on_click=_set_offer_active_tab, args=(state_key, label))

    if not approved:
        st.caption("Tracking and Finance are enabled after the offer is approved.")
    return active


def _fmt_tracking_stamp(value) -> str:
    raw = _text(value)
    if not raw:
        return ""
    try:
        return dt.datetime.fromisoformat(raw).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return raw.replace("T", " ")[:16]


def _tracking_center_text(col, value):
    text = html.escape(_text(value)) or "&nbsp;"
    col.markdown(f"<div class='tracking-center-cell'>{text}</div>", unsafe_allow_html=True)


def _tracking_keys(lid: int, key_name: str):
    value_key = f"trk_status_{key_name}_{lid}"
    stamp_key = f"trk_{key_name}_stamp_{lid}"
    qty_key = f"trkqty_{key_name}_{lid}"
    return value_key, stamp_key, qty_key


def _bounded_tracking_qty(value, max_qty: float) -> float:
    qty = max(_safe_float(value), 0.0)
    if max_qty > 0:
        qty = min(qty, max_qty)
    return round(qty, 4)


def _set_tracking_qty(value_key: str, stamp_key: str, qty_key: str,
                      qty, max_qty: float):
    qty = _bounded_tracking_qty(qty, max_qty)
    st.session_state[qty_key] = qty
    checked = qty > 0
    st.session_state[value_key] = checked
    if checked and not st.session_state.get(stamp_key):
        st.session_state[stamp_key] = dt.datetime.now().isoformat(timespec="minutes")
    elif not checked:
        st.session_state[stamp_key] = ""


def _open_tracking_qty_prompt(value_key: str, stamp_key: str, qty_key: str,
                              action: str, description: str, full_qty: float):
    st.session_state.tracking_qty_prompt = {
        "value_key": value_key,
        "stamp_key": stamp_key,
        "qty_key": qty_key,
        "action": action,
        "description": description,
        "full_qty": _bounded_tracking_qty(full_qty, full_qty),
    }
    st.session_state["tracking_prompt_qty"] = _bounded_tracking_qty(full_qty, full_qty)


def _clear_tracking_qty_prompt():
    st.session_state.pop("tracking_qty_prompt", None)
    st.session_state.pop("tracking_prompt_qty", None)


def _render_tracking_qty_prompt():
    prompt = st.session_state.get("tracking_qty_prompt")
    if not prompt:
        return
    action = _text(prompt.get("action")).title()
    desc = _text(prompt.get("description"), "this item")
    full_qty = _bounded_tracking_qty(prompt.get("full_qty"), prompt.get("full_qty") or 0.0)

    def prompt_body():
        st.write(f"**{desc}**")
        st.write(
            f"Is the total quantity ({full_qty:g}) {action.lower()}, "
            "or only part of it?"
        )
        if st.button(f"Full quantity ({full_qty:g})", key="tracking_prompt_full",
                     type="primary", use_container_width=True):
            _set_tracking_qty(prompt["value_key"], prompt["stamp_key"], prompt["qty_key"],
                              full_qty, full_qty)
            _clear_tracking_qty_prompt()
            st.rerun()
        partial_qty = st.number_input(
            "Partial quantity",
            min_value=0.0,
            max_value=full_qty,
            step=1.0,
            format="%.2f",
            key="tracking_prompt_qty",
        )
        b1, b2 = st.columns(2)
        if b1.button(f"Use partial", key="tracking_prompt_partial",
                     use_container_width=True):
            _set_tracking_qty(prompt["value_key"], prompt["stamp_key"], prompt["qty_key"],
                              partial_qty, full_qty)
            _clear_tracking_qty_prompt()
            st.rerun()
        if b2.button("Cancel", key="tracking_prompt_cancel", use_container_width=True):
            _clear_tracking_qty_prompt()
            st.rerun()

    if hasattr(st, "dialog"):
        @st.dialog(f"{action} quantity")
        def tracking_qty_dialog():
            prompt_body()
        tracking_qty_dialog()
    else:
        with st.container(border=True):
            prompt_body()


def _toggle_tracking_status(value_key: str, stamp_key: str,
                            qty_key: str | None = None, full_qty: float = 0.0):
    checked = not bool(st.session_state.get(value_key))
    st.session_state[value_key] = checked
    if checked and not st.session_state.get(stamp_key):
        st.session_state[stamp_key] = dt.datetime.now().isoformat(timespec="minutes")
    elif not checked:
        st.session_state[stamp_key] = ""
    if qty_key:
        st.session_state[qty_key] = _bounded_tracking_qty(full_qty, full_qty) if checked else 0.0


def _handle_tracking_status_click(value_key: str, stamp_key: str,
                                  qty_key: str | None = None,
                                  full_qty: float = 0.0,
                                  action: str = "",
                                  description: str = ""):
    if qty_key and not bool(st.session_state.get(value_key)):
        current_qty = _bounded_tracking_qty(st.session_state.get(qty_key), full_qty)
        if current_qty <= 0:
            _open_tracking_qty_prompt(value_key, stamp_key, qty_key,
                                      action, description, full_qty)
            return
    _toggle_tracking_status(value_key, stamp_key, qty_key, full_qty)


def _sync_tracking_qty_status(qty_key: str, value_key: str, stamp_key: str, max_qty: float):
    _set_tracking_qty(value_key, stamp_key, qty_key,
                      st.session_state.get(qty_key), max_qty)


def _tracking_status_cell(col, lid: int, key_name: str, current: bool, stamp_value,
                          full_qty: float | None = None, current_qty=0.0,
                          description: str = ""):
    value_key, stamp_key, qty_key = _tracking_keys(lid, key_name)
    if value_key not in st.session_state:
        st.session_state[value_key] = bool(current)
    if stamp_key not in st.session_state:
        st.session_state[stamp_key] = _text(stamp_value)
    max_qty = _bounded_tracking_qty(full_qty or 0.0, full_qty or 0.0)
    if full_qty is not None:
        default_qty = _bounded_tracking_qty(current_qty, max_qty)
        if bool(st.session_state[value_key]) and default_qty <= 0 and max_qty > 0:
            default_qty = max_qty
        if qty_key not in st.session_state:
            st.session_state[qty_key] = default_qty
        st.session_state[qty_key] = _bounded_tracking_qty(st.session_state.get(qty_key), max_qty)

    checked = bool(st.session_state[value_key])
    stamp = _text(st.session_state.get(stamp_key))
    if checked and not stamp:
        stamp = dt.datetime.now().isoformat(timespec="minutes")
        st.session_state[stamp_key] = stamp
    elif not checked:
        stamp = ""
        st.session_state[stamp_key] = ""

    btn_state = "on" if checked else "off"
    if full_qty is not None and checked:
        tracked_qty = _bounded_tracking_qty(st.session_state.get(qty_key), max_qty)
        if max_qty > 0 and 0 < tracked_qty < max_qty:
            btn_state = "partial"
    btn_key = f"trkbtn_{btn_state}_{key_name}_{lid}"
    label = "âœ“" if checked else " "
    col.button(label, key=btn_key, disabled=not can("tracking"), use_container_width=True,
               on_click=_handle_tracking_status_click,
               args=(value_key, stamp_key,
                     qty_key if full_qty is not None else None,
                     _bounded_tracking_qty(full_qty or 0.0, full_qty or 0.0),
                     key_name, description))
    stamp_text = html.escape(_fmt_tracking_stamp(stamp)) if stamp else "&nbsp;"
    col.markdown(f"<div class='tracking-stamp'>{stamp_text}</div>", unsafe_allow_html=True)
    return checked, stamp


def _tracking_qty_cell(col, lid: int, key_name: str, current_qty, line_qty: float) -> float:
    value_key, stamp_key, qty_key = _tracking_keys(lid, key_name)
    max_qty = max(float(line_qty or 0.0), 0.0)
    default_qty = _bounded_tracking_qty(current_qty, max_qty)
    if bool(st.session_state.get(value_key)) and default_qty <= 0 and max_qty > 0:
        default_qty = max_qty
    if qty_key not in st.session_state:
        st.session_state[qty_key] = default_qty
    st.session_state[qty_key] = _bounded_tracking_qty(st.session_state.get(qty_key), max_qty)
    col.number_input(
        f"{key_name} quantity",
        min_value=0.0,
        max_value=max_qty,
        step=1.0,
        format="%.2f",
        key=qty_key,
        label_visibility="collapsed",
        disabled=not can("tracking"),
        on_change=_sync_tracking_qty_status,
        args=(qty_key, value_key, stamp_key, max_qty),
    )
    return _bounded_tracking_qty(st.session_state.get(qty_key), max_qty)


def _render_finance_tab(project_id: int, grand_total: float):
    """Two side-by-side tables for an offer: client payments/invoices and purchases/costs."""
    if not can("finance"):
        st.info("ðŸ”’ Your role doesn't have Finance access. "
                "An owner can grant it in Settings â†’ Roles & permissions.")
        return

    gt = float(grand_total or 0.0)
    # Cache the editor sources once per offer so in-progress edits stay consistent
    # (we never reload from the DB mid-edit; we just persist changes back to it).
    pay_src_key, pur_src_key = f"fin_pay_src_{project_id}", f"fin_pur_src_{project_id}"
    if pay_src_key not in st.session_state or pur_src_key not in st.session_state:
        pays, purs = repo.get_finance(project_id)
        st.session_state[pay_src_key] = pd.DataFrame(
            [{"Description": r["Description"], "Amount (SAR)": r["AmountSAR"],
              "Invoice #": r["InvoiceNo"] or ""} for r in pays]
            or [{"Description": d, "Amount (SAR)": 0.0, "Invoice #": ""}
                for d in ("Downpayment", "Payment 1", "Payment 2")])
        st.session_state[pur_src_key] = pd.DataFrame(
            [{"Description": r["Description"], "Cost (SAR)": r["AmountSAR"],
              "PO #": r["PORef"] or ""} for r in purs]
            or [{"Description": "", "Cost (SAR)": 0.0, "PO #": ""}])

    col_pay, col_pur = st.columns(2)
    with col_pay:
        st.markdown("#### ðŸ’µ Payments / Invoices")
        pay_cfg = {
            "Description": st.column_config.TextColumn("Payment Description", width="medium"),
            "Amount (SAR)": st.column_config.NumberColumn("Amount (SAR)", format="%.2f", min_value=0.0),
            "Invoice #": st.column_config.TextColumn("Invoice #", help="Invoice number (free text)"),
        }
        pay_edit = st.data_editor(st.session_state[pay_src_key], column_config=pay_cfg,
                                  num_rows="dynamic", hide_index=True, use_container_width=True,
                                  key=f"fin_pay_{project_id}")
        collected = pay_edit["Amount (SAR)"].map(calc._num).sum()
        remaining = gt - collected
        pct = (collected / gt * 100) if gt else 0.0
        _fin_bubble("Collected", f"SAR {collected:,.2f}",
                    [f"{pct:.0f}% collected", f"of SAR {gt:,.0f}"],
                    "Remaining / Due", f"SAR {remaining:,.2f}",
                    positive=remaining >= 0)

    with col_pur:
        st.markdown("#### ðŸ§¾ Purchases / Costs")
        pur_cfg = {
            "Description": st.column_config.TextColumn("Dispense Description", width="medium"),
            "Cost (SAR)": st.column_config.NumberColumn("Cost (SAR)", format="%.2f", min_value=0.0),
            "PO #": st.column_config.TextColumn("PO #", help="Purchase-order reference (free text)"),
        }
        pur_edit = st.data_editor(st.session_state[pur_src_key], column_config=pur_cfg,
                                  num_rows="dynamic", hide_index=True, use_container_width=True,
                                  key=f"fin_pur_{project_id}")
        cost_total = pur_edit["Cost (SAR)"].map(calc._num).sum()
        vat = gt * calc.VAT_RATE
        net_profit = gt - cost_total - vat
        markup = (gt / cost_total) if cost_total > 0 else None
        margin_pct = (net_profit / gt * 100) if gt else 0.0
        markup_txt = f"Markup Ã-{markup:.2f}" if markup else "Markup â€”"
        _fin_bubble("ðŸ§¾ Cost (POs)", f"SAR {cost_total:,.2f}",
                    [markup_txt, f"Margin {margin_pct:.1f}%",
                     f"VAT ({calc.VAT_RATE * 100:g}%) SAR {vat:,.0f}"],
                    "ðŸ’° Net Profit", f"SAR {net_profit:,.2f}",
                    positive=net_profit >= 0)

    # Auto-save on any change (no Save button).
    sig = (pay_edit.to_json(), pur_edit.to_json())
    sig_key = f"fin_sig_{project_id}"
    if sig_key not in st.session_state:
        st.session_state[sig_key] = sig
    elif st.session_state[sig_key] != sig:
        repo.save_finance(project_id, pay_edit.to_dict("records"), pur_edit.to_dict("records"))
        st.session_state[sig_key] = sig
        st.toast("Finance saved", icon="ðŸ’¾")


def _render_tracking_tab(project_id: int, sheet_name: str | None):
    track = repo.load_tracking(project_id, sheet_name).reset_index(drop=True)
    if track.empty:
        st.info("No line items to track.")
        return

    _render_tracking_qty_prompt()

    cols_meta = [("Description", 2.75), ("Brand", 1.0), ("Model", 1.05), ("Qty", 0.5),
                 ("PO Number", 1.25), ("Paid", 0.9), ("Received", 0.9), ("Rec. Qty", 0.75),
                 ("Delivery Note", 1.25), ("Delivered", 0.9), ("Deliv. Qty", 0.75)]
    widths = [w for _, w in cols_meta]
    hdr = st.columns(widths)
    for (label, _), col in zip(cols_meta, hdr):
        col.markdown(f"<div class='tracking-header'>{label}</div>", unsafe_allow_html=True)
    st.markdown("<hr style='margin:0.1rem 0 0.3rem;border:none;"
                "border-top:1px solid #e6e6e6'>", unsafe_allow_html=True)

    collected = []
    for row_idx, (_, row) in enumerate(track.iterrows()):
        lid = int(row["LineID"])
        rc = st.columns(widths, vertical_alignment="center")
        j = 0
        rc[j].write(row["Description"] or ""); j += 1
        rc[j].write(row["Brand"] or ""); j += 1
        rc[j].write(row["Model"] or ""); j += 1
        line_qty = max(_safe_float(row.get("Qty")), 0.0)
        qty_text = f"{int(line_qty)}" if float(line_qty).is_integer() else f"{line_qty:.2f}"
        _tracking_center_text(rc[j], qty_text if pd.notna(row["Qty"]) else ""); j += 1
        po = rc[j].text_input("po", value=str(row.get("PONumber") or ""),
                              key=f"po_{lid}", label_visibility="collapsed"); j += 1
        paid, paid_at = _tracking_status_cell(rc[j], lid, "paid", bool(row["Paid"]), row.get("PaidAt")); j += 1
        rec_current_qty = _bounded_tracking_qty(row.get("ReceivedQty"), line_qty)
        rec_current = bool(row["Received"]) or rec_current_qty > 0
        rec, rec_at = _tracking_status_cell(rc[j], lid, "received", rec_current,
                                            row.get("ReceivedAt"), full_qty=line_qty,
                                            current_qty=rec_current_qty,
                                            description=_text(row["Description"])); j += 1
        rec_qty = _tracking_qty_cell(rc[j], lid, "received", rec_current_qty, line_qty); j += 1
        delivery_note = rc[j].text_input("delivery note", value=str(row.get("DeliveryNote") or ""),
                                         key=f"dn_{lid}", label_visibility="collapsed"); j += 1
        deliv_current_qty = _bounded_tracking_qty(row.get("DeliveredQty"), line_qty)
        deliv_current = bool(row["Delivered"]) or deliv_current_qty > 0
        deliv, deliv_at = _tracking_status_cell(rc[j], lid, "delivered", deliv_current,
                                                row.get("DeliveredAt"), full_qty=line_qty,
                                                current_qty=deliv_current_qty,
                                                description=_text(row["Description"])); j += 1
        deliv_qty = _tracking_qty_cell(rc[j], lid, "delivered", deliv_current_qty, line_qty); j += 1
        rec = rec or rec_qty > 0
        deliv = deliv or deliv_qty > 0
        _, rec_stamp_key, _ = _tracking_keys(lid, "received")
        _, deliv_stamp_key, _ = _tracking_keys(lid, "delivered")
        rec_at = _text(st.session_state.get(rec_stamp_key))
        deliv_at = _text(st.session_state.get(deliv_stamp_key))
        collected.append((lid, paid, rec, deliv, po, delivery_note, paid_at, rec_at, deliv_at,
                          rec_qty, deliv_qty))
        if row_idx < len(track) - 1:
            st.markdown("<div class='tracking-row-separator'></div>", unsafe_allow_html=True)

    tot = len(track)
    total_qty = sum(max(_safe_float(row.get("Qty")), 0.0) for _, row in track.iterrows())
    rec_total = sum(c[9] for c in collected)
    deliv_total = sum(c[10] for c in collected)
    st.markdown("<div style='height:1.75rem'></div>", unsafe_allow_html=True)  # gap before totals
    pc1, pc2, pc3 = st.columns(3)
    pc1.metric("Paid", f"{sum(1 for c in collected if c[1])}/{tot}")
    pc2.metric("Received Qty", f"{rec_total:g}/{total_qty:g}")
    pc3.metric("Delivered Qty", f"{deliv_total:g}/{total_qty:g}")
    if can("tracking"):
        sig = repr(collected)                       # auto-save on any change (no button)
        sig_key = f"trk_sig_{project_id}_{sheet_name or 'all'}"
        if sig_key not in st.session_state:
            st.session_state[sig_key] = sig
        elif st.session_state[sig_key] != sig:
            repo.update_tracking(collected)
            st.session_state[sig_key] = sig
            st.toast("Tracking saved", icon="ðŸ’¾")
    else:
        st.caption("ðŸ”’ Your role can view tracking but not change it.")


def _render_login():
    c = st.columns([1, 2, 1])[1]
    if os.path.exists(_LOGO):
        c.image(_LOGO, use_container_width=True)
    c.subheader("Sign in")
    if auth.user_count() == 0:
        c.info("First run â€” create the **owner** account (full access).")
        with c.form("create_owner"):
            u = st.text_input("Username")
            dn = st.text_input("Display name")
            p1 = st.text_input("Password", type="password")
            p2 = st.text_input("Confirm password", type="password")
            if st.form_submit_button("Create owner & sign in", type="primary"):
                if not u.strip() or not p1:
                    st.error("Username and password are required.")
                elif p1 != p2:
                    st.error("Passwords don't match.")
                elif auth.create_user(u, p1, dn, role="owner"):
                    st.session_state.auth_user = auth.verify_login(u, p1)
                    st.rerun()
                else:
                    st.error("Could not create user (name already taken?).")
    else:
        with c.form("login"):
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Sign in", type="primary"):
                user = auth.verify_login(u, p)
                if user:
                    st.session_state.auth_user = user
                    st.rerun()
                else:
                    st.error("Invalid username or password (or account disabled).")


# ----------------------------- UI -----------------------------
if "db_init" not in st.session_state:
    db.init_db()                       # create Users table + apply migrations
    auth.ensure_roles_seeded()         # seed default roles/permissions on first run
    st.session_state.db_init = True

if "auth_user" not in st.session_state:
    _render_login()
    st.stop()

USER = st.session_state.auth_user
ROLE = USER.get("Role", "viewer")
PERMS = auth.role_perms(ROLE)
def can(p):
    return p in PERMS

_ensure_state()
# Sidebar shows the standalone LOGO (full width); falls back to the banner if no logo.
if os.path.exists(db.logo_path()):
    st.sidebar.image(db.logo_path(), use_container_width=True)
elif os.path.exists(_LOGO):
    st.sidebar.image(_LOGO, use_container_width=True)
st.sidebar.title(_COMPANY)
st.sidebar.caption(f"ðŸ‘¤ **{USER.get('DisplayName') or USER.get('Username')}** Â· _{ROLE}_")
if st.sidebar.button("ðŸ”’ Log out", use_container_width=True):
    st.session_state.pop("auth_user", None)
    st.rerun()

_SECTIONS = [("New Project", "new_offer"), ("Load Project", "load"),
             ("Catalogue", "catalogue"), ("Settings", "settings"), ("Users", "users")]
_allowed = [name for name, p in _SECTIONS if can(p)]
if not _allowed:
    st.error("Your account has no accessible sections â€” contact the owner.")
    st.stop()
_nav_mode = st.session_state.pop("_nav_mode", None)
if _nav_mode in _allowed:
    st.session_state["workspace_mode"] = _nav_mode
if st.session_state.get("workspace_mode") not in _allowed:
    st.session_state["workspace_mode"] = _allowed[0]
mode = st.sidebar.radio("Workspace", _allowed, key="workspace_mode")
admin = can("view_costs")          # on-screen internal cost metrics (client PDF never shows costs)

# Refresh per-currency -> USD rates from Settings (EUR configurable; SAR pegged).
try:
    calc.CURRENCY_RATES["EUR"] = float(repo.get_setting("eur_to_usd") or 1.08)
except (TypeError, ValueError):
    pass
calc.CURRENCY_RATES["SAR"] = 1.0 / calc.SAR_PER_USD
calc.CURRENCY_RATES["AED"] = 1.0 / calc.AED_PER_USD
try:
    calc.VAT_RATE = float(repo.get_setting("vat_percent") or 15) / 100
except (TypeError, ValueError):
    pass

# DB stats
try:
    nproj = len(repo.list_projects())
    ncat = len(repo.search_catalog("", limit=100000))
    st.sidebar.metric("Projects in DB", nproj)
    st.sidebar.metric("Catalogue items", ncat)
except Exception as e:
    st.sidebar.error(f"DB: {e}")


# ============================ NEW OFFER ============================
if mode == "New Project":
    duplicate = st.session_state.pop("_duplicate_offer", None)
    if duplicate:
        _prime_new_offer_form(duplicate["header"], duplicate["grid"], duplicate["discount"])
        st.success(f"Duplicated from **{duplicate['source']}**. Review and save as a new offer.")
    elif st.session_state.pop("_no_reset_all", False):
        _prime_new_offer_form()
    elif st.session_state.pop("_no_reset_option", False):
        st.session_state["no_option"] = ""
        st.session_state.pop("editor", None)

    st.subheader("New Project")
    h = st.session_state.header

    # Live offer reference (from the System Offer + override below) â€” shown as a top bar.
    # Once the first option is saved, the offer # is "locked" so further options share it.
    _sel = st.session_state.get("no_offer_type", "(none)")
    _otype = "" if _sel == "(none)" else _sel
    _ov = (st.session_state.get("no_offer_ov") or "").strip()
    h["system"] = _otype
    h["offer_override"] = _ov
    _locked = st.session_state.get("no_offer_lock")
    h["offer"] = _locked or _ov or repo.make_offer_no(_otype)
    _saved = st.session_state.get("no_saved_options", [])
    _extra = (f"<span style='font-size:.8rem;opacity:.85'> &nbsp;Â·&nbsp; options saved: "
              f"{', '.join(_saved)}</span>") if _saved else ""
    st.markdown(
        f"<div style='background:#002060;color:#fff;padding:10px 16px;border-radius:8px;"
        f"font-size:1.2rem;margin:2px 0 12px'>ðŸ§¾&nbsp;&nbsp;Offer #:&nbsp; <b>{h['offer']}</b>{_extra}</div>",
        unsafe_allow_html=True)

    with st.expander("Offer header", expanded=True):
        c1, c2, c3 = st.columns(3)
        h["client"] = c1.text_input("Client", h["client"], key="no_client")
        h["project"] = c1.text_input("Project", h["project"], key="no_project")
        h["contact"] = c2.text_input("Contact", h["contact"], key="no_contact")
        h["phone"] = c2.text_input("Phone", h["phone"], key="no_phone")
        p1, p2, p3 = st.columns(3)
        h["sales"] = _person_select(p1, "Sales Person", PEOPLE_ROLES["sales"],
                                    h.get("sales", ""), "no_sales")
        h["presales"] = _person_select(p2, "Pre-sales Engineer", PEOPLE_ROLES["presales"],
                                       h.get("presales", ""), "no_presales")
        h["pm"] = _person_select(p3, "Project Manager", PEOPLE_ROLES["pm"],
                                 h.get("pm", ""), "no_pm")
        # "System Offer" drives BOTH the offer-ref type segment and the BOQ system suffix.
        c3.selectbox("System Offer", ["(none)"] + repo.offer_types(), key="no_offer_type",
                     help="The system being quoted (AV, LCS, â€¦). Used in the offer reference "
                          "and as the BOQ system. Manage the list in Settings.")
        c3.text_input("Offer # (blank = auto)", key="no_offer_ov",
                      help="Leave blank to auto-number; type a value to override.")
        h["option"] = c3.text_input("Option label (optional)", key="no_option",
                                    help="Name this alternative (e.g. Dynalite, KNX). "
                                         "Leave blank for a single-option offer.")

    terms_form(st.session_state.header, "no")
    project_sheet_info_form(st.session_state.header, "no_ps")

    # ---- Add items from catalogue ----
    _dm = float(repo.get_setting("default_margin") or 1.6)   # default margin from Settings
    catalogue_add("grid", _dm, "no", st.session_state.header["system"], show_clear=True)

    # ---- Editable grid (builder always shows costs) ----
    st.caption("Edit **Qty Â· Ex Unit Cost Â· Shipping % Â· Margin Ã-** â†’ prices recalc automatically. "
               "Locked columns are computed.")
    grid = render_editable_grid("grid", "editor")

    # ---- Discount + totals ----
    st.markdown("##### Totals")
    calc_grid = calc.recompute(st.session_state.grid)
    base_summary = calc.summarize(calc_grid, 0)
    dcol, pcol, _ = st.columns([1, 1, 2])
    discount_sar = _discount_inputs("no", "discount", base_summary["subtotal_sar"], dcol, pcol)
    s = calc.summarize(calc_grid, discount_sar)

    m1, m2, m3 = st.columns(3)
    _subtotal_metric(m1, s)
    m2.metric(f"VAT {calc.VAT_RATE * 100:g}% (SAR)", f"{s['vat_amount_sar']:,.0f}")
    m3.metric("Grand Total (SAR)", f"{s['grand_total_sar']:,.0f}")
    _profit_banner(s)
    if admin:
        a1, a2, a3 = st.columns(3)
        a1.metric("Total Cost (USD)", f"{s['total_cost_usd']:,.2f}")
        a2.metric("Cost in SAR", f"{s['cost_sar']:,.0f}")
        a3.metric("Total Selling (USD)", f"{s['total_sell_usd']:,.2f}")

    # ---- Actions ----
    st.divider()
    if st.session_state.get("no_offer_lock"):
        st.caption(f"Adding options to **{h['offer']}** â€” build this option, name it, then "
                   "**Save option**. Use **âž• Add another option** to start the next one, or "
                   "**ðŸ†• New offer** to begin a fresh offer.")
    ac1, ac2, ac3, ac4, ac5 = st.columns(5)
    _optname = (h.get("option") or "").strip()
    if ac1.button("ðŸ’¾ Save option" if st.session_state.get("no_offer_lock") else "ðŸ’¾ Save offer",
                  type="primary", use_container_width=True):
        _locked_now = st.session_state.get("no_offer_lock")
        _done = st.session_state.get("no_saved_options", [])
        if st.session_state.grid.empty:
            st.warning("Grid is empty.")
        elif _locked_now and not _optname:
            st.warning("Enter an Option label for this alternative (e.g. KNX).")
        elif _optname and _optname in _done:
            st.warning(f"Option '{_optname}' is already saved for this offer.")
        else:
            name = (h["project"] or "Untitled") + (f" ({_optname})" if _optname else "")
            pid = repo.save_offer(
                name=name, client=h["client"], contact=h["contact"],
                offer_no=h["offer"], system_suffix=h["system"],
                grid=calc.recompute(st.session_state.grid),
                discount_sar=discount_sar,
                factors=(s["markup_factor"], None, None),
                sales_person=h.get("sales"), presales_engineer=h.get("presales"),
                project_manager=h.get("pm"),
                terms={k: h.get(k) for k in TERMS_KEYS}, option_label=_optname,
                project_sheet_info=h.get("project_sheet"))
            st.session_state.no_offer_lock = h["offer"]          # lock # for further options
            st.session_state.setdefault("no_saved_options", []).append(_optname or "Main")
            st.success(f"Saved {('option ' + _optname) if _optname else 'offer'} (ProjectID {pid}).")

    if ac2.button("âž• Add another option", use_container_width=True,
                  disabled=not st.session_state.get("no_offer_lock")):
        st.session_state.grid = _empty_grid()
        st.session_state["_no_reset_option"] = True   # clear option label on next run
        st.rerun()

    if ac3.button("ðŸ†• New offer", use_container_width=True):
        st.session_state.grid = _empty_grid()
        st.session_state.no_offer_lock = None
        st.session_state.no_saved_options = []
        st.session_state["_no_reset_all"] = True       # clear option label + override
        st.session_state.pop("pdf_bytes", None)
        st.session_state.pop("project_sheet_bytes", None)
        st.rerun()

    if ac4.button("ðŸ“„ Generate Offer PDF", use_container_width=True):
        _make_pdf_download(h, st.session_state.grid, s)

    if ac5.button("ðŸ“Š Generate Project Sheet", use_container_width=True):
        _make_project_sheet_download(h, s)

    dl1, dl2 = st.columns(2)
    if "pdf_bytes" in st.session_state:
        dl1.download_button("â¬‡ï¸ Download Offer PDF", st.session_state.pdf_bytes,
                            file_name=f"Quotation_{h['offer']}{(' '+_optname) if _optname else ''}.pdf",
                            mime="application/pdf", use_container_width=True)
    if "project_sheet_bytes" in st.session_state:
        dl2.download_button("â¬‡ï¸ Download Project Sheet", st.session_state.project_sheet_bytes,
                            file_name=f"Project_Sheet_{_safe_filename(h.get('offer') or h.get('project'))}.xlsx",
                            mime=("application/vnd.openxmlformats-officedocument."
                                  "spreadsheetml.sheet"),
                            use_container_width=True)


# ============================ LOAD EXISTING ============================
elif mode == "Load Project":
    st.subheader("Load Project")
    if st.session_state.pop("_del_reset", False):        # clear delete confirm widgets
        st.session_state.pop("del_confirm", None)
        st.session_state.pop("del_scope", None)
    projects = repo.list_projects()
    if projects.empty:
        st.info("No projects ingested yet. Run `python ingest.py`.")
    else:
        # Group revisions into offer families; pick a family, then a revision.
        projects["_fam"] = projects.apply(
            lambda r: repo.family_key(r.get("OfferNo"), r.get("ProjectName")), axis=1)
        fams = []
        for fam, grp in projects.groupby("_fam", sort=False):
            grp = grp.sort_values(["RevisionNo", "OptionLabel"], na_position="first")
            rep = grp.iloc[-1]
            base = _text(rep.get("BaseName")) or repo.base_name(_text(rep.get("ProjectName"), "Offer"))
            client = _text(rep.get("ClientName"))
            offer_nos = sorted({_text(v) for v in grp["OfferNo"].tolist() if _text(v)})
            project_names = sorted({_text(v) for v in grp["ProjectName"].tolist() if _text(v)})
            n_rev = int(grp["RevisionNo"].fillna(0).astype(int).nunique())          # distinct revisions
            n_opt = int(grp.groupby(grp["RevisionNo"].fillna(0)).size().max())      # most options in a revision
            fams.append({"fam": fam, "base": base, "client": client,
                         "offer_nos": offer_nos, "project_names": project_names,
                         "name_search": " ".join([base, client] + project_names).lower(),
                         "offer_search": " ".join(offer_nos).lower(),
                         "n_rev": n_rev, "n_opt": n_opt,
                         "approved": bool(grp["Approved"].fillna(0).max()),
                         "date": _text(grp["CreationDate"].max())})
        fams.sort(key=lambda f: f["date"], reverse=True)

        def _famlabel(f):
            parts = [_text(f["base"], "Offer")] + ([_text(f["client"])] if _text(f["client"]) else [])
            if f["offer_nos"]:
                parts.append(", ".join(f["offer_nos"][:3]))
            parts.append(f"{f['n_rev']} rev. - {f['n_opt']} opt.")
            return ("âœ… " if f["approved"] else "") + " Â· ".join(parts)

        sc1, sc2 = st.columns([2, 1])
        q_name = _text(sc1.text_input("Search by name", key="load_search_name")).lower()
        q_offer = _text(sc2.text_input("Search by offer #", key="load_search_offer")).lower()
        query_key = f"{q_name}\0{q_offer}"
        if st.session_state.get("load_query") != query_key:
            st.session_state.load_query = query_key
            st.session_state.pop("view_pid", None)
            st.session_state.pop("load_fam", None)
            st.session_state.pop("pdf_bytes", None)
            st.session_state.pop("project_sheet_bytes", None)
            st.session_state.edit_mode = False

        if not q_name and not q_offer:
            st.session_state.pop("view_pid", None)
            st.session_state.pop("load_fam", None)
            st.info("Search by project name, client name, or offer number.")
            st.stop()

        matches = [
            f for f in fams
            if (not q_name or q_name in f["name_search"])
            and (not q_offer or q_offer in f["offer_search"])
        ]
        if not matches:
            st.session_state.pop("view_pid", None)
            st.session_state.pop("load_fam", None)
            st.warning("No matching offers found.")
            st.stop()

        current_fam = st.session_state.get("load_fam")
        match_fams = {f["fam"] for f in matches}
        if current_fam not in match_fams:
            current_fam = None
            st.session_state.pop("load_fam", None)
            st.session_state.pop("view_pid", None)

        st.markdown("**Matching offers**")
        widths = [2.4, 1.4, 2.3, 0.9, 0.7, 0.7, 0.9, 0.8]
        hc = st.columns(widths)
        for col, t in zip(hc, ["Project", "Client", "Offer #", "Date", "Rev.", "Opt.", "Approved", ""]):
            col.caption(t)
        for idx, f in enumerate(matches):
            selected = f["fam"] == current_fam
            rc = st.columns(widths, vertical_alignment="center")
            rc[0].write(("â–¶ " if selected else "") + _text(f["base"], "Offer"))
            rc[1].write(_text(f["client"], "-"))
            rc[2].write(", ".join(f["offer_nos"][:3]) if f["offer_nos"] else "-")
            rc[3].write(_text(f["date"])[:10])
            rc[4].write(str(f["n_rev"]))
            rc[5].write(str(f["n_opt"]))
            rc[6].write("âœ…" if f["approved"] else "")
            if rc[7].button("View", key=f"match_view_{idx}_{f['fam']}",
                            disabled=selected, use_container_width=True):
                st.session_state.load_fam = f["fam"]
                st.session_state.pop("view_pid", None)
                st.session_state.pop("pdf_bytes", None)
                st.session_state.pop("project_sheet_bytes", None)
                st.session_state.edit_mode = False
                st.rerun()

        if not current_fam:
            st.info(f"{len(matches)} matching offer{'s' if len(matches) != 1 else ''} found. Click View to open one.")
            st.stop()

        selected_fam = current_fam

        grp = (projects[projects["_fam"] == selected_fam]
               .sort_values(["RevisionNo", "OptionLabel"], na_position="first"))

        # Hide archived entries by default; toggle to reveal them.
        n_arch = int(grp["Archived"].fillna(0).sum())
        show_arch = st.checkbox(f"Show archived ({n_arch})", value=False, key="show_archived",
                                disabled=not n_arch)
        shown = grp if show_arch else grp[grp["Archived"].fillna(0) == 0]
        if shown.empty:                                  # all archived -> show them anyway
            shown = grp

        # Default selection: the approved entry, else the newest active, else newest.
        rev_ids = shown["ProjectID"].tolist()
        approved_ids = shown[shown["Approved"].fillna(0) == 1]["ProjectID"].tolist()
        active_ids = shown[shown["Archived"].fillna(0) == 0]["ProjectID"].tolist()
        default_id = int(approved_ids[-1] if approved_ids
                         else (active_ids[-1] if active_ids else rev_ids[-1]))
        if st.session_state.get("view_pid") not in rev_ids:
            st.session_state.view_pid = default_id

        # Revisions & options, each with its own View button.
        st.markdown("**Revisions & options**")
        widths = [1.0, 1.2, 1.8, 1.0, 0.5, 1.0, 0.9]
        hc = st.columns(widths)
        for col, t in zip(hc, ["Revision", "Option", "Offer #", "Date", "âœ“", "Status", ""]):
            col.caption(t)
        for _, row in shown.iterrows():
            rid = int(row["ProjectID"])
            rn = int(row["RevisionNo"]) if pd.notna(row["RevisionNo"]) else 0
            sel = (rid == int(st.session_state.view_pid))
            rc = st.columns(widths)
            rc[0].write(("â–¶ " if sel else "") + (repo.revision_token(rn) if rn > 0 else "Original"))
            rc[1].write(_text(row["OptionLabel"], "-"))
            rc[2].write(_text(row["OfferNo"]))
            rc[3].write(_text(row["CreationDate"])[:10])
            rc[4].write("âœ…" if row["Approved"] else "")
            rc[5].write("ðŸ“¦ Archived" if row["Archived"] else "Active")
            if rc[6].button("View", key=f"view_{rid}", disabled=sel, use_container_width=True):
                st.session_state.view_pid = rid
                st.session_state.pop("pdf_bytes", None)
                st.session_state.pop("project_sheet_bytes", None)
                st.rerun()

        pid = int(st.session_state.view_pid)
        systems = repo.list_systems(pid)
        sheet = systems[0] if systems else None      # auto-pick the system sheet
        meta = repo.project_meta(pid)
        cur_key = f"{pid}::{sheet}"
        editing = (st.session_state.get("edit_mode")
                   and st.session_state.get("edit_key") == cur_key)

        if not editing:
            # -------------------- VIEW: tabbed offer view --------------------
            grid = repo.load_project_grid(pid, sheet)
            disp = grid.copy()
            for col in MONEY_COLS:
                if col in disp.columns:
                    disp[col] = disp[col].map(lambda v: calc.roundup(v, 0))
            s = calc.summarize(disp, meta.get("DiscountAmount") or 0)

            rev = meta.get("RevisionNo") or 0
            _subj = repo.load_terms(meta).get("subject")
            if _subj:
                st.markdown(f"#### ðŸ“„ {_subj}")
            _opt = meta.get("OptionLabel") or ""
            _ttl = (repo.revision_token(rev) if rev else "original") + (f" Â· Option: {_opt}" if _opt else "")
            if meta.get("Approved"):
                st.caption(f"âœ… Approved Â· {_ttl}")
            else:
                st.caption(f"Active Â· {_ttl} â€” click **Edit** to make a new revision or option.")

            active_tab = _offer_tab_selector(pid, bool(meta.get("Approved")))
            if active_tab == "BoQ":
                _summary_metrics(s)

                # Approval + archive controls (BoQ tab only).
                apc1, apc2, apc3 = st.columns([2.2, 1, 1], vertical_alignment="center")
                if meta.get("Archived"):
                    apc1.warning("ðŸ“¦ **Archived**" + (" Â· was âœ… approved" if meta.get("Approved") else ""))
                elif meta.get("Approved"):
                    at = (meta.get("ApprovedAt") or "")[:16].replace("T", " ")
                    apc1.success(f"âœ… **Approved**{(' Â· ' + at) if at else ''}")
                else:
                    apc1.info("Active Â· not approved")
                if meta.get("Approved"):
                    if can("approve") and apc2.button("â†©ï¸ Unapprove", use_container_width=True):
                        r = repo.unapprove_offer(pid)
                        st.toast(f"Unapproved. {r} auto-archived entr{'y' if r == 1 else 'ies'} restored."
                                 if r else "Unapproved.", icon="â†©ï¸")
                        st.rerun()
                elif can("approve"):
                    if apc2.button("âœ… Approve", type="primary", use_container_width=True):
                        n = repo.approve_offer(pid)
                        st.toast(f"Approved. {n} other entr{'y' if n == 1 else 'ies'} archived."
                                 if n else "Approved.", icon="âœ…")
                        st.rerun()
                if can("archive"):
                    if meta.get("Archived"):
                        if apc3.button("â™»ï¸ Restore", use_container_width=True):
                            repo.unarchive_project(pid)
                            st.rerun()
                    elif apc3.button("ðŸ“¦ Archive", use_container_width=True):
                        repo.archive_project(pid)
                        st.rerun()
                if meta.get("SalesPerson") or meta.get("PresalesEngineer") or meta.get("ProjectManager"):
                    st.caption(f"ðŸ‘¤ Sales: {meta.get('SalesPerson') or 'â€”'}  Â·  "
                               f"Pre-sales Eng.: {meta.get('PresalesEngineer') or 'â€”'}  Â·  "
                               f"Project Mgr: {meta.get('ProjectManager') or 'â€”'}")
                if admin:                   # gross-profit line (internal cost view) â€” BoQ tab only
                    _profit_banner(s)
                cfg = {c: st.column_config.NumberColumn(c, format="%.0f") for c in MONEY_COLS}
                cfg["Qty"] = st.column_config.NumberColumn("Qty", format="%d")
                cfg["Shipping %"] = st.column_config.NumberColumn("Shipping %", format="%.2f")
                st.dataframe(disp[[c for c in BUILDER_COLS if c in disp.columns]],
                             use_container_width=True, hide_index=True, column_config=cfg)
            elif active_tab == "Tracking":
                _render_tracking_tab(pid, sheet)
            else:
                _render_finance_tab(pid, s["grand_total_sar"])

            b1, b2, b3, b4 = st.columns(4)
            if can("edit") and b1.button("âœï¸ Edit / new revision or option", type="primary",
                                         use_container_width=True):
                eg = repo.load_project_grid(pid, sheet).copy()
                eg["Margin x"] = 0.0   # keep loaded prices; set a margin per line to re-price
                st.session_state.edit_grid = calc.recompute(eg)
                st.session_state.edit_key = cur_key
                st.session_state.edit_pid = pid
                st.session_state.edit_system = repo.base_name(sheet or "").replace("BOQ", "").strip() or "LCS"
                st.session_state.edit_discount = abs(float(meta.get("DiscountAmount") or 0))
                st.session_state.edit_terms = {**DEFAULT_TERMS, **repo.load_terms(meta)}
                st.session_state.edit_project_sheet = {
                    **DEFAULT_PROJECT_SHEET_INFO, **repo.load_project_sheet_info(meta)
                }
                for src, key in {
                    "job_reference": "ed_ps_job_reference",
                    "sheet_date": "ed_ps_sheet_date",
                    "lead_source": "ed_ps_lead_source",
                    "commission": "ed_ps_commission",
                    "shipment_by": "ed_ps_shipment_by",
                    "downpayment_date": "ed_ps_downpayment_date",
                    "invoice_to": "ed_ps_invoice_to",
                    "delivery_instructions": "ed_ps_delivery_instructions",
                    "salesman_signature": "ed_ps_salesman_signature",
                    "gm_signature": "ed_ps_gm_signature",
                }.items():
                    st.session_state[key] = st.session_state.edit_project_sheet.get(
                        src, DEFAULT_PROJECT_SHEET_INFO.get(src, ""))
                st.session_state["ed_option"] = meta.get("OptionLabel") or ""
                for k in ("ed_discount_percent", "ed_discount_driver", "ed_discount_subtotal"):
                    st.session_state.pop(k, None)
                st.session_state.edit_mode = True
                st.session_state.pop("pdf_bytes", None)
                st.session_state.pop("project_sheet_bytes", None)
                st.session_state.pop("saved_rev", None)
                st.rerun()
            if can("new_offer") and b2.button("ðŸ“‹ Duplicate", use_container_width=True):
                dg = repo.load_project_grid(pid, sheet).copy()
                if dg.empty:
                    st.warning("This offer has no lines to duplicate.")
                else:
                    dg["Margin x"] = 0.0   # preserve copied selling prices until user re-prices
                    system_suffix = repo.base_name(sheet or "").replace("BOQ", "").strip() or "LCS"
                    copied_terms = repo.load_terms(meta)
                    copied_header = {
                        **copied_terms,
                        "client": _text(meta.get("ClientName")),
                        "project": _text(meta.get("ProjectName")),
                        "contact": _text(meta.get("ContactName")),
                        "phone": "",
                        "sales": _text(meta.get("SalesPerson")),
                        "presales": _text(meta.get("PresalesEngineer")),
                        "pm": _text(meta.get("ProjectManager")),
                        "system": system_suffix,
                        "project_sheet": repo.load_project_sheet_info(meta),
                    }
                    st.session_state["_duplicate_offer"] = {
                        "source": _text(meta.get("OfferNo")) or _text(meta.get("ProjectName"), "existing offer"),
                        "header": copied_header,
                        "grid": calc.recompute(dg),
                        "discount": abs(float(meta.get("DiscountAmount") or 0)),
                    }
                    st.session_state["_nav_mode"] = "New Project"
                    st.session_state.edit_mode = False
                    st.rerun()
            opts = revision_options(pid)
            _multi = len(opts) > 1
            _export_header = {**DEFAULT_TERMS, **repo.load_terms(meta),
                              "client": meta.get("ClientName"), "project": meta.get("ProjectName"),
                              "contact": meta.get("ContactName"), "phone": "",
                              "sales": meta.get("SalesPerson"), "presales": meta.get("PresalesEngineer"),
                              "pm": meta.get("ProjectManager"),
                              "offer": meta.get("OfferNo"), "date": meta.get("CreationDate"),
                              "project_sheet": repo.load_project_sheet_info(meta)}
            if b3.button(f"ðŸ“„ Generate Offer PDF{f' ({len(opts)} options)' if _multi else ''}",
                         use_container_width=True):
                _make_pdf_download(_export_header, disp, s, options=opts if _multi else None)
            if b4.button("ðŸ“Š Generate Project Sheet", use_container_width=True):
                _make_project_sheet_download(_export_header, s)
            dl1, dl2 = st.columns(2)
            if "pdf_bytes" in st.session_state and not st.session_state.get("saved_rev"):
                dl1.download_button(
                    "â¬‡ï¸ Download Offer PDF", st.session_state.pdf_bytes,
                    file_name=f"Quotation_{meta.get('OfferNo') or meta.get('ProjectName')}.pdf",
                    mime="application/pdf", use_container_width=True)
            if "project_sheet_bytes" in st.session_state and not st.session_state.get("saved_rev"):
                dl2.download_button(
                    "â¬‡ï¸ Download Project Sheet", st.session_state.project_sheet_bytes,
                    file_name=f"Project_Sheet_{_safe_filename(meta.get('OfferNo') or meta.get('ProjectName'))}.xlsx",
                    mime=("application/vnd.openxmlformats-officedocument."
                          "spreadsheetml.sheet"),
                    use_container_width=True)

            if can("delete"):
              with st.expander("ðŸ-‘ï¸ Deleteâ€¦"):
                _rn = int(meta.get("RevisionNo") or 0)
                _rlbl = repo.revision_token(_rn) if _rn > 0 else "Original"
                _opt = meta.get("OptionLabel") or "â€”"
                scopes = {
                    f"This option only  ({_rlbl} Â· option {_opt})": "option",
                    f"This revision  ({_rlbl} and all its options)": "revision",
                    "This entire offer  (all revisions & options)": "offer",
                }
                pick = st.radio("What to delete", list(scopes), key="del_scope")
                ids = repo.deletion_ids(pid, scopes[pick])
                st.warning(f"Permanently deletes **{len(ids)}** entr"
                           f"{'y' if len(ids) == 1 else 'ies'} (and their line items). "
                           "This cannot be undone.")
                ok = st.checkbox("Yes, permanently delete", key="del_confirm")
                if st.button("ðŸ-‘ï¸ Delete now", type="primary", disabled=not ok):
                    n = repo.delete_projects(ids)
                    st.session_state.pop("view_pid", None)
                    st.session_state["_del_reset"] = True
                    st.success(f"Deleted {n} entr{'y' if n == 1 else 'ies'}.")
                    st.rerun()
        else:
            # -------------------- EDIT: all columns -> new revision OR new option --------------------
            base = meta.get("BaseName") or repo.base_name(meta.get("ProjectName") or "Offer")
            src_rev = int(meta.get("RevisionNo") or 0)
            nextrev = repo.next_revision(base)
            st.info(f"âœï¸ Editing **{meta.get('ProjectName')}**. Save as a **new revision** "
                    f"(â†’ {repo.revision_token(nextrev)}) â€” a changed version â€” or as **another option** of the "
                    f"current revision ({repo.revision_token(src_rev) if src_rev else 'Original'}) â€” an "
                    "alternative like Dynalite vs KNX. Loaded lines have Margin Ã- = 0 (prices kept); "
                    "set a margin to re-price a line.")
            if "edit_terms" not in st.session_state:
                st.session_state.edit_terms = {**DEFAULT_TERMS, **repo.load_terms(meta)}
            if "edit_project_sheet" not in st.session_state:
                st.session_state.edit_project_sheet = {
                    **DEFAULT_PROJECT_SHEET_INFO, **repo.load_project_sheet_info(meta)
                }
            terms_form(st.session_state.edit_terms, "ed")
            edit_header_for_ps = {
                **st.session_state.edit_terms,
                "client": meta.get("ClientName"),
                "project": meta.get("ProjectName"),
                "contact": meta.get("ContactName"),
                "phone": "",
                "sales": meta.get("SalesPerson"),
                "date": meta.get("CreationDate"),
                "project_sheet": st.session_state.edit_project_sheet,
            }
            project_sheet_info_form(edit_header_for_ps, "ed_ps")
            st.session_state.edit_project_sheet = edit_header_for_ps["project_sheet"]
            catalogue_add("edit_grid", float(repo.get_setting("default_margin") or 1.6), "ed",
                          st.session_state.get("edit_system", ""))
            grid = render_editable_grid("edit_grid", "edit_editor")

            st.markdown("##### Totals & save")
            edit_calc_grid = calc.recompute(st.session_state.edit_grid)
            edit_base_summary = calc.summarize(edit_calc_grid, 0)
            tcol, pcol, ocol = st.columns([1, 1, 2])
            edit_discount = _discount_inputs(
                "ed", "edit_discount", edit_base_summary["subtotal_sar"], tcol, pcol)
            opt_label = ocol.text_input("Option label (e.g. Dynalite, KNX)", key="ed_option",
                                        help="Names this alternative. Required for 'Save as new option'.")
            s = calc.summarize(edit_calc_grid, edit_discount)
            m1, m2, m3 = st.columns(3)
            _subtotal_metric(m1, s)
            m2.metric(f"VAT {calc.VAT_RATE * 100:g}% (SAR)", f"{s['vat_amount_sar']:,.0f}")
            m3.metric("Grand Total (SAR)", f"{s['grand_total_sar']:,.0f}")
            _profit_banner(s)

            edit_terms = st.session_state.get("edit_terms", dict(DEFAULT_TERMS))
            edit_project_sheet = st.session_state.get("edit_project_sheet", dict(DEFAULT_PROJECT_SHEET_INFO))

            def _post_save(npid, nname, nrev):
                offer_rev = repo.project_meta(npid).get("OfferNo") or nname   # actual saved offer #
                h = {**edit_terms, "client": meta.get("ClientName"), "project": nname,
                     "contact": meta.get("ContactName"), "phone": "",
                     "sales": meta.get("SalesPerson"), "presales": meta.get("PresalesEngineer"),
                     "pm": meta.get("ProjectManager"),
                     "offer": offer_rev, "date": dt.date.today().isoformat(),
                     "project_sheet": edit_project_sheet}
                _make_pdf_download(h, calc.recompute(st.session_state.edit_grid), s)
                st.session_state.saved_rev = (npid, nname, nrev)

            st.divider()
            _cur_rev = int(meta.get("RevisionNo") or 0)
            e1, e2, e3, e4 = st.columns(4)
            if e1.button("ðŸ’¾ Save on this revision", type="primary", use_container_width=True,
                         help="Overwrite the current revision/option in place â€” keeps the same "
                              "offer #, revision, option and approval."):
                repo.update_offer(
                    st.session_state.edit_pid, calc.recompute(st.session_state.edit_grid),
                    discount_sar=edit_discount,
                    factors=(s["markup_factor"], None, None),
                    system_suffix=st.session_state.get("edit_system", "LCS"), terms=edit_terms,
                    project_sheet_info=edit_project_sheet)
                _post_save(st.session_state.edit_pid, meta.get("ProjectName"), _cur_rev)
                st.success(f"Updated **{meta.get('ProjectName')}** in place.")
            if e2.button("ðŸ’¾ Save as new revision", use_container_width=True):
                npid, nname, nrev = repo.save_revision(
                    st.session_state.edit_pid, calc.recompute(st.session_state.edit_grid),
                    discount_sar=edit_discount,
                    factors=(s["markup_factor"], None, None),
                    system_suffix=st.session_state.get("edit_system", "LCS"),
                    terms=edit_terms, option_label=opt_label.strip(),
                    project_sheet_info=edit_project_sheet)
                _post_save(npid, nname, nrev)
                st.success(f"Saved **{nname}** as ProjectID {npid}.")
            if e3.button("ðŸ’¾ Save as new option", use_container_width=True):
                if not opt_label.strip():
                    st.warning("Enter an Option label first (e.g. Dynalite / KNX).")
                else:
                    npid, nname, nrev = repo.save_option(
                        st.session_state.edit_pid, calc.recompute(st.session_state.edit_grid),
                        option_label=opt_label.strip(),
                        discount_sar=edit_discount,
                        factors=(s["markup_factor"], None, None),
                        system_suffix=st.session_state.get("edit_system", "LCS"), terms=edit_terms,
                        project_sheet_info=edit_project_sheet)
                    _post_save(npid, nname, nrev)
                    st.success(f"Saved option **{nname}** as ProjectID {npid}.")
            if e4.button("âœ– Cancel edit", use_container_width=True):
                st.session_state.edit_mode = False
                for k in ("pdf_bytes", "project_sheet_bytes", "edit_terms", "edit_project_sheet", "ed_option",
                          "ed_discount_percent", "ed_discount_driver", "ed_discount_subtotal"):
                    st.session_state.pop(k, None)
                st.rerun()
            if "pdf_bytes" in st.session_state and st.session_state.get("saved_rev"):
                fn = f"Quotation_{st.session_state.saved_rev[1]}.pdf".replace(" ", "")
                st.download_button("â¬‡ï¸ Download Offer PDF", st.session_state.pdf_bytes,
                                   file_name=fn, mime="application/pdf")


# ============================ CATALOGUE ============================
elif mode == "Catalogue":
    st.subheader("Catalogue")
    _cat_edit = can("catalogue_edit")

    # ---- Add a new item ----
    if _cat_edit:
      with st.expander("âž• Add new item"):
        with st.form("add_cat_item", clear_on_submit=True):
            ac1, ac2, ac3 = st.columns(3)
            a_brand = ac1.text_input("Brand")
            a_model = ac2.text_input("Model")
            a_desc = ac3.text_input("Description")
            fc, f1, f2, f3, f4, f5, f6 = st.columns([1, 1.2, 1.4, 1.2, 1.4, 1.5, 1.6])
            a_cur = fc.selectbox("Currency", calc.CURRENCIES, index=0,
                                 help="Currency of List Price & Ex Unit Cost. "
                                      "Unit Cost is stored in USD.")
            a_list = f1.number_input("List Price", min_value=0.0, value=0.0, step=1.0)
            a_ex = f2.number_input("Ex Unit Cost", min_value=0.0, value=0.0, step=1.0)
            a_ship = f3.number_input("Shipping %", min_value=0.0, value=30.0, step=5.0)
            a_unit = f4.number_input("Unit Cost (USD)", min_value=0.0, value=0.0, step=1.0,
                                     help="If Ex Unit Cost is entered, this is calculated (in USD) from Shipping %.")
            a_up = f5.number_input("Default U.Price $", min_value=0.0, value=0.0, step=1.0)
            a_ups = f6.number_input("Default U.Price SAR", min_value=0.0, value=0.0, step=10.0)
            if st.form_submit_button("âž• Add item", type="primary"):
                if not (a_model.strip() or a_desc.strip()):
                    st.warning("Enter at least a Model or a Description.")
                else:
                    iid = repo.add_catalog_item(
                        a_brand.strip(), a_model.strip(), a_desc.strip(),
                        a_list or None, a_ex or None, a_ship, a_unit or None, a_up or None,
                        a_ups or None, currency=a_cur)
                    if iid:
                        st.success(f"Added '{a_model or a_desc}' (ItemID {iid}).")
                        st.rerun()
                    else:
                        st.warning("An item with the same Brand + Model + Description already exists.")

    st.caption("Prices shown **rounded up**. Edit cost / default-price cells inline, or tick "
               "**Del** to remove items. Brand / Model / Description are read-only here â€” "
               "use **Add new item** above to create one.")
    term = st.text_input("Search", placeholder="Model / Description / Brand")
    res = repo.search_catalog(term, limit=300).reset_index(drop=True)
    st.caption(f"{len(res)} item(s)")
    if not res.empty:
        rename = {"ListPriceUSD": "List Price $", "ExUnitCostUSD": "Ex Unit Cost $",
                  "Currency": "Cur",
                  "ShippingPercent": "Shipping %", "UnitCostUSD": "Unit Cost $", "DefaultUPriceUSD": "Default U.Price $",
                  "DefaultUPriceSAR": "Default U.Price SAR", "TimesQuoted": "Times Quoted"}
        edit_cols = list(repo.CATALOG_EDITABLE.keys())
        money_cols = [c for c in edit_cols if c != "Shipping %"]
        disp = res.rename(columns=rename).copy()
        if "Cur" not in disp.columns:
            disp["Cur"] = "USD"
        disp["Cur"] = disp["Cur"].fillna("USD")
        for c in money_cols:                       # always display rounded up
            disp[c] = disp[c].map(lambda v: calc.roundup(v, 0))
        if "Shipping %" in disp.columns:
            disp["Shipping %"] = disp["Shipping %"].map(lambda v: calc.shipping_percent(v))
        base_cols = ["Brand", "Model", "Description", "Cur"] + edit_cols + ["Times Quoted"]
        colcfg = {c: st.column_config.NumberColumn(c, format="%.0f", min_value=0.0)
                  for c in money_cols}
        colcfg["Cur"] = st.column_config.SelectboxColumn(
            "Cur", options=calc.CURRENCIES, required=False,
            help="Currency of List Price & Ex Unit Cost. Unit Cost (USD) recomputes when changed.")
        colcfg["List Price $"] = st.column_config.NumberColumn("List Price", format="%.0f", min_value=0.0)
        colcfg["Ex Unit Cost $"] = st.column_config.NumberColumn("Ex Unit Cost", format="%.0f", min_value=0.0)
        colcfg["Unit Cost $"] = st.column_config.NumberColumn("Unit Cost (USD)", format="%.0f", min_value=0.0)
        colcfg["Shipping %"] = st.column_config.NumberColumn("Shipping %", format="%.2f", min_value=0.0, step=5.0)
        colcfg["Times Quoted"] = st.column_config.NumberColumn("Times Quoted", format="%d")
        if not _cat_edit:                              # read-only catalogue
            st.dataframe(disp[base_cols], use_container_width=True, hide_index=True,
                         column_config=colcfg)
        else:
            disp["Del"] = False
            colcfg["Del"] = st.column_config.CheckboxColumn("Del", help="Tick to delete this item")
            edited = st.data_editor(
                disp[["Del"] + base_cols], column_config=colcfg, num_rows="fixed", hide_index=True,
                use_container_width=True, key=f"cat_editor::{term}",
                disabled=["Brand", "Model", "Description", "Times Quoted"])
            del_ids = [int(res.iloc[i]["ItemID"]) for i in range(len(edited))
                       if bool(edited.iloc[i]["Del"])]
            b1, b2 = st.columns(2)
            if b1.button("ðŸ’¾ Save price changes", type="primary", use_container_width=True):
                n = 0
                for i in range(len(edited)):
                    changes = {}
                    for c in edit_cols:
                        new = edited.iloc[i][c]
                        if pd.isna(new):
                            continue
                        if abs(float(new) - float(disp.iloc[i][c])) > 1e-9:
                            changes[repo.CATALOG_EDITABLE[c]] = float(new)
                    new_cur = str(edited.iloc[i]["Cur"])
                    if new_cur in calc.CURRENCIES and new_cur != str(disp.iloc[i]["Cur"]):
                        changes["Currency"] = new_cur
                    if changes:
                        repo.update_catalog_item(int(res.iloc[i]["ItemID"]), changes)
                        n += 1
                st.success(f"Updated {n} catalogue item(s).")
                st.rerun()
            if b2.button(f"ðŸ-‘ï¸ Delete {len(del_ids)} checked item(s)", use_container_width=True,
                         disabled=not del_ids):
                n = repo.delete_catalog_items(del_ids)
                st.success(f"Deleted {n} item(s).")
                st.rerun()


# ============================ SETTINGS ============================
elif mode == "Settings":
    st.subheader("Settings")
    st.caption("Offer reference numbers are built from a **template** with variables.")

    with st.form("settings_form"):
        template = st.text_input("Offer # template", repo.get_setting("offer_template"),
                                 help="Variables: *TYPE* = System Offer (AV, LCSâ€¦), *YY* = 2-digit "
                                      "year, *YYYY* = 4-digit year, and a run of x's = the auto-number "
                                      "(its length is the zero-padding). e.g. LG-*TYPE*-*YY*/xxxx â†’ "
                                      "LG-AV-26/0053. Omit *TYPE* for a fixed prefix (e.g. SWS-*YY*-xxxx).")
        c1, c2, c3 = st.columns(3)
        pad = c1.number_input("Fallback digits (no x-run)", min_value=1, max_value=8,
                              value=int(repo.get_setting("offer_number_pad") or 3),
                              help="Padding used only when the template has no x's.")
        types = c2.text_input("System Offer types (comma-separated)", repo.get_setting("offer_types"),
                              help="Shown as the 'System Offer' dropdown on a new offer. "
                                   "Pick '(none)' there to skip the *TYPE* segment.")
        dmargin = c3.number_input("Default margin Ã-", min_value=0.0, step=0.05,
                                  value=float(repo.get_setting("default_margin") or 1.6),
                                  help="Applied to new blank rows and catalogue items with no "
                                       "historical price, in New Project and Edit.")

        st.markdown("**Revision label**")
        rcol1, rcol2 = st.columns(2)
        rev_fmt = rcol1.text_input("Revision format", repo.get_setting("revision_format"),
                                   help="A run of x's = the revision number (length = padding). "
                                        "e.g. Rev.x â†’ Rev.1 / Rev.10 ;  Rxx â†’ R01 / R10.")
        sep_opts = {"Dash   (â€¦0053-Rev.1)": "-", "Space   (â€¦0053 Rev.1)": " ",
                    "Underscore   (â€¦0053_Rev.1)": "_"}
        _cur_sep = repo.get_setting("revision_separator")
        _cur_lbl = next((k for k, v in sep_opts.items() if v == _cur_sep), list(sep_opts)[0])
        rev_sep_lbl = rcol2.selectbox("Separator (offer # â†’ revision)", list(sep_opts.keys()),
                                      index=list(sep_opts.keys()).index(_cur_lbl))

        st.markdown("**Tax**")
        vcol1, vcol2 = st.columns([1, 2])
        vat_pct = vcol1.number_input("VAT %", min_value=0.0, max_value=100.0, step=0.5,
                                     value=float(repo.get_setting("vat_percent") or 15),
                                     help="VAT rate applied across offers, quotations and the Finance tab. "
                                          "KSA = 15%; change it for other countries.")
        vcol2.caption("Applies everywhere VAT is shown â€” new offers, loaded offers, the client PDF "
                      "and the Finance tab. Changing it re-computes VAT on all offers.")

        st.markdown("**Currencies / exchange rates**")
        ecol1, ecol2 = st.columns([1, 2])
        eur_rate = ecol1.number_input("1 EUR = ? USD", min_value=0.0, step=0.01, format="%.4f",
                                      value=float(repo.get_setting("eur_to_usd") or 1.08),
                                      help="Converts EUR buy prices to USD when computing the Unit Cost.")
        ecol2.caption(f"Pegged (fixed, not editable): 1 USD = {calc.SAR_PER_USD:g} SAR "
                      f"(1 SAR â‰ˆ {1 / calc.SAR_PER_USD:.4f} USD)  Â·  "
                      f"1 USD = {calc.AED_PER_USD:g} AED (1 AED â‰ˆ {1 / calc.AED_PER_USD:.4f} USD).")

        st.markdown("**Company / Branding**")
        gc1, gc2 = st.columns([3, 1])
        company_name = gc1.text_input("Company name", repo.get_setting("company_name") or "",
                                      help="Shown in the page title, client PDF and project sheet.")
        brand_color = gc2.color_picker("Brand color", repo.get_setting("company_brand_color") or "#002060",
                                       help="Primary colour for PDF titles, table headers and footer.")
        company_tagline = st.text_input("Tagline", repo.get_setting("company_tagline") or "")
        company_contact = st.text_input("Contact line (city / country)",
                                        repo.get_setting("company_contact") or "")

        saved = st.form_submit_button("ðŸ’¾ Save settings", type="primary")
        if saved:
            repo.set_setting("offer_template", template.strip())
            repo.set_setting("offer_number_pad", int(pad))
            repo.set_setting("offer_types", types.strip())
            repo.set_setting("default_margin", float(dmargin))
            repo.set_setting("revision_format", rev_fmt.strip() or "Rev.x")
            repo.set_setting("revision_separator", sep_opts[rev_sep_lbl])
            repo.set_setting("eur_to_usd", float(eur_rate))
            repo.set_setting("vat_percent", float(vat_pct))
            repo.set_setting("company_name", company_name.strip() or "Company")
            repo.set_setting("company_tagline", company_tagline.strip())
            repo.set_setting("company_contact", company_contact.strip())
            repo.set_setting("company_brand_color", brand_color)
            st.success("Settings saved. (Page title updates on next reload.)")

    # ---- Branding images (per company; outside the form for file upload) ----
    st.divider()
    st.markdown("##### Branding images")
    bcol, lcol = st.columns([2, 1])
    with bcol:
        st.markdown("**Banner** â€” full-width header (app, PDF, project sheet)")
        if os.path.exists(db.banner_path()):
            st.image(db.banner_path(), use_container_width=True)
        else:
            st.info("No banner yet.")
        up_b = st.file_uploader("Upload / replace banner (PNG)", type=["png"], key="banner_up")
        if up_b is not None and st.button("ðŸ’¾ Save banner", key="save_banner"):
            with open(db.banner_path(), "wb") as f:
                f.write(up_b.getbuffer())
            st.success("Banner updated. (Reload to see it in the header/sidebar.)")
            st.rerun()
    with lcol:
        st.markdown("**Logo** â€” standalone mark")
        if os.path.exists(db.logo_path()):
            st.image(db.logo_path(), width=160)
        else:
            st.info("No logo yet.")
        up_l = st.file_uploader("Upload / replace logo (PNG)", type=["png"], key="logo_up")
        if up_l is not None and st.button("ðŸ’¾ Save logo", key="save_logo"):
            with open(db.logo_path(), "wb") as f:
                f.write(up_l.getbuffer())
            st.success("Logo updated.")
            st.rerun()
    st.caption("Banner â‰ˆ 1400Ã-155 px. Logo: a square / transparent PNG works best.")

    st.divider()
    st.markdown("##### Offer-number preview")
    st.caption("Numbering is **per series** â€” each rendered template (type + year) keeps its own "
               "counter, so LG-AV-26/â€¦ and LG-LC-26/â€¦ don't conflict.")
    ex_type = (repo.offer_types() or ["AV"])[0]
    _ex = repo.make_offer_no(ex_type)
    st.write("Next offer # examples:")
    st.code(f"no type      :  {repo.make_offer_no('')}\n"
            f"type {ex_type:<6}  :  {_ex}\n"
            f"revision     :  {_ex}{repo.revision_separator()}{repo.revision_token(1)}"
            f"   /   {_ex}{repo.revision_separator()}{repo.revision_token(2)}")

    st.divider()
    st.markdown("##### Reset / force a starting number")
    st.caption("Force a series to begin at a chosen number; numbering then continues "
               "incrementing from there. (Numbers at or below the current next have no effect "
               "â€” it never reuses an existing number.)")
    rc1, rc2, rc3 = st.columns([2, 1, 1])
    rsel = rc1.selectbox("Series (System Offer)", ["(none)"] + repo.offer_types(), key="reset_series")
    r_otype = "" if rsel == "(none)" else rsel
    r_next = repo.next_offer_number(r_otype)
    r_floor = repo.get_series_start(repo.series_key(r_otype))
    rc1.caption(f"Next: `{repo.make_offer_no(r_otype)}`"
                + (f" Â· forced start: {r_floor}" if r_floor else ""))
    start_at = rc2.number_input("Start at", min_value=1, value=max(int(r_next), 1), step=1,
                                key="reset_start_val")
    if rc3.button("âœ… Apply", use_container_width=True):
        repo.set_series_start(r_otype, int(start_at))
        st.success(f"This series will start at {int(start_at)}, then continue incrementing.")
        st.rerun()
    if r_floor and rc3.button("â†©ï¸ Clear", use_container_width=True):
        repo.clear_series_start(r_otype)
        st.success("Forced start cleared â€” back to automatic numbering.")
        st.rerun()


# ============================ USERS (owner) ============================
elif mode == "Users":
    st.subheader("Users & access")
    tab_users, tab_roles = st.tabs(["ðŸ‘¤ Users", "ðŸ›¡ï¸ Roles & permissions"])

    # ---------------------------- USERS TAB ----------------------------
    with tab_users:
        roles = auth.list_roles()
        _default_role = "viewer" if "viewer" in roles else roles[-1]
        with st.expander("âž• Add user", expanded=False):
            with st.form("add_user", clear_on_submit=True):
                uc1, uc2 = st.columns(2)
                nu = uc1.text_input("Username")
                ndn = uc2.text_input("Display name")
                uc3, uc4, uc5 = st.columns(3)
                np1 = uc3.text_input("Password", type="password")
                np2 = uc4.text_input("Confirm", type="password")
                nrole = uc5.selectbox("Role", roles, index=roles.index(_default_role))
                if st.form_submit_button("âž• Create user", type="primary"):
                    if not nu.strip() or not np1:
                        st.warning("Username and password are required.")
                    elif np1 != np2:
                        st.warning("Passwords don't match.")
                    elif auth.create_user(nu, np1, ndn, nrole):
                        st.success(f"Created user '{nu}' ({nrole}).")
                        st.rerun()
                    else:
                        st.warning("Username already exists.")

        st.markdown("##### Existing users")
        users = auth.list_users()
        owners = sum(1 for x in users if x["Role"] == auth.PROTECTED_ROLE and x["Active"])
        for u in users:
            uid = u["UserID"]
            is_self = (uid == USER.get("UserID"))
            last_owner = (u["Role"] == auth.PROTECTED_ROLE and owners <= 1)
            with st.container(border=True):
                cc = st.columns([2, 2, 1.5, 1, 1.2], vertical_alignment="center")
                cc[0].markdown(f"**{u['Username']}**" + (" Â· _(you)_" if is_self else ""))
                cc[1].write(u.get("DisplayName") or "")
                r_idx = roles.index(u["Role"]) if u["Role"] in roles else 0
                new_role = cc[2].selectbox("Role", roles, key=f"role_{uid}",
                                           index=r_idx, label_visibility="collapsed")
                active = cc[3].toggle("On", value=bool(u["Active"]), key=f"act_{uid}",
                                      label_visibility="collapsed")
                if (new_role != u["Role"] or active != bool(u["Active"])):
                    if cc[4].button("ðŸ’¾ Save", key=f"saveu_{uid}", use_container_width=True):
                        if last_owner and (new_role != auth.PROTECTED_ROLE or not active):
                            st.warning("Can't remove the last active owner.")
                        else:
                            auth.update_user(uid, role=new_role, active=active)
                            st.toast(f"Updated {u['Username']}.")
                            st.rerun()
                pc = st.columns([3, 1.3, 1.3])
                newpw = pc[0].text_input("pw", type="password", key=f"pw_{uid}",
                                         label_visibility="collapsed",
                                         placeholder="New password (blank = keep)")
                if pc[1].button("ðŸ”‘ Set password", key=f"setpw_{uid}", use_container_width=True,
                                disabled=not newpw):
                    auth.set_password(uid, newpw)
                    st.toast(f"Password updated for {u['Username']}.")
                    st.rerun()
                if pc[2].button("ðŸ-‘ï¸ Delete", key=f"delu_{uid}", use_container_width=True,
                                disabled=is_self or last_owner):
                    auth.delete_user(uid)
                    st.toast(f"Deleted {u['Username']}.")
                    st.rerun()

    # ---------------------- ROLES & PERMISSIONS TAB --------------------
    with tab_roles:
        st.caption("Tick what each role is allowed to do, then **Save**. "
                   "The **owner** always has full access and can't be changed.")
        roles = auth.list_roles()
        rows = []
        for role in roles:
            perms = auth.role_perms(role)
            row = {"Role": role}
            for k, label in auth.PERMISSIONS:
                row[label] = (k in perms)
            rows.append(row)
        mat = pd.DataFrame(rows, columns=["Role"] + [lbl for _, lbl in auth.PERMISSIONS])

        colcfg = {"Role": st.column_config.TextColumn("Role", disabled=True, width="small")}
        for _, label in auth.PERMISSIONS:
            colcfg[label] = st.column_config.CheckboxColumn(label)
        edited = st.data_editor(mat, column_config=colcfg, hide_index=True,
                                num_rows="fixed", use_container_width=True,
                                disabled=["Role"], key="role_matrix")

        if st.button("ðŸ’¾ Save role permissions", type="primary"):
            for _, r in edited.iterrows():
                role = r["Role"]
                if role == auth.PROTECTED_ROLE:
                    continue
                granted = {auth.LABEL_TO_PERM[lbl] for _, lbl in auth.PERMISSIONS if bool(r[lbl])}
                auth.set_role_perms(role, granted)
            st.success("Role permissions saved.")
            st.rerun()

        st.divider()
        rc1, rc2 = st.columns(2)
        with rc1:
            st.markdown("**Add a role**")
            arc1, arc2 = st.columns([3, 1.4], vertical_alignment="bottom")
            new_role_name = arc1.text_input("New role name", key="new_role_name",
                                            label_visibility="collapsed",
                                            placeholder="e.g. Estimator")
            if arc2.button("âž• Add", use_container_width=True):
                if auth.add_role(new_role_name):
                    st.toast(f"Added role '{new_role_name.strip()}'.")
                    st.rerun()
                else:
                    st.warning("Empty or duplicate role name.")
        with rc2:
            st.markdown("**Delete a role**")
            deletable = [r for r in roles if r != auth.PROTECTED_ROLE]
            drc1, drc2 = st.columns([3, 1.4], vertical_alignment="bottom")
            drole = drc1.selectbox("Role to delete", deletable, key="del_role",
                                   label_visibility="collapsed") if deletable else None
            if drc2.button("ðŸ-‘ï¸ Delete", use_container_width=True, disabled=not deletable):
                in_use = auth.role_user_count(drole)
                if in_use:
                    st.warning(f"{in_use} user(s) still have the '{drole}' role â€” "
                               "reassign them first.")
                elif auth.delete_role(drole):
                    st.toast(f"Deleted role '{drole}'.")
                    st.rerun()
