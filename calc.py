"""
Pricing calculation engine - the single source of truth for every grid formula.

Mirrors the Excel sheet exactly (validated against the historical workbooks):
    Unit Cost  = Ex Unit Cost x (1 + Shipping % / 100)  (manual if no Ex cost)
    U. Price $ = ROUNDUP(Unit Cost x Margin, 0)   (Margin = per-line multiplier)
    U. Price SAR = ROUNDUP(U. Price $ x 3.75 -> next multiple of 10)
    Total Cost = Qty x Unit Cost
    T. Price   = Qty x U. Price
    T. Price SAR = Qty x U. Price SAR
"""
from __future__ import annotations
import math
import pandas as pd

SAR_PER_USD = 3.75      # USD->SAR peg (confirmed: 99.9% of historical rows)
AED_PER_USD = 3.6725    # USD->AED peg (fixed since 1997)
DEFAULT_SHIPPING_PERCENT = 30.0
COST_BUFFER = 1 + DEFAULT_SHIPPING_PERCENT / 100  # legacy alias for older code/comments
VAT_RATE = 0.15         # 15% KSA VAT
USD_ROUND_DEC = 0       # U. Price $ rounds UP to a whole dollar
SAR_ROUND_TO = 10       # U. Price SAR rounds UP to the next multiple of 10

# Currencies a buy price (List Price / Ex Unit Cost) can be entered in.
# CURRENCY_RATES = USD value of 1 unit of the currency; the app refreshes the
# EUR rate from Settings at runtime. SAR is the fixed peg (1 SAR = 1/3.75 USD).
CURRENCIES = ["USD", "EUR", "SAR", "AED"]
CURRENCY_COL = "Cur"
DEFAULT_CURRENCY = "USD"
CURRENCY_RATES = {"USD": 1.0, "EUR": 1.08, "SAR": 1.0 / SAR_PER_USD, "AED": 1.0 / AED_PER_USD}

# Canonical grid columns, left-to-right. "Markup x" drives the selling price.
GRID_COLUMNS = [
    "Area", "System", "Description", "Brand", "Model", "Qty",
    "Cur", "List Price $", "Ex Unit Cost $", "Shipping %", "Unit Cost $", "Total Cost $",
    "Markup x", "U. Price $", "T. Price $", "U. Price SAR", "T. Price SAR",
]
TEXT_COLUMNS = {"Area", "System", "Description", "Brand", "Model", "Cur"}

# Columns hidden from the client-facing Quotation/PDF (internal cost metrics).
COST_COLUMNS = ["Cur", "List Price $", "Ex Unit Cost $", "Shipping %", "Unit Cost $",
                "Total Cost $", "Markup x"]


def currency_rate(currency) -> float:
    """USD value of 1 unit of `currency` (defaults to USD = 1.0)."""
    return CURRENCY_RATES.get(str(currency).strip() or DEFAULT_CURRENCY, 1.0)


def to_usd(amount, currency=DEFAULT_CURRENCY) -> float:
    """Convert a buy amount expressed in `currency` to USD."""
    return _num(amount) * currency_rate(currency)


def _num(x):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)) or x == "":
            return 0.0
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def _blank(x) -> bool:
    try:
        return x is None or x == "" or pd.isna(x)
    except (TypeError, ValueError):
        return False


def roundup(x, decimals=0):
    """Excel ROUNDUP(x, decimals) - always rounds away from zero (ceiling for +)."""
    x = _num(x)
    if x == 0:
        return 0.0
    f = 10 ** decimals
    return math.ceil(round(x * f, 6)) / f


def roundup_to(x, step):
    """Round UP to the next multiple of `step` (Excel CEILING(x, step))."""
    x = _num(x)
    if x == 0:
        return 0.0
    return math.ceil(round(x / step, 6)) * step


def shipping_percent(value, ex_unit_cost=None, unit_cost=None) -> float:
    """Shipping percentage as a user-facing percent number, defaulting to 30."""
    if not _blank(value):
        return max(_num(value), 0.0)
    return infer_shipping_percent(ex_unit_cost, unit_cost)


def infer_shipping_percent(ex_unit_cost, unit_cost=None) -> float:
    """Infer Shipping % from stored costs, else return the default 30%."""
    ex = _num(ex_unit_cost)
    unit = _num(unit_cost)
    if ex > 0 and unit > 0:
        return round(max((unit / ex - 1) * 100, 0.0), 2)
    return DEFAULT_SHIPPING_PERCENT


def unit_cost_from_ex(ex_unit_cost, shipping_pct=None) -> float:
    """Unit Cost = ROUNDUP(Ex Unit Cost x (1 + Shipping % / 100))."""
    ship = shipping_percent(shipping_pct, ex_unit_cost)
    return roundup(_num(ex_unit_cost) * (1 + ship / 100), 0)


def u_price_from_margin(unit_cost, margin) -> float:
    """U. Price $ = ROUNDUP(Unit Cost x Margin, 0)."""
    return roundup(_num(unit_cost) * _num(margin), USD_ROUND_DEC)


def u_price_sar_from_usd(u_price_usd) -> float:
    """U. Price SAR = ROUNDUP(U. Price $ x 3.75 -> next 10)."""
    return roundup_to(_num(u_price_usd) * SAR_PER_USD, SAR_ROUND_TO)


def recompute(df: pd.DataFrame) -> pd.DataFrame:
    """Recalculate every derived column following the exact Excel chain.

    Drivers (editable):  Ex Unit Cost $, Shipping %, Markup x, and Qty.
                         Unit Cost $ stays manual when no Ex Unit Cost exists.
                         A manual U. Price $ is honoured only when
                         Markup x is blank/0 (lets you price odd items directly).
    Discount rows (LineType == 'discount') are passed through untouched.
    """
    df = df.copy()
    for col in GRID_COLUMNS:
        if col not in df.columns:
            df[col] = "" if col in TEXT_COLUMNS else 0.0
    if "LineType" not in df.columns:
        df["LineType"] = "item"
    # Normalise the per-line currency (blank / unknown -> USD).
    df[CURRENCY_COL] = [c if str(c).strip() in CURRENCIES else DEFAULT_CURRENCY
                        for c in df[CURRENCY_COL]]

    ship_l, unit_l, tc_l, up_l, tp_l, usar_l, tpsar_l = [], [], [], [], [], [], []
    for _, r in df.iterrows():
        if str(r.get("LineType", "item")) == "discount":
            ship_l.append(shipping_percent(r.get("Shipping %")))
            unit_l.append(_num(r.get("Unit Cost $")))
            tc_l.append(_num(r.get("Total Cost $")))
            up_l.append(_num(r.get("U. Price $")))
            tp_l.append(_num(r.get("T. Price $")))
            usar_l.append(_num(r.get("U. Price SAR")))
            tpsar_l.append(_num(r.get("T. Price SAR")))
            continue

        qty = _num(r.get("Qty"))
        ex = _num(r.get("Ex Unit Cost $"))
        ex_usd = to_usd(ex, r.get(CURRENCY_COL))     # Ex cost may be in EUR / SAR
        ship = shipping_percent(r.get("Shipping %"), ex_usd, r.get("Unit Cost $"))
        # Unit Cost rounded UP to a whole dollar (landed cost, always USD).
        unit = roundup(ex_usd * (1 + ship / 100), 0) if ex > 0 else roundup(_num(r.get("Unit Cost $")), 0)

        margin = round(_num(r.get("Markup x")), 4)
        if margin > 0:
            uprice = u_price_from_margin(unit, margin)        # formula-driven
        else:
            uprice = roundup(_num(r.get("U. Price $")), 0)    # manual / from catalogue

        if uprice > 0:
            usar = u_price_sar_from_usd(uprice)          # normal: SAR derived from the dollar price
        elif unit > 0 or margin > 0:
            usar = 0.0                                   # priced at 0 against a real cost/margin -> SAR follows to 0
        else:
            usar = roundup_to(_num(r.get("U. Price SAR")), SAR_ROUND_TO)  # SAR-only line (no USD basis) -> keep it

        # Every monetary value is rounded UP to a whole number.
        ship_l.append(ship)
        unit_l.append(unit)
        tc_l.append(roundup(qty * unit, 0))
        up_l.append(uprice)
        tp_l.append(roundup(qty * uprice, 0))
        usar_l.append(usar)
        tpsar_l.append(roundup(qty * usar, 0))

    df["Shipping %"] = ship_l
    # Margin is a pricing driver with at most four decimal places. Keeping the
    # numeric values rounded lets the UI display only meaningful digits (1.6,
    # 1.25, 1.2345) instead of padding every cell with trailing zeroes.
    df["Markup x"] = [round(_num(value), 4) for value in df["Markup x"]]
    df["Unit Cost $"], df["Total Cost $"] = unit_l, tc_l
    df["U. Price $"], df["T. Price $"] = up_l, tp_l
    df["U. Price SAR"], df["T. Price SAR"] = usar_l, tpsar_l
    return df


def effective_margin(df: pd.DataFrame) -> pd.Series:
    """Read-only effective margin = U. Price $ / Unit Cost $ (for display)."""
    uc = df["Unit Cost $"].map(_num)
    up = df["U. Price $"].map(_num)
    return (up / uc.where(uc > 0)).round(4)


def increase_margins(df: pd.DataFrame, percentage: float) -> tuple[pd.DataFrame, int]:
    """Adjust every positive item margin by a percentage and recalculate prices."""
    grid = recompute(df)
    percentage = max(_num(percentage), -100.0)
    factor = 1 + percentage / 100
    line_types = grid.get("LineType", pd.Series("item", index=grid.index))
    margins = grid["Markup x"].map(_num)
    mask = line_types.astype(str).str.lower().ne("discount") & margins.gt(0)
    grid.loc[mask, "Markup x"] = (margins.loc[mask] * factor).round(4)
    return recompute(grid), int(mask.sum())


def summarize(df: pd.DataFrame, discount_sar: float = 0.0,
              commission_sar: float = 0.0) -> dict:
    """Client totals plus internal commission-adjusted profit metrics."""
    if "LineType" in df:
        lt = df["LineType"].astype(str)
        cost_items = df[lt != "discount"]                    # cost includes included rows
        sell_items = df[~lt.isin({"discount", "included"})] # selling excludes included rows
    else:
        cost_items = sell_items = df
    product_cost_usd = cost_items["Total Cost $"].map(_num).sum()
    total_sell_usd = sell_items["T. Price $"].map(_num).sum()
    subtotal_sar = sell_items["T. Price SAR"].map(_num).sum()

    discount_sar = min(abs(_num(discount_sar)), subtotal_sar)
    discount_percent = (discount_sar / subtotal_sar * 100) if subtotal_sar else 0.0
    discounted = subtotal_sar - discount_sar
    commission_sar = max(_num(commission_sar), 0.0)
    commission_percent = (commission_sar / discounted * 100) if discounted else 0.0
    # Commission is an internal expense. It is never included in client-facing
    # subtotal, VAT, grand total, or exported quotation totals.
    vat = round(discounted * VAT_RATE, 2)
    grand_total = round(discounted + vat, 2)

    product_cost_sar = product_cost_usd * SAR_PER_USD
    cost_sar = product_cost_sar + commission_sar
    total_cost_usd = product_cost_usd + commission_sar / SAR_PER_USD
    markup_factor = (discounted / cost_sar) if cost_sar else None

    return {
        "total_cost_usd": round(total_cost_usd, 2),
        "product_cost_usd": round(product_cost_usd, 2),
        "total_sell_usd": round(total_sell_usd, 2),
        "subtotal_sar": round(subtotal_sar, 2),
        "discount_sar": round(discount_sar, 2),
        "discount_percent": round(discount_percent, 4),
        "discounted_subtotal_sar": round(discounted, 2),
        "commission_sar": round(commission_sar, 2),
        "commission_percent": round(commission_percent, 4),
        "vat_rate": VAT_RATE,
        "vat_amount_sar": vat,
        "grand_total_sar": grand_total,
        "product_cost_sar": round(product_cost_sar, 2),
        "cost_sar": round(cost_sar, 2),
        "markup_factor": round(markup_factor, 4) if markup_factor else None,
        # Commission is internal cost, so it reduces markup and profit.
        "gross_margin_sar": round(discounted - cost_sar, 2),
        "gross_margin_usd": round((discounted - cost_sar) / SAR_PER_USD, 2),
    }


def blank_row(area="", system="", line_type="item") -> dict:
    row = {c: ("" if c in TEXT_COLUMNS else 0.0) for c in GRID_COLUMNS}
    row["Area"], row["System"] = area, system
    row["Cur"] = DEFAULT_CURRENCY
    row["Shipping %"] = DEFAULT_SHIPPING_PERCENT
    row["LineType"] = line_type
    row["_IncludedInItems"] = False
    row["_RowOrder"] = 0
    return row


def apply_inclusion(df: pd.DataFrame, markup: float = 1.0) -> pd.DataFrame:
    """Zero selling prices for included rows and mark them as LineType='included'."""
    df = df.copy()
    if "_IncludedInItems" not in df.columns:
        return df
    included_mask = df["_IncludedInItems"].fillna(False).astype(bool)
    df.loc[included_mask, "LineType"] = "included"
    for col in ["Markup x", "U. Price $", "T. Price $", "U. Price SAR", "T. Price SAR"]:
        df.loc[included_mask, col] = 0.0
    return df
