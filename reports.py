"""
Reporting engine: denormalized datasets + filtering/aggregation for the
Reports workspace (Report Builder + Dashboard). All money is SAR unless noted.

Three datasets, each one flat DataFrame ready for filtering / group-by:
  - "Offers / Projects" : one row per offer (meta + totals + profit).
  - "Line items"        : one row per BoQ item line (meta + line metrics).
  - "Finance"           : one row per offer (collected / due / PO spend / net).
"""
from __future__ import annotations
import datetime as dt
import pandas as pd

import db
import calc


def _conn():
    return db.connect()


def _month(series: pd.Series) -> pd.Series:
    d = pd.to_datetime(series, errors="coerce")
    return d.dt.strftime("%Y-%m")


# ----------------------------- datasets -----------------------------

def offers_df(include_archived: bool = False) -> pd.DataFrame:
    """One row per offer/revision/option with totals and profit."""
    sql = """
        SELECT p.ProjectID, p.ProjectName, p.ClientName, p.SalesPerson, p.PresalesEngineer,
               p.ProjectManager, p.OfferNo, p.CreationDate, p.RevisionNo, p.OptionLabel,
               IFNULL(p.Approved,0) AS Approved, IFNULL(p.Archived,0) AS Archived,
               IFNULL(p.DiscountAmount,0) AS DiscountAmount,
               (SELECT s.SystemSuffix FROM Project_Sheets s
                  WHERE s.ProjectID = p.ProjectID LIMIT 1) AS System,
               (SELECT IFNULL(SUM(l.TPriceSAR),0) FROM Project_BoQ_Lines l
                  WHERE l.ProjectID = p.ProjectID AND l.LineType='item') AS SubtotalSAR,
               (SELECT IFNULL(SUM(l.TotalCostUSD),0) FROM Project_BoQ_Lines l
                  WHERE l.ProjectID = p.ProjectID AND l.LineType='item') AS TotalCostUSD
        FROM Projects_Master p
    """
    with _conn() as c:
        df = pd.DataFrame([dict(r) for r in c.execute(sql)])
    if df.empty:
        return df
    if not include_archived:
        df = df[df["Archived"] == 0]
    sub = df["SubtotalSAR"].astype(float)
    disc = df["DiscountAmount"].abs().clip(upper=sub)
    discounted = sub - disc
    vat = discounted * calc.VAT_RATE
    cost_sar = df["TotalCostUSD"].astype(float) * calc.SAR_PER_USD
    df["Subtotal SAR"] = sub.round(0)
    df["Discount SAR"] = disc.round(0)
    df["Grand Total SAR"] = (discounted + vat).round(0)
    df["Cost SAR"] = cost_sar.round(0)
    df["Gross Profit SAR"] = (discounted - cost_sar).round(0)
    df["Margin %"] = (df["Gross Profit SAR"] / discounted.where(discounted > 0) * 100).round(1).fillna(0)
    df["Status"] = df["Approved"].map(lambda v: "Approved" if v else "Pending")
    df["Month"] = _month(df["CreationDate"])
    df["System"] = df["System"].fillna("")
    df = df.rename(columns={"ClientName": "Client", "SalesPerson": "Sales Person",
                            "PresalesEngineer": "Pre-sales", "ProjectManager": "Project Mgr",
                            "ProjectName": "Project", "OfferNo": "Offer #", "CreationDate": "Date"})
    return df


def lines_df(include_archived: bool = False) -> pd.DataFrame:
    """One row per BoQ item line with its project meta and line metrics."""
    sql = """
        SELECT p.ProjectName, p.ClientName, p.SalesPerson, p.PresalesEngineer, p.ProjectManager,
               p.OfferNo, p.CreationDate, IFNULL(p.Approved,0) AS Approved,
               IFNULL(p.Archived,0) AS Archived,
               l.System, l.Area, l.Description, l.Brand, l.Model, IFNULL(l.Qty,0) AS Qty,
               IFNULL(l.TotalCostUSD,0) AS TotalCostUSD, IFNULL(l.TPriceSAR,0) AS TPriceSAR,
               l.Currency
        FROM Project_BoQ_Lines l
        JOIN Projects_Master p ON p.ProjectID = l.ProjectID
        WHERE l.LineType = 'item'
    """
    with _conn() as c:
        df = pd.DataFrame([dict(r) for r in c.execute(sql)])
    if df.empty:
        return df
    if not include_archived:
        df = df[df["Archived"] == 0]
    cost_sar = df["TotalCostUSD"].astype(float) * calc.SAR_PER_USD
    df["Cost SAR"] = cost_sar.round(0)
    df["Total Price SAR"] = df["TPriceSAR"].round(0)
    df["Profit SAR"] = (df["TPriceSAR"] - cost_sar).round(0)
    df["Status"] = df["Approved"].map(lambda v: "Approved" if v else "Pending")
    df["Month"] = _month(df["CreationDate"])
    for col in ("System", "Area", "Brand", "Model", "Currency"):
        df[col] = df[col].fillna("")
    df = df.rename(columns={"ClientName": "Client", "SalesPerson": "Sales Person",
                            "PresalesEngineer": "Pre-sales", "ProjectManager": "Project Mgr",
                            "ProjectName": "Project", "OfferNo": "Offer #", "CreationDate": "Date"})
    return df


def finance_df(include_archived: bool = False) -> pd.DataFrame:
    """One row per offer with collected / due / PO spend / net profit."""
    base = offers_df(include_archived=include_archived)
    if base.empty:
        return base
    # Map ProjectID -> total (works even when the finance tables are empty).
    with _conn() as c:
        pay = {r["ProjectID"]: r["Collected"] for r in c.execute(
            "SELECT ProjectID, IFNULL(SUM(AmountSAR),0) AS Collected "
            "FROM Finance_Payments GROUP BY ProjectID")}
        pur = {r["ProjectID"]: r["PO_Spend"] for r in c.execute(
            "SELECT ProjectID, IFNULL(SUM(AmountSAR),0) AS PO_Spend "
            "FROM Finance_Purchases GROUP BY ProjectID")}
    df = base.copy()
    df["Collected SAR"] = df["ProjectID"].map(pay).fillna(0).round(0)
    df["PO Spend SAR"] = df["ProjectID"].map(pur).fillna(0).round(0)
    gt = df["Grand Total SAR"]
    df["Remaining SAR"] = (gt - df["Collected SAR"]).round(0)
    df["Net Profit SAR"] = (gt - df["PO Spend SAR"] - gt * calc.VAT_RATE).round(0)
    return df


# Per-dataset metadata: which columns are categorical filters / dates / metrics / shown.
DATASETS = {
    "Offers / Projects": {
        "builder": offers_df,
        "filters": ["Client", "Sales Person", "Pre-sales", "Project Mgr", "System", "Status"],
        "date": "Date",
        "metrics": ["Subtotal SAR", "Discount SAR", "Grand Total SAR", "Cost SAR",
                    "Gross Profit SAR", "Margin %"],
        "show": ["Offer #", "Project", "Client", "Sales Person", "System", "Status", "Date",
                 "Grand Total SAR", "Cost SAR", "Gross Profit SAR", "Margin %"],
    },
    "Line items": {
        "builder": lines_df,
        "filters": ["Client", "Sales Person", "System", "Brand", "Status", "Currency"],
        "date": "Date",
        "metrics": ["Qty", "Cost SAR", "Total Price SAR", "Profit SAR"],
        "show": ["Client", "System", "Brand", "Model", "Description", "Qty",
                 "Total Price SAR", "Cost SAR", "Profit SAR", "Status"],
    },
    "Finance": {
        "builder": finance_df,
        "filters": ["Client", "Sales Person", "Project Mgr", "System", "Status"],
        "date": "Date",
        "metrics": ["Grand Total SAR", "Collected SAR", "Remaining SAR", "PO Spend SAR",
                    "Net Profit SAR"],
        "show": ["Offer #", "Project", "Client", "Status", "Grand Total SAR",
                 "Collected SAR", "Remaining SAR", "PO Spend SAR", "Net Profit SAR"],
    },
}


# ----------------------------- filter + aggregate -----------------------------

def apply_filters(df: pd.DataFrame, selections: dict, date_col: str | None,
                  date_from=None, date_to=None) -> pd.DataFrame:
    """selections: {column: [allowed values]}; empty/None list = no filter on that column."""
    out = df
    for col, allowed in (selections or {}).items():
        if allowed and col in out.columns:
            out = out[out[col].isin(allowed)]
    if date_col and date_col in out.columns and (date_from or date_to):
        d = pd.to_datetime(out[date_col], errors="coerce")
        if date_from:
            out = out[d >= pd.Timestamp(date_from)]
            d = pd.to_datetime(out[date_col], errors="coerce")
        if date_to:
            out = out[d <= pd.Timestamp(date_to)]
    return out


def aggregate(df: pd.DataFrame, group_by: list, metrics: list,
              add_count: bool = True) -> pd.DataFrame:
    """Group by `group_by`, summing `metrics` (+ a Count column)."""
    if not group_by:
        return df
    agg = {m: "sum" for m in metrics if m in df.columns}
    grouped = df.groupby(group_by, dropna=False)
    out = grouped.agg(agg).reset_index() if agg else grouped.size().reset_index(name="Count")
    if add_count and agg:
        out["Count"] = grouped.size().values
    # round money/metric columns
    for m in metrics:
        if m in out.columns:
            out[m] = out[m].round(1 if "%" in m else 0)
    return out


def totals_row(df: pd.DataFrame, metrics: list) -> dict:
    """Sum of each metric column over the (already filtered) DataFrame."""
    out = {}
    for m in metrics:
        if m in df.columns and "%" not in m:
            out[m] = round(float(df[m].sum()), 0)
    return out
