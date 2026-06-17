"""
Client-facing PDF generator (Workflow 3) — replicates the `Quotation` sheet.

Layout mirrors the Excel Quotation tab: corporate header block, grouped
System/Description/Brand/Model/Qty/Unit Price (SAR)/Total Price (SAR) table,
then the Subtotal / Discount / VAT / Grand Total block and the notes section.

Internal cost columns are hidden unless show_costs=True (admin toggle).
"""
from __future__ import annotations
import datetime as dt

import os
import math
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable, Image, PageBreak,
)

import calc
import db

# Corporate palette. BRAND is the primary colour and is overridden per company
# (from Settings) at render time; the rest are neutral accents.
BRAND = colors.HexColor("#002060")      # primary (default navy) — set per company
BRAND_LIGHT = colors.HexColor("#E6ECF5")
ACCENT = colors.HexColor("#62B22F")     # green accent rule
GREY = colors.HexColor("#6B7280")
LINE = colors.HexColor("#D1D5DB")


def _tint(hex_color, factor=0.90):
    """A light tint of a colour (blended toward white) for row/block backgrounds."""
    c = colors.HexColor(hex_color)
    return colors.Color(c.red + (1 - c.red) * factor,
                        c.green + (1 - c.green) * factor,
                        c.blue + (1 - c.blue) * factor)


def _styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("Brand", fontName="Helvetica-Bold", fontSize=20,
                          textColor=BRAND, leading=22))
    ss.add(ParagraphStyle("BrandSub", fontName="Helvetica", fontSize=8.5,
                          textColor=GREY, leading=11))
    ss.add(ParagraphStyle("DocTitle", fontName="Helvetica-Bold", fontSize=13,
                          textColor=BRAND, alignment=TA_RIGHT, leading=16))
    ss.add(ParagraphStyle("DocTitle2", fontName="Helvetica-Bold", fontSize=17,
                          textColor=BRAND, alignment=TA_LEFT, leading=19))
    ss.add(ParagraphStyle("MetaR", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, alignment=TA_RIGHT, leading=12))
    ss.add(ParagraphStyle("Lbl", fontName="Helvetica-Bold", fontSize=8.5,
                          textColor=BRAND, leading=12))
    ss.add(ParagraphStyle("Val", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, leading=12))
    ss.add(ParagraphStyle("Cell", fontName="Helvetica", fontSize=8,
                          textColor=colors.black, leading=10))
    ss.add(ParagraphStyle("CellR", fontName="Helvetica", fontSize=8,
                          alignment=TA_RIGHT, leading=10))
    ss.add(ParagraphStyle("CellC", fontName="Helvetica", fontSize=8,
                          alignment=TA_CENTER, leading=10))
    ss.add(ParagraphStyle("Grp", fontName="Helvetica-Bold", fontSize=8.5,
                          textColor=BRAND, leading=11))
    ss.add(ParagraphStyle("Note", fontName="Helvetica", fontSize=8,
                          textColor=colors.HexColor("#374151"), leading=10))
    ss.add(ParagraphStyle("NoteH", fontName="Helvetica-Bold", fontSize=8.5,
                          textColor=BRAND, leading=12))
    return ss


def _rt(s):
    """Escape text for reportlab markup and turn newlines into line breaks."""
    s = str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s.replace("\r\n", "\n").replace("\n", "<br/>")


def _sar(x):
    """Format as a whole number, rounded UP (matches the grid's round-up rule)."""
    try:
        if x is None or x == "":
            return ""
        return f"{math.ceil(round(float(x), 6)):,}"
    except (TypeError, ValueError):
        return ""


def _pct(x):
    try:
        if x is None or x == "":
            return ""
        return f"{float(x):.2f}%"
    except (TypeError, ValueError):
        return ""


def _header(story, ss, h, company):
    # The brand banner is drawn full-bleed on the page canvas (see generate_*);
    # the flowing content starts here, just below it, with the title + offer meta.
    title = _rt((h.get("title") or "Quotation").upper())
    content_w = 176 * mm

    # Document title (left) + offer meta (right), divided by the green rule.
    meta = [
        Paragraph(f"Offer #: <b>{h.get('offer','') or ''}</b>", ss["MetaR"]),
        Paragraph(f"Date: <b>{h.get('date','') or ''}</b>", ss["MetaR"]),
    ]
    t = Table([[Paragraph(title, ss["DocTitle2"]), meta]],
              colWidths=[content_w - 70 * mm, 70 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(t)
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=2, color=ACCENT))
    story.append(Spacer(1, 8))

    # Billed-to block
    meta = [
        [Paragraph("Client", ss["Lbl"]), Paragraph(h.get("client", "") or "", ss["Val"]),
         Paragraph("Project", ss["Lbl"]), Paragraph(h.get("project", "") or "", ss["Val"])],
        [Paragraph("Contact", ss["Lbl"]), Paragraph(h.get("contact", "") or "", ss["Val"]),
         Paragraph("Phone", ss["Lbl"]), Paragraph(h.get("phone", "") or "", ss["Val"])],
    ]
    if (h.get("sales") or h.get("presales")):
        meta.append(
            [Paragraph("Sales", ss["Lbl"]), Paragraph(h.get("sales", "") or "", ss["Val"]),
             Paragraph("Pre-sales", ss["Lbl"]), Paragraph(h.get("presales", "") or "", ss["Val"])])
    if h.get("pm"):
        meta.append(
            [Paragraph("Project Mgr", ss["Lbl"]), Paragraph(h.get("pm", "") or "", ss["Val"]),
             Paragraph("", ss["Lbl"]), Paragraph("", ss["Val"])])
    mt = Table(meta, colWidths=[20 * mm, 67 * mm, 20 * mm, 68 * mm])
    mt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BRAND_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.5, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.white),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(mt)
    story.append(Spacer(1, 7))
    if h.get("greeting"):
        story.append(Paragraph(_rt(h["greeting"]), ss["Note"]))
        story.append(Spacer(1, 6))


def _items_table(ss, grid: pd.DataFrame, show_costs: bool):
    if show_costs:
        headers = ["System", "Description", "Brand", "Model", "Qty",
                   "Ex Cost $", "Shipping %", "Unit Cost $", "Total Cost $",
                   "Unit Price (SAR)", "Total Price (SAR)"]
        widths = [16, 39, 15, 19, 8, 14, 13, 14, 15, 15, 16]
    else:
        headers = ["System", "Description", "Brand", "Model", "Qty",
                   "Unit Price (SAR)", "Total Price (SAR)"]
        widths = [22, 62, 22, 23, 11, 16, 19]
    scale = 175.0 / sum(widths)               # fit content to ~175mm printable width
    widths = [w * scale * mm for w in widths]

    data = [[Paragraph(f"<b>{c}</b>", ParagraphStyle("h", parent=ss["Cell"],
            textColor=colors.white, fontName="Helvetica-Bold", fontSize=8,
            alignment=TA_CENTER if i >= 4 else TA_LEFT)) for i, c in enumerate(headers)]]

    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BRAND),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    last_system = None
    r = 0
    for _, row in grid.iterrows():
        if str(row.get("LineType", "item")) == "discount":
            continue
        r += 1
        system = str(row.get("System") or "")
        # group header when System changes
        if system and system != last_system:
            data.append([Paragraph(system, ss["Grp"])] + [""] * (len(headers) - 1))
            style.append(("SPAN", (0, len(data) - 1), (-1, len(data) - 1)))
            style.append(("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), BRAND_LIGHT))
            last_system = system
        cells = [
            Paragraph("", ss["Cell"]),
            Paragraph(str(row.get("Description") or ""), ss["Cell"]),
            Paragraph(str(row.get("Brand") or ""), ss["Cell"]),
            Paragraph(str(row.get("Model") or ""), ss["Cell"]),
            Paragraph(_sar(row.get("Qty")), ss["CellC"]),
        ]
        if show_costs:
            cells += [Paragraph(_sar(row.get("Ex Unit Cost $")), ss["CellR"]),
                      Paragraph(_pct(row.get("Shipping %")), ss["CellR"]),
                      Paragraph(_sar(row.get("Unit Cost $")), ss["CellR"]),
                      Paragraph(_sar(row.get("Total Cost $")), ss["CellR"])]
        cells += [Paragraph(_sar(row.get("U. Price SAR")), ss["CellR"]),
                  Paragraph(_sar(row.get("T. Price SAR")), ss["CellR"])]
        data.append(cells)
        if r % 2 == 0:
            style.append(("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1),
                          colors.HexColor("#F7F9FC")))

    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle(style))
    return t


def _totals(ss, s: dict):
    discount_label = "Discount"
    if s.get("discount_sar") and s.get("discount_percent"):
        discount_label = f"Discount ({s['discount_percent']:.2f}%)"
    rows = [
        ("Subtotal", _sar(s["subtotal_sar"])),
        (discount_label, _sar(s["discount_sar"])),
        ("Discounted Subtotal", _sar(s["discounted_subtotal_sar"])),
        (f"VAT ({int(s['vat_rate']*100)}%)", _sar(s["vat_amount_sar"])),
        ("Grand Total (SAR)", _sar(s["grand_total_sar"])),
    ]
    data = [[Paragraph(lbl, ss["Lbl"] if i < 4 else
            ParagraphStyle("g", parent=ss["Lbl"], textColor=colors.white)),
             Paragraph(val, ss["MetaR"] if i < 4 else
            ParagraphStyle("gv", parent=ss["MetaR"], textColor=colors.white,
                           fontName="Helvetica-Bold"))]
            for i, (lbl, val) in enumerate(rows)]
    t = Table(data, colWidths=[45 * mm, 30 * mm])
    t.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -2), 0.25, LINE),
        ("BACKGROUND", (0, -1), (-1, -1), BRAND),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    # right-align the whole totals block
    wrap = Table([[t]], colWidths=[175 * mm])
    wrap.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "RIGHT")]))
    return wrap


NOTE_ORDER = ["System", "Scope", "Exclusions", "Pre-requirements", "Delivery",
              "Payment Terms", "Validity", "Notes"]


def _notes(ss, notes: dict):
    flow = []
    for label in NOTE_ORDER:
        val = notes.get(label)
        if val:
            flow.append(Paragraph(f"<b>{label}:</b>&nbsp; {_rt(val)}", ss["Note"]))
            flow.append(Spacer(1, 2))
    return flow


def generate_options_pdf(out_path, header: dict, options: list,
                         notes: dict | None = None, company: dict | None = None,
                         show_costs: bool = False) -> str:
    """Render one quotation document. `options` is a list of
    {'label', 'grid', 'summary'} — each becomes its own section (table + totals).
    A single option renders as a normal quotation."""
    global BRAND, BRAND_LIGHT
    company = company or {"name": "SmartWay Systems",
                          "tagline": "Smart &amp; Low-Current Systems",
                          "contact": "Riyadh, Kingdom of Saudi Arabia"}
    _bc = company.get("color") or "#002060"          # per-company brand colour
    BRAND = colors.HexColor(_bc)
    BRAND_LIGHT = _tint(_bc, 0.90)
    ss = _styles()
    notes = notes or {}
    options = options or []
    page_w, page_h = A4

    # Full-bleed brand banner: spans the whole paper width, flush to the top edge.
    banner = db.banner_path()
    if os.path.exists(banner):
        iw, ih = ImageReader(banner).getSize()
        banner_h = page_w * ih / iw
    else:
        banner_h = 0
    top_margin = (banner_h + 5 * mm) if banner_h else 12 * mm

    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=17 * mm, rightMargin=17 * mm,
                            topMargin=top_margin, bottomMargin=13 * mm,
                            title=f"Quotation {header.get('offer','')}")
    story = []
    _header(story, ss, header, company)
    multi = len(options) > 1
    for i, opt in enumerate(options):
        if i > 0:
            story.append(PageBreak())
        if multi:
            label = opt.get("label") or f"Option {i + 1}"
            story.append(Paragraph(f"Option {i + 1}: {_rt(label)}", ss["DocTitle2"]))
            story.append(Spacer(1, 2))
            story.append(HRFlowable(width="100%", thickness=1.5, color=ACCENT))
            story.append(Spacer(1, 5))
        story.append(_items_table(ss, opt["grid"], show_costs))
        story.append(Spacer(1, 6))
        story.append(_totals(ss, opt["summary"]))
        story.append(Spacer(1, 7))
    for f in _notes(ss, notes):
        story.append(f)
    story.append(Spacer(1, 4))
    story.append(Paragraph("Yours sincerely,", ss["Note"]))
    story.append(Paragraph(f"<b>{_rt(company['name'])}</b>", ss["NoteH"]))

    def _decorate(canvas, d):
        if banner_h:
            canvas.drawImage(banner, 0, page_h - banner_h, width=page_w, height=banner_h,
                             preserveAspectRatio=False, mask='auto')
        canvas.saveState()
        canvas.setStrokeColor(LINE)
        canvas.line(17 * mm, 12 * mm, 193 * mm, 12 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(GREY)
        canvas.drawString(17 * mm, 8 * mm, f"{company['name']} — {header.get('project','')}")
        canvas.drawRightString(193 * mm, 8 * mm, f"Page {d.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_decorate, onLaterPages=_decorate)
    return out_path


def generate_quotation_pdf(out_path, header: dict, grid: pd.DataFrame, summary: dict,
                           notes: dict | None = None, company: dict | None = None,
                           show_costs: bool = False) -> str:
    """Single-option quotation (delegates to generate_options_pdf)."""
    return generate_options_pdf(out_path, header,
                                [{"label": "", "grid": grid, "summary": summary}],
                                notes=notes, company=company, show_costs=show_costs)


if __name__ == "__main__":
    # smoke test from a real project
    import repo
    projects = repo.list_projects()
    pid = int(projects.iloc[0]["ProjectID"])
    meta = repo.project_meta(pid)
    grid = repo.load_project_grid(pid)
    summ = calc.summarize(grid)
    hdr = {"title": "Quotation", "client": meta.get("ClientName"),
           "project": meta.get("ProjectName"), "contact": meta.get("ContactName"),
           "offer": meta.get("OfferNo"), "date": meta.get("CreationDate"),
           "greeting": "Dear Sir, Thank you for the opportunity to quote the below scope of works."}
    out = generate_quotation_pdf("sample_quotation.pdf", hdr, grid, summ,
                                 notes={"Scope": "Supply, installation, testing & commissioning.",
                                        "Payment Terms": "70% advance, 30% on completion.",
                                        "Validity": "30 days from offer date."})
    print("Wrote", out, "| grand total SAR:", summ["grand_total_sar"])
