"""
Client-facing PDF generator (Workflow 3) - replicates the `Quotation` sheet.

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
    LongTable,
)

import calc
import db

# Corporate palette. BRAND is the primary colour and is overridden per company
# (from Settings) at render time; the rest are neutral accents.
BRAND = colors.HexColor("#002060")      # primary (default navy) - set per company
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


def _sar2(x):
    try:
        if x is None or x == "":
            return ""
        return f"{float(x):,.2f}"
    except (TypeError, ValueError):
        return ""


def _display_date(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            d = dt.datetime.strptime(text, fmt).date()
            return d.strftime("%B ") + str(d.day) + d.strftime(",%Y")
        except ValueError:
            pass
    return text


def _slot_text(template: str, company: dict, header: dict, page: int) -> str:
    values = {
        "company": company.get("name") or "",
        "project": header.get("project") or "",
        "offer": header.get("offer") or "",
        "page": page,
        "vat": company.get("vat_number") or "",
        "vat_number": company.get("vat_number") or "",
        "cr": company.get("cr_number") or "",
        "cr_number": company.get("cr_number") or "",
    }
    try:
        return str(template or "").format(**values)
    except (KeyError, ValueError):
        return str(template or "")


def _fit_size(path: str, max_w: float, max_h: float) -> tuple[float, float]:
    iw, ih = ImageReader(path).getSize()
    scale = min(max_w / iw, max_h / ih)
    return iw * scale, ih * scale


def _asset_path(filename: str) -> str:
    base = getattr(db, "ASSETS_DIR", "")
    if not base:
        data_dir = getattr(db, "DATA_DIR", getattr(db, "APP_DIR", os.path.dirname(os.path.abspath(__file__))))
        base = os.path.join(data_dir, "assets")
    return os.path.join(base, filename)


def _db_asset_path(helper_name: str, filename: str) -> str:
    helper = getattr(db, helper_name, None)
    if callable(helper):
        return helper()
    return _asset_path(filename)


def _draw_fit_image(canvas, path: str, x: float, y: float, max_w: float, max_h: float,
                    align: str = "center") -> None:
    if not os.path.exists(path):
        return
    w, h = _fit_size(path, max_w, max_h)
    if align == "right":
        ix = x + max_w - w
    elif align == "left":
        ix = x
    else:
        ix = x + (max_w - w) / 2
    iy = y + (max_h - h) / 2
    canvas.drawImage(path, ix, iy, width=w, height=h, preserveAspectRatio=True, mask="auto")


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


def _template2_styles(ss):
    ss.add(ParagraphStyle("T2MetaL", fontName="Helvetica-Bold", fontSize=9.5,
                          textColor=colors.black, leading=12))
    ss.add(ParagraphStyle("T2MetaR", fontName="Helvetica", fontSize=9.5,
                          textColor=colors.black, leading=12, alignment=TA_RIGHT))
    ss.add(ParagraphStyle("T2Subject", fontName="Helvetica-Bold", fontSize=9.5,
                          textColor=colors.black, leading=12))
    ss.add(ParagraphStyle("T2Body", fontName="Helvetica", fontSize=9.5,
                          textColor=colors.black, leading=13))
    ss.add(ParagraphStyle("T2TermH", fontName="Helvetica-Bold", fontSize=9.5,
                          textColor=colors.black, leading=11))
    ss.add(ParagraphStyle("T2TableH", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.white, leading=11))
    ss.add(ParagraphStyle("T2Cell", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, leading=11))
    ss.add(ParagraphStyle("T2CellC", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, leading=11, alignment=TA_CENTER))
    ss.add(ParagraphStyle("T2CellR", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, leading=11, alignment=TA_RIGHT))
    ss.add(ParagraphStyle("T2Group", fontName="Helvetica", fontSize=9.5,
                          textColor=colors.black, leading=12))
    ss.add(ParagraphStyle("T2Total", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.white, leading=10))
    ss.add(ParagraphStyle("T2TotalR", fontName="Helvetica-Bold", fontSize=8.5,
                          textColor=colors.white, leading=10, alignment=TA_RIGHT))
    ss.add(ParagraphStyle("T2TotalDark", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, leading=10))
    ss.add(ParagraphStyle("T2TotalDarkR", fontName="Helvetica", fontSize=8.5,
                          textColor=colors.black, leading=10, alignment=TA_RIGHT))


def _template2_add_text(flow, text: str, style, blank_space: float = 5) -> None:
    lines = str(text or "").replace("\r\n", "\n").split("\n")
    for line in lines:
        if line.strip():
            flow.append(Paragraph(_rt(line), style))
        else:
            flow.append(Spacer(1, blank_space))


def _template2_intro(story, ss, h, notes: dict, company: dict) -> None:
    content_w = 176 * mm
    left = [
        Paragraph(f"Project:{_rt(h.get('project') or '')}", ss["T2MetaL"]),
        Spacer(1, 19),
        Paragraph(f"M/S {_rt(h.get('client') or '')}", ss["T2MetaL"]),
    ]
    right_lines = []
    if h.get("sales"):
        right_lines.append(f"From: {h.get('sales')}")
    if h.get("phone"):
        right_lines.append(f"Phone: {h.get('phone')}")
    if h.get("offer"):
        right_lines.append(f"Reference: {h.get('offer')}")
    if h.get("date"):
        right_lines.append(f"Date:{_display_date(h.get('date'))}")
    right = [Paragraph(_rt(line), ss["T2MetaR"]) for line in right_lines]
    meta = Table([[left, right]], colWidths=[content_w * 0.58, content_w * 0.42])
    meta.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(meta)
    story.append(Spacer(1, 28))

    title = h.get("title") or "Quotation"
    story.append(Paragraph(f"<u>Subject: {_rt(title)}</u>", ss["T2Subject"]))
    story.append(Spacer(1, 22))

    greeting = h.get("greeting") or ""
    if greeting:
        _template2_add_text(story, greeting, ss["T2Body"], blank_space=10)
        story.append(Spacer(1, 10))

    free_notes = str(notes.get("Notes") or "").strip()
    if free_notes:
        story.append(Paragraph("Below are some notes to be taken into consideration:", ss["T2Body"]))
        story.append(Spacer(1, 12))
        _template2_add_text(story, free_notes, ss["T2Body"], blank_space=3)
        story.append(Spacer(1, 14))

    intro_text = f"{greeting}\n{free_notes}".lower()
    if "we look forward" not in intro_text:
        story.append(Paragraph(
            "We look forward for your favorable reply, meanwhile we remain at your disposal "
            "for any further information that you may need.",
            ss["T2Body"],
        ))
        story.append(Spacer(1, 18))
    else:
        story.append(Spacer(1, 8))

    term_order = [
        ("Scope", notes.get("Scope")),
        ("System", notes.get("System")),
        ("Exclusions", notes.get("Exclusions")),
        ("Pre-requirements", notes.get("Pre-requirements")),
        ("Payment Terms", notes.get("Payment Terms")),
        ("Validity", notes.get("Validity")),
        ("Delivery", notes.get("Delivery")),
    ]
    for label, value in term_order:
        if not str(value or "").strip():
            continue
        story.append(Paragraph(f"<u>{_rt(label)}</u>", ss["T2TermH"]))
        _template2_add_text(story, value, ss["T2Body"], blank_space=2)
        story.append(Spacer(1, 6))


def _template2_price_cells(row, included: bool):
    if included:
        return "", "Included"
    return _sar(row.get("U. Price SAR")), _sar(row.get("T. Price SAR"))


def _template2_items_table(ss, grid: pd.DataFrame, summary: dict, show_costs: bool):
    dark = colors.HexColor("#37689B")
    group_fill = colors.HexColor("#92D0DC")
    system_fill = colors.HexColor("#96B3D3")
    body_fill = colors.HexColor("#C5D9F1")
    headers = ["System", "Description", "Brand", "Model", "Qty", "U. Price<br/>SAR", "T. Price SAR"]
    widths = [18, 69, 17, 22, 12, 19, 19]
    widths = [w * mm for w in widths]
    data = [[Paragraph(c, ss["T2TableH"]) for c in headers]]
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), dark),
        ("LINEABOVE", (0, 0), (-1, 0), 0.7, colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, 0), 4),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]

    last_area_key = None
    last_system = None
    for _, row in grid.iterrows():
        if str(row.get("LineType", "item")) == "discount":
            continue
        area = str(row.get("Area") or "").strip()
        system = str(row.get("System") or "").strip()
        group_label = area or system
        group_key = group_label.casefold()
        if group_label and group_key != last_area_key:
            data.append([Paragraph(_rt(group_label), ss["T2Group"])] + [""] * 6)
            rr = len(data) - 1
            style += [
                ("SPAN", (0, rr), (-1, rr)),
                ("BACKGROUND", (0, rr), (-1, rr), group_fill),
                ("TOPPADDING", (0, rr), (-1, rr), 2),
                ("BOTTOMPADDING", (0, rr), (-1, rr), 2),
            ]
            last_area_key = group_key
            last_system = None

        show_system = system if system and system != last_system else ""
        if system:
            last_system = system
        included = str(row.get("LineType", "item")).lower() in {"service", "included"}
        unit_price, total_price = _template2_price_cells(row, included)
        data.append([
            Paragraph(_rt(show_system), ss["T2Cell"]),
            Paragraph(_rt(row.get("Description") or ""), ss["T2Cell"]),
            Paragraph(_rt(row.get("Brand") or ""), ss["T2Cell"]),
            Paragraph(_rt(row.get("Model") or ""), ss["T2Cell"]),
            Paragraph(_sar(row.get("Qty")), ss["T2CellC"]),
            Paragraph(unit_price, ss["T2CellR"]),
            Paragraph(total_price, ss["T2CellR"]),
        ])
        rr = len(data) - 1
        style += [
            ("BACKGROUND", (0, rr), (0, rr), system_fill),
            ("BACKGROUND", (1, rr), (1, rr), body_fill),
            ("BACKGROUND", (3, rr), (3, rr), body_fill),
            ("TOPPADDING", (0, rr), (-1, rr), 2),
            ("BOTTOMPADDING", (0, rr), (-1, rr), 2),
        ]

    total_rows = [
        ("Grand Total", summary.get("subtotal_sar"), True),
    ]
    if summary.get("discount_sar"):
        total_rows.append((f"Discount ({summary.get('discount_percent', 0):.2f}%)",
                           summary.get("discount_sar"), False))
        total_rows.append(("Total After Discount", summary.get("discounted_subtotal_sar"), True))
    total_rows += [
        (f"VAT {int(summary.get('vat_rate', 0) * 100)}%", summary.get("vat_amount_sar"), False),
        ("Net Total", summary.get("grand_total_sar"), True),
    ]

    for label, value, filled in total_rows:
        if filled:
            row = [
                Paragraph(_rt(label), ss["T2Total"]), "", "", "", "",
                Paragraph("SAR", ss["T2TotalR"]),
                Paragraph(_sar2(value), ss["T2TotalR"]),
            ]
        else:
            row = [
                Paragraph(_rt(label), ss["T2TotalDark"]), "", "", "", "",
                Paragraph("SAR", ss["T2TotalDarkR"]),
                Paragraph(_sar2(value), ss["T2TotalDarkR"]),
            ]
        data.append(row)
        rr = len(data) - 1
        style += [
            ("SPAN", (0, rr), (4, rr)),
            ("BACKGROUND", (0, rr), (-1, rr), dark if filled else colors.white),
            ("TOPPADDING", (0, rr), (-1, rr), 2),
            ("BOTTOMPADDING", (0, rr), (-1, rr), 2),
        ]

    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle(style))
    return t


def _template2_boq_section(story, ss, opt: dict, option_label: str = "") -> None:
    title = "BOQ:" if not option_label else f"BOQ: {option_label}"
    story.append(Paragraph(_rt(title), ss["T2MetaL"]))
    story.append(Spacer(1, 14))
    story.append(_template2_items_table(ss, opt["grid"], opt["summary"], show_costs=False))
    story.append(Spacer(1, 34))


def _template2_signature(story, ss, h, company):
    story.append(Paragraph("Yours sincerely,", ss["T2Body"]))
    story.append(Spacer(1, 22))
    if h.get("sales"):
        story.append(Paragraph(_rt(h.get("sales")), ss["T2Body"]))
    story.append(Paragraph(_rt(company.get("name") or ""), ss["T2Body"]))


def generate_options_pdf(out_path, header: dict, options: list,
                         notes: dict | None = None, company: dict | None = None,
                         show_costs: bool = False, template: str = "template1") -> str:
    """Render one quotation document. `options` is a list of
    {'label', 'grid', 'summary'} - each becomes its own section (table + totals).
    A single option renders as a normal quotation."""
    global BRAND, BRAND_LIGHT
    company = company or {"name": "Company Name",
                          "tagline": "Smart &amp; Low-Current Systems",
                          "contact": "Riyadh, Kingdom of Saudi Arabia",
                          "header_left": "{company}",
                          "header_middle": "",
                          "header_right": "{offer}",
                          "footer_left": "{company} - {project}",
                          "footer_middle": "",
                          "footer_right": "Page {page}"}
    _bc = company.get("color") or "#002060"          # per-company brand colour
    BRAND = colors.HexColor(_bc)
    BRAND_LIGHT = _tint(_bc, 0.90)
    ss = _styles()
    template_key = str(template or "template1").strip().lower().replace(" ", "")
    is_template2 = template_key in {"2", "template2", "template-2"}
    if is_template2:
        _template2_styles(ss)
    notes = notes or {}
    options = options or []
    page_w, page_h = A4

    # Full-bleed brand banner: spans the whole paper width, flush to the top edge.
    # If no banner is present, the three-section PDF header is drawn instead.
    banner = _db_asset_path("banner_path", "header_banner.png")
    if os.path.exists(banner):
        iw, ih = ImageReader(banner).getSize()
        banner_h = page_w * ih / iw
    else:
        banner_h = 0
    section_header_h = 22 * mm
    header_image_paths = [
        _db_asset_path("header_left_path", "header_left.png"),
        _db_asset_path("header_middle_path", "header_middle.png"),
        _db_asset_path("header_right_path", "header_right.png"),
    ]
    header_text_templates = [
        company.get("header_left", ""),
        company.get("header_middle", ""),
        company.get("header_right", ""),
    ]
    has_section_header = (not banner_h) and (
        any(os.path.exists(path) for path in header_image_paths) or
        any(str(text or "").strip() for text in header_text_templates)
    )
    top_margin = (banner_h + 5 * mm) if banner_h else (
        (section_header_h + 5 * mm) if has_section_header else 12 * mm
    )

    full_footer = _db_asset_path("footer_full_path", "footer_full.png")
    if os.path.exists(full_footer):
        fiw, fih = ImageReader(full_footer).getSize()
        footer_h = page_w * fih / fiw
    else:
        footer_h = 22 * mm
    bottom_margin = max(13 * mm, footer_h + 7 * mm)

    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=17 * mm, rightMargin=17 * mm,
                            topMargin=top_margin, bottomMargin=bottom_margin,
                            title=f"Quotation {header.get('offer','')}")
    story = []
    if is_template2:
        _template2_intro(story, ss, header, notes, company)
        multi = len(options) > 1
        for i, opt in enumerate(options):
            story.append(PageBreak())
            label = opt.get("label") or f"Option {i + 1}"
            _template2_boq_section(story, ss, opt, label if multi else "")
            if i == len(options) - 1:
                _template2_signature(story, ss, header, company)
    else:
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
        canvas.saveState()
        if banner_h:
            canvas.drawImage(banner, 0, page_h - banner_h, width=page_w, height=banner_h,
                             preserveAspectRatio=False, mask="auto")
        elif has_section_header:
            left_x = 17 * mm
            right_x = 193 * mm
            bottom_y = page_h - section_header_h
            canvas.setStrokeColor(LINE)
            canvas.line(left_x, bottom_y, right_x, bottom_y)
            col_gap = 5 * mm
            col_w = (right_x - left_x - 2 * col_gap) / 3
            cols = [
                (left_x, "left", _db_asset_path("header_left_path", "header_left.png"), company.get("header_left", "")),
                (left_x + col_w + col_gap, "center", _db_asset_path("header_middle_path", "header_middle.png"), company.get("header_middle", "")),
                (left_x + 2 * (col_w + col_gap), "right", _db_asset_path("header_right_path", "header_right.png"), company.get("header_right", "")),
            ]
            align_map = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT}
            for x, align, image_path, text_template in cols:
                has_image = os.path.exists(image_path)
                text = _slot_text(text_template, company, header, d.page)
                if has_image:
                    _draw_fit_image(canvas, image_path, x, page_h - 13 * mm, col_w, 8 * mm, align=align)
                if text:
                    style = ParagraphStyle(
                        f"header_{align}",
                        fontName="Helvetica",
                        fontSize=8,
                        leading=9,
                        textColor=BRAND,
                        alignment=align_map[align],
                    )
                    p = Paragraph(_rt(text), style)
                    _, used_h = p.wrap(col_w, 10 * mm)
                    p.drawOn(canvas, x, bottom_y + 5 * mm + max(0, (8 * mm - used_h) / 2))

        if os.path.exists(full_footer):
            canvas.drawImage(full_footer, 0, 0, width=page_w, height=footer_h,
                             preserveAspectRatio=False, mask="auto")
        else:
            left_x = 17 * mm
            right_x = 193 * mm
            top_y = 24 * mm
            canvas.setStrokeColor(LINE)
            canvas.line(left_x, top_y, right_x, top_y)
            col_gap = 5 * mm
            col_w = (right_x - left_x - 2 * col_gap) / 3
            cols = [
                (left_x, "left", _db_asset_path("footer_left_path", "footer_left.png"), company.get("footer_left", "")),
                (left_x + col_w + col_gap, "center", _db_asset_path("footer_middle_path", "footer_middle.png"), company.get("footer_middle", "")),
                (left_x + 2 * (col_w + col_gap), "right", _db_asset_path("footer_right_path", "footer_right.png"), company.get("footer_right", "")),
            ]
            align_map = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT}
            for x, align, image_path, text_template in cols:
                has_image = os.path.exists(image_path)
                text = _slot_text(text_template, company, header, d.page)
                if has_image:
                    _draw_fit_image(canvas, image_path, x, 13 * mm, col_w, 8 * mm, align=align)
                if text:
                    style = ParagraphStyle(
                        f"footer_{align}",
                        fontName="Helvetica",
                        fontSize=7,
                        leading=8,
                        textColor=GREY,
                        alignment=align_map[align],
                    )
                    p = Paragraph(_rt(text), style)
                    _, used_h = p.wrap(col_w, 10 * mm)
                    p.drawOn(canvas, x, 5 * mm + max(0, (8 * mm - used_h) / 2))
        canvas.restoreState()

    doc.build(story, onFirstPage=_decorate, onLaterPages=_decorate)
    return out_path


def generate_quotation_pdf(out_path, header: dict, grid: pd.DataFrame, summary: dict,
                           notes: dict | None = None, company: dict | None = None,
                           show_costs: bool = False, template: str = "template1") -> str:
    """Single-option quotation (delegates to generate_options_pdf)."""
    return generate_options_pdf(out_path, header,
                                [{"label": "", "grid": grid, "summary": summary}],
                                notes=notes, company=company, show_costs=show_costs,
                                template=template)


# ============================ REPORTS PDF ============================

def _fmt_cell(col, val):
    if isinstance(val, bool):
        return "Yes" if val else "No"
    if isinstance(val, (int, float)):
        return f"{val:.1f}" if "%" in str(col) else f"{val:,.0f}"
    return str(val if val is not None else "")


def _report_table(ss, df, totals):
    cols = list(df.columns)
    numeric = {c for c in cols if pd.api.types.is_numeric_dtype(df[c])}
    head = ParagraphStyle("rh", parent=ss["Cell"], textColor=colors.white,
                          fontName="Helvetica-Bold", fontSize=7.5, leading=9)
    cell = ParagraphStyle("rc", parent=ss["Cell"], fontSize=7.5, leading=9)
    cellR = ParagraphStyle("rcr", parent=ss["CellR"], fontSize=7.5, leading=9)
    data = [[Paragraph(_rt(c), head) for c in cols]]
    for _, row in df.iterrows():
        data.append([Paragraph(_rt(_fmt_cell(c, row[c])), cellR if c in numeric else cell)
                     for c in cols])
    if totals:
        tcell = ParagraphStyle("tt", parent=cell, textColor=colors.white, fontName="Helvetica-Bold")
        tcellR = ParagraphStyle("ttr", parent=cellR, textColor=colors.white, fontName="Helvetica-Bold")
        trow = []
        for i, c in enumerate(cols):
            if i == 0:
                trow.append(Paragraph("TOTAL", tcell))
            elif c in totals:
                trow.append(Paragraph(f"{totals[c]:,.0f}", tcellR))
            else:
                trow.append("")
        data.append(trow)
    n = len(cols)
    avail = 182 * mm
    t = LongTable(data, colWidths=[avail / n] * n, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BRAND),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2 if totals else -1),
         [colors.white, colors.HexColor("#F4F7FB")]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    if totals:
        style.append(("BACKGROUND", (0, -1), (-1, -1), BRAND))
    t.setStyle(TableStyle(style))
    return t


def _report_decorate(canvas, d, banner, banner_h, company, title):
    page_w, page_h = A4
    canvas.saveState()
    if banner_h:
        canvas.drawImage(banner, 0, page_h - banner_h, width=page_w, height=banner_h,
                         preserveAspectRatio=False, mask="auto")
    else:
        canvas.setFillColor(BRAND)
        canvas.rect(0, page_h - 16 * mm, page_w, 16 * mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 13)
        canvas.drawString(14 * mm, page_h - 11 * mm, company.get("name") or "")
    canvas.setStrokeColor(LINE)
    canvas.line(14 * mm, 12 * mm, 196 * mm, 12 * mm)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(GREY)
    canvas.drawString(14 * mm, 8 * mm, f"{company.get('name','')} — {title}")
    canvas.drawRightString(196 * mm, 8 * mm, f"Page {d.page}")
    canvas.restoreState()


def generate_report_pdf(out_path, title: str, subtitle_lines, table_df=None, totals=None,
                        company: dict | None = None, chart_paths=None) -> str:
    """Branded internal report: title, filter summary, optional charts, and a (multi-page) table."""
    global BRAND, BRAND_LIGHT
    company = company or {"name": "Company"}
    _bc = company.get("color") or "#002060"
    BRAND = colors.HexColor(_bc)
    BRAND_LIGHT = _tint(_bc, 0.90)
    ss = _styles()
    page_w, page_h = A4
    banner = db.banner_path()
    if os.path.exists(banner):
        iw, ih = ImageReader(banner).getSize()
        banner_h = page_w * ih / iw
    else:
        banner_h = 0
    top_margin = (banner_h + 6 * mm) if banner_h else 20 * mm

    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=top_margin, bottomMargin=15 * mm, title=title)
    story = [Paragraph(_rt(title), ss["DocTitle2"]), Spacer(1, 2),
             HRFlowable(width="100%", thickness=2, color=ACCENT), Spacer(1, 6)]
    for line in (subtitle_lines or []):
        if line:
            story.append(Paragraph(_rt(line), ss["Note"]))
    if subtitle_lines:
        story.append(Spacer(1, 8))
    for cp in (chart_paths or []):
        if cp and os.path.exists(cp):
            iw, ih = ImageReader(cp).getSize()
            w = 182 * mm
            story.append(Image(cp, width=w, height=w * ih / iw))
            story.append(Spacer(1, 10))
    if table_df is not None and not table_df.empty:
        story.append(_report_table(ss, table_df, totals))

    def _dec(canvas, d):
        _report_decorate(canvas, d, banner, banner_h, company, title)

    doc.build(story, onFirstPage=_dec, onLaterPages=_dec)
    return out_path


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
