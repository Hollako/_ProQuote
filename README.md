# ProQuote — BoQ pricing & quotation engine

Python + SQLite system that ingests the historical project workbooks under
`J:\My Drive\1-Projects`, builds a deduplicated item catalogue, and powers a
Streamlit interface for creating new offers and exporting client-facing
Quotation PDFs (SAR, VAT, no internal costs).

## Layout
| File | Role |
|------|------|
| `db.py` | SQLite schema + connection (`proquote.db`) |
| `ingest.py` | Folder-wide Excel scanner → catalogue + project history |
| `calc.py` | Pricing formulas (single source of truth) |
| `repo.py` | Catalogue search + project load/save |
| `pdf_export.py` | Client-facing Quotation PDF (reportlab) |
| `app.py` | Streamlit UI (New Offer / Load Existing / Catalogue) |

## Setup (one-time)
```powershell
cd "J:\My Drive\1-Projects\_ProQuote"
py -m pip install -r requirements.txt
py ingest.py            # build / refresh proquote.db from the project folder
```

## Run the app
```powershell
cd "J:\My Drive\1-Projects\_ProQuote"
py -m streamlit run app.py
```
Then open http://localhost:8501.

## Re-ingest when new projects are added
`py ingest.py` is **idempotent** — it re-reads every workbook and refreshes
the catalogue. Point it at a different folder with
`py ingest.py "C:\some\other\folder"`.

## Key conventions (from the real sheets)
- Sheets come in `BOQ <System>` / `Pivot <System>` / `Quotation <System>` triplets.
- Pricing chain (validated against the historical sheets, 99.9% match):
  - `Unit Cost $ = Ex Unit Cost × 1.3`  (landed cost; manual if no Ex cost)
  - `U. Price $  = ROUNDUP(Unit Cost × Margin, 0)`  (**Margin ×** is per-line input)
  - `U. Price SAR = ROUNDUP(U. Price $ × 3.75 → next multiple of 10)`
  - `Total Cost = Qty × Unit Cost`; `T. Price = Qty × U. Price`;
    `T. Price SAR = Qty × U. Price SAR`
  - Set **Margin × = 0** on a line to type `U. Price $` manually (e.g. services).
- USD→SAR peg = **3.75**; VAT = **15%**. Constants live at the top of `calc.py`
  (`SAR_PER_USD`, `COST_BUFFER`, `SAR_ROUND_TO`).
- Markup factor = `Total Selling SAR ÷ (Total Cost USD × 3.75)` (≈1.69 headline).
- The PDF hides List/Ex/Unit/Total Cost columns unless the **Admin** toggle is on.

## Branding
The client PDF uses the real **SmartWay Systems** banner at
`assets/header_banner.png` (navy #002060 + logo green #62B22F accent). To
rebrand, replace that PNG with one of similar aspect ratio — the header scales
to fit the page width automatically; no code change needed.

## Notes
- 12 non-template workbooks (third-party video-wall quotes, MEP electrical BOQs)
  are intentionally skipped to keep the catalogue clean.
- `Items_Catalog` is keyed on `(Brand, Model, Description)`; re-ingest refreshes
  default prices to the most recently seen values.
