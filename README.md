# ProQuote - BoQ pricing & quotation engine

Python + PostgreSQL/SQLite system that ingests the historical project workbooks under
`J:\My Drive\1-Projects`, builds a deduplicated item catalogue, and powers a
Streamlit interface for creating new offers and exporting client-facing
Quotation PDFs (SAR, VAT, no internal costs).

## Layout
| File | Role |
|------|------|
| `db.py` | Database selector: PostgreSQL via `DATABASE_URL`, otherwise local SQLite |
| `db_postgres.py` | PostgreSQL schema, pooled connections, triggers and SQL compatibility |
| `db_transfer.py` | Portable backups and SQLite-to-PostgreSQL migration |
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

## Persistent PostgreSQL deployment

ProQuote uses PostgreSQL whenever `DATABASE_URL` is configured. Without it, the
app retains the local `proquote.db` fallback. For Streamlit Community Cloud, add
the **pooled** runtime URL in the app's Secrets panel:

```toml
DATABASE_URL = "postgresql://USER:PASSWORD@HOST-pooler/DATABASE?sslmode=require"
POSTGRES_MIGRATION_URL = "postgresql://USER:PASSWORD@DIRECT-HOST/DATABASE?sslmode=require"
```

Never commit a real database URL or password to GitHub. For a local PostgreSQL
run, set `DATABASE_URL` in the environment before starting Streamlit.

### One-time SQLite migration

1. Create an empty PostgreSQL database.
2. Put its **direct** (non-pooled) URL in `POSTGRES_MIGRATION_URL` locally.
3. Run:

```powershell
py migrate_to_postgres.py --sqlite "C:\path\to\proquote.db"
```

The migration preserves IDs, users/password hashes, roles, offers, sheets,
catalogue items, line items, finance, audit history and branding assets. It
verifies every table count and the aggregate offer value before reporting
success. Use `--replace` only when intentionally replacing an already populated
PostgreSQL target.

For normal Streamlit traffic, switch `DATABASE_URL` back to the provider's
**pooled** URL. PostgreSQL profile backups remain downloadable ZIP files and can
be restored from Settings > Backup & Restore.

## Re-ingest when new projects are added
`py ingest.py` is **idempotent** - it re-reads every workbook and refreshes
the catalogue. Point it at a different folder with
`py ingest.py "C:\some\other\folder"`.

## Key conventions (from the real sheets)
- Sheets come in `BOQ <System>` / `Pivot <System>` / `Quotation <System>` triplets.
- Pricing chain (validated against the historical sheets, 99.9% match):
  - `Unit Cost $ = Ex Unit Cost x 1.3`  (landed cost; manual if no Ex cost)
  - `U. Price $  = ROUNDUP(Unit Cost x Margin, 0)`  (**Margin x** is per-line input)
  - `U. Price SAR = ROUNDUP(U. Price $ x 3.75 → next multiple of 10)`
  - `Total Cost = Qty x Unit Cost`; `T. Price = Qty x U. Price`;
    `T. Price SAR = Qty x U. Price SAR`
  - Set **Margin x = 0** on a line to type `U. Price $` manually (e.g. services).
- USD→SAR peg = **3.75**; VAT = **15%**. Constants live at the top of `calc.py`
  (`SAR_PER_USD`, `COST_BUFFER`, `SAR_ROUND_TO`).
- Markup factor = `Total Selling SAR ÷ (Total Cost USD x 3.75)` (≈1.69 headline).
- The PDF hides List/Ex/Unit/Total Cost columns unless the **Admin** toggle is on.

## Branding
The client PDF uses the real **SmartWay Systems** banner at
`assets/header_banner.png` (navy #002060 + logo green #62B22F accent). To
rebrand, replace that PNG with one of similar aspect ratio - the header scales
to fit the page width automatically; no code change needed.

## Notes
- 12 non-template workbooks (third-party video-wall quotes, MEP electrical BOQs)
  are intentionally skipped to keep the catalogue clean.
- `Items_Catalog` is keyed on `(Brand, Model, Description)`; re-ingest refreshes
  default prices to the most recently seen values.
