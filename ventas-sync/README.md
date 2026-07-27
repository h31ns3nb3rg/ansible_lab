# Phase 1 — Ventas PDF → Excel (La Cata LMF)

Drop a daily `VENTAS Y COBROS` PDF into `inbox/`. The script reads **Totales Generales**, upserts **Ventas Diarias** for **La Cata LMF** only, then archives the PDF.

## Setup (Mac)

```bash
cd ventas-sync
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp config.example.json config.json
```

Edit `config.json`:

- `excel_path` — path to your workbook (start with the test copy under `data/`)
- `ubicacion` — keep `La Cata LMF` for this PDF source
- `sheet_name` — `Ventas Diarias`

## Run

Any PDF filename works (`ventas.pdf`, `ventas-2.pdf`, `Ventas (1).pdf`, etc.):

```bash
# optional: check PDF extraction only
python run_ventas_sync.py --parse-only samples/ventas.pdf

# process everything in inbox/
cp /path/to/daily.pdf inbox/
python run_ventas_sync.py
```

Dates are written as real Excel dates with the same Spanish long format as your sheet (`lunes 20 de julio de 2026`), and amounts use `RD$` formatting.
## What gets written

| Excel column | Source |
|---|---|
| FECHA | PDF date |
| UBICACION | `La Cata LMF` |
| EFECTIVO | Totales Generales Efectivo |
| TARJETA (BRUTA) | Totales Generales Tarjeta |
| TRANSFERENCIAS | Totales Generales **Otros Pagos** |
| NOTAS DE CREDITO | Totales Generales Nota Cr. |
| ITBS COBRADO | PDF Itbis (when present) |
| FEE / NETA / TOTAL / VENTAS SIN ITBIS | Excel formulas (same as your sheet) |
| DEPOSITO | left alone if already filled |

Upsert key: **FECHA + UBICACION**. When a new date is added from the LMF PDF, the script also creates placeholder rows for **La Cata DF** and **La Cata SJM** (amounts 0 until filled). Multiple PDFs are processed oldest→newest and inserted in chronological order.

## Production switch

Use `launchd` or cron to run periodically:

```bash
cd /path/to/ventas-sync && /path/to/ventas-sync/.venv/bin/python run_ventas_sync.py
```

## Production switch

After testing the copy workbook, point `excel_path` in `config.json` at your main file. No other changes needed.

## Phase 2 (later)

Gmail auto-fetch, other stores, AdControl API.
