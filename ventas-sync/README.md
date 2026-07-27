# Phase 1 + Gmail OAuth — Ventas PDF → Excel (La Cata LMF)

Drop a daily `VENTAS Y COBROS` PDF into `inbox/`, or let the script fetch it from Gmail with **OAuth (read-only)**. The script reads **Totales Generales**, upserts **Ventas Diarias** for **La Cata LMF**, creates DF/SJM stub rows for that date, then archives the PDF.

## Setup (Mac)

```bash
cd ventas-sync   # or ~/Cata-Ventas-Auto
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp config.example.json config.json
```

Edit `config.json` → `excel_path` (your OneDrive workbook).

## Manual inbox run

Any PDF filename works (`ventas.pdf`, `ventas-2.pdf`, etc.):

```bash
cp /path/to/daily.pdf inbox/
python run_ventas_sync.py --no-fetch-gmail
```

## Phase 2 — Gmail OAuth (read-only)

### 1) Google Cloud (once)
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create/select a project
3. Enable **Gmail API**
4. Configure OAuth consent screen (External is fine for personal use; add your Gmail as a test user)
5. Create credentials → **OAuth client ID** → Application type **Desktop app**
6. Download the JSON and save as:

```text
secrets/credentials.json
```

Never commit this file (gitignored).

### 2) Authorize once on your Mac

```bash
source .venv/bin/activate
python run_ventas_sync.py --auth-gmail
```

Browser opens → sign in → Allow. Creates `secrets/token.json` (also gitignored).

### 3) Fetch + sync

```bash
# one-shot
python run_ventas_sync.py --fetch-gmail

# or set in config.json: "gmail": { "enabled": true, ... }
python run_ventas_sync.py
```

Default search: PDFs from the last 14 days whose filename matches `ventas*.pdf`. Tune `gmail.query` / `gmail.filename_regex` in config.

**Security:** scope is `gmail.readonly` only. Revoke anytime in Google Account → Security → Third-party access.

## What gets written

| Excel column | Source |
|---|---|
| FECHA | Spanish text date (`lunes 20 de julio de 2026`) |
| UBICACION | `La Cata LMF` |
| EFECTIVO | Totales Generales Efectivo |
| TARJETA (BRUTA) | Totales Generales Tarjeta |
| TRANSFERENCIAS | Totales Generales **Otros Pagos** |
| NOTAS DE CREDITO | Totales Generales Nota Cr. |
| ITBS COBRADO | PDF Itbis (when present) |
| FEE / NETA / TOTAL / VENTAS SIN ITBIS | Excel formulas |
| DEPOSITO | left alone if already filled |

Also creates stub rows for **La Cata DF** and **La Cata SJM** on that date.

## Schedule (optional)

```bash
cd /path/to/ventas-sync && .venv/bin/python run_ventas_sync.py --fetch-gmail
```
