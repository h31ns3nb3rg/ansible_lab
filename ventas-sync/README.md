# Cata Ventas Auto

Syncs daily sales into **`Ventas Diarias`** in `LaCata_Contabilidad_v3.xlsx` (OneDrive).

| Store | Source | How it gets in |
|---|---|---|
| **La Cata LMF** | Gmail PDF `RESUMEN DE VENTAS Y COBROS` | Auto fetch, or drop PDF in `inbox/` |
| **La Cata DF** | AdControl register-report PDF | Drop PDF in `inbox/` |
| **La Cata SJM** | AdControl register-report PDF | Drop PDF in `inbox/` |
| **DF / SJM backfill** | AdControl *Informe avanzado de cierres de caja* `.xlsx` | One-time / bulk `--import-cierres` |

Upsert key: **`FECHA + UBICACION`**. Always keeps **3 rows per day** (LMF, DF, SJM).

---

## One-time setup (Mac)

```bash
cd ~/Cata-Ventas-Auto
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

1. Confirm `config.json` → `excel_path` points at your OneDrive workbook.
2. Put Google OAuth `credentials.json` in `secrets/` (never commit secrets).
3. Authenticate Gmail once (browser login):

```bash
cd ~/Cata-Ventas-Auto
source .venv/bin/activate
python run_ventas_sync.py --auth-gmail
```

4. Install the daily schedule (optional but recommended):

```bash
chmod +x scripts/run_job.sh schedule/macos/*.sh
./schedule/macos/install_schedule.sh
```

**Before any write:** close Excel if possible so OneDrive does not lock the file.

---

## Manual runs

Activate the venv first:

```bash
cd ~/Cata-Ventas-Auto
source .venv/bin/activate
```

### 1) LMF — fetch from Gmail and update Excel

```bash
python run_ventas_sync.py --fetch-gmail
```

Downloads matching `VENTAS Y COBROS` PDFs, processes them, writes `Ventas Diarias`.

### 2) DF / SJM — process PDFs already in `inbox/`

1. In AdControl, open each cash-register detail and save/print as PDF.
2. Copy the PDFs into `~/Cata-Ventas-Auto/inbox/`.
3. Run:

```bash
python run_ventas_sync.py --no-fetch-gmail
```

Same-day shifts for the same store are **summed**. PDF type is auto-detected.

### 3) LMF PDF without Gmail

Drop a `VENTAS Y COBROS` PDF into `inbox/`, then:

```bash
python run_ventas_sync.py --no-fetch-gmail
```

### 4) Bulk backfill DF / SJM from cierres Excel

Export **Informe avanzado de cierres de caja** from AdControl, then:

```bash
# Preview only (no Excel write)
python run_ventas_sync.py --import-cierres "/full/path/to/Informe_avanzado_de_cierres_de_caja.xlsx" --dry-run

# Write into the workbook (close Excel first)
python run_ventas_sync.py --import-cierres "/full/path/to/Informe_avanzado_de_cierres_de_caja.xlsx"
```

### 5) Parse one PDF (debug, no Excel write)

```bash
python run_ventas_sync.py --parse-only "/full/path/to/file.pdf"
```

### 6) Manual run via the scheduled wrapper (writes under `logs/`)

```bash
VENTAS_JOB_TAG=manual-gmail ./scripts/run_job.sh --fetch-gmail
VENTAS_JOB_TAG=manual-inbox ./scripts/run_job.sh --no-fetch-gmail
```

### CLI quick reference

| Flag | Purpose |
|---|---|
| `--auth-gmail` | One-time OAuth login; saves `secrets/token.json` |
| `--fetch-gmail` | Download LMF PDFs from Gmail, then process inbox |
| `--no-fetch-gmail` | Process `inbox/` only (DF/SJM and/or dropped LMF PDFs) |
| `--import-cierres FILE` | Bulk-load DF/SJM from cierres `.xlsx` |
| `--dry-run` | With `--import-cierres`: parse/summarize only |
| `--excel FILE` | Override workbook path (testing a copy) |
| `--parse-only PDF` | Print parsed totals as JSON |
| `--config PATH` | Alternate config (default: `config.json`) |

---

## Automatic schedule (macOS)

| Job | Times (Mac local) | Command |
|---|---|---|
| Gmail / LMF | **02:00** | `--fetch-gmail` |
| DF / SJM inbox | **11:00** | `--no-fetch-gmail` |

```bash
# Install / reload after schedule changes
./schedule/macos/install_schedule.sh

# Remove
./schedule/macos/uninstall_schedule.sh
```

**DF/SJM habit:** drop **all** register-report PDFs for the day (both shifts) into `inbox/` before **11:00**. One daily run aggregates every PDF in the folder for that date/store.

**Sleep:** if the Mac is fully asleep at a scheduled time, that run may be skipped. Keep it plugged in or allow wake for scheduled tasks.

### Logs

| Location | What |
|---|---|
| `logs/scheduled_gmail_*.log` | Timed / wrapper Gmail runs |
| `logs/scheduled_inbox_*.log` | Timed / wrapper inbox runs |
| `logs/launchd_gmail.*.log` | launchd stdout/stderr (Gmail) |
| `logs/launchd_inbox.*.log` | launchd stdout/stderr (inbox) |
| `logs/ventas_sync_YYYYMMDD.log` | App log for the day |

---

## Field mapping

### Register PDF (DF / SJM)

| Excel column | PDF field |
|---|---|
| FECHA | Start date from `Detalles de caja (...)` |
| UBICACION | `(SAN JUAN)` → `La Cata SJM`, `(DEFILLO)` → `La Cata DF` |
| EFECTIVO | `Efectivo del Dia` |
| TARJETA (BRUTA) | `Pago con tarjeta` |
| TRANSFERENCIAS | `Transferencia bancaria` + `Otros pagos` |
| PEDIDOS YA | `PedidosYA` |
| ITBS COBRADO | `Impuesto` |

### Cierres Excel (bulk DF / SJM)

| Excel column | Cierres column |
|---|---|
| FECHA | Date from `Hora de apertura` |
| UBICACION | `(SAN JUAN)` → `La Cata SJM`, `(DEFILLO)` → `La Cata DF` |
| EFECTIVO | `Cantidad de cierre` (same as PDF `Efectivo del Dia`) |
| TARJETA (BRUTA) | `Total en pago con tarjeta` |
| TRANSFERENCIAS | `Transeferencia bancaria` + `Total en otros pagos` |
| PEDIDOS YA | `Total en PedidosYA` |

### Derived Excel formulas (unchanged)

- FEE = tarjeta × `Setup!$B$6` (3.45%)
- TARJETA NETA = tarjeta − fee
- TOTAL = efectivo + tarjeta neta + transferencias − notas crédito
- VENTAS SIN ITBIS = total − ITBS

Filled **DEPOSITO EFECTIVO** is never overwritten.

---

## Folders

| Folder | Role |
|---|---|
| `inbox/` | Drop PDFs here to process |
| `processed/` | Successfully handled PDFs |
| `failed/` | PDFs that failed to parse/write |
| `pending/` | Queued Excel updates when the file was locked |
| `logs/` | Run logs |
| `secrets/` | `credentials.json` + `token.json` (do not commit) |

---

## Troubleshooting

| Symptom | What to do |
|---|---|
| Excel locked / OneDrive “in use” | Close Excel; re-run. Updates may sit in `pending/` until the next successful run. |
| Gmail 403 / consent | Add your Google account as a test user on the OAuth consent screen; re-run `--auth-gmail`. |
| `--import-cierres` unrecognized | Your folder is on an old copy — refresh from the latest branch, then retry. |
| Schedule did not fire | Check Mac sleep; confirm agents with `launchctl print gui/$(id -u)/com.lacata.ventas.gmail` (and `.inbox`). Re-run `install_schedule.sh`. |
| Wrong / missing store | Confirm PDF has `(DEFILLO)` or `(SAN JUAN)` in ubicación comercial. |

---

## Update the app code

```bash
cd /tmp
rm -rf ansible_lab_tmp
git clone --branch cursor/ventas-pdf-excel-phase1-5495 --single-branch \
  https://github.com/h31ns3nb3rg/ansible_lab.git ansible_lab_tmp

rsync -a --exclude 'secrets' --exclude 'config.json' --exclude '.venv' \
  --exclude 'inbox' --exclude 'processed' --exclude 'failed' \
  --exclude 'pending' --exclude 'logs' \
  /tmp/ansible_lab_tmp/ventas-sync/ ~/Cata-Ventas-Auto/

# If schedule files changed:
cd ~/Cata-Ventas-Auto && ./schedule/macos/install_schedule.sh
```
