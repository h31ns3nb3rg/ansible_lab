# Cata Ventas Auto

Syncs daily sales into **`Ventas Diarias`** in `LaCata_Contabilidad_v3.xlsx` (OneDrive).

| Store | Source | How it gets in |
|---|---|---|
| **La Cata LMF** | Gmail PDF `RESUMEN DE VENTAS Y COBROS` | Auto fetch 02:00, or drop PDF in `inbox/` |
| **La Cata DF** | AdControl **Informe avanzado de cierres de caja** `.xlsx` | Drop export in `inbox/` (11:00 job) |
| **La Cata SJM** | Same cierres `.xlsx` | Same file — both stores in one export |

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

4. Install the daily schedule:

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

### 2) DF / SJM — daily cierres Excel (default)

1. In AdControl, export **Informe avanzado de cierres de caja** (`.xlsx`) for the day/range you need.
2. Drop the file into `~/Cata-Ventas-Auto/inbox/`.
3. Run:

```bash
cd ~/Cata-Ventas-Auto
source .venv/bin/activate
python run_ventas_sync.py --no-fetch-gmail
```

Or import a path directly:

```bash
python run_ventas_sync.py --import-cierres "/full/path/to/Informe_avanzado_de_cierres_de_caja.xlsx" --dry-run
python run_ventas_sync.py --import-cierres "/full/path/to/Informe_avanzado_de_cierres_de_caja.xlsx"
```

**Jul 27 example from your export (both shifts summed):**

| Store | EFECTIVO | TARJETA | PEDIDOS YA |
|---|---|---|---|
| La Cata SJM | **48225.00** | 10210.00 | 0 |
| La Cata DF | **9915.00** | 3005.00 | 1580.00 |

EFECTIVO = `Cantidad de cierre` (same as PDF `Efectivo del Dia`). Location from `Ubicación` column `(SAN JUAN)` / `(DEFILLO)`.

### 3) LMF PDF without Gmail

Drop a `VENTAS Y COBROS` PDF into `inbox/`, then:

```bash
python run_ventas_sync.py --no-fetch-gmail
```

### 4) Parse one LMF PDF (debug, no Excel write)

```bash
python run_ventas_sync.py --parse-only "/full/path/to/file.pdf"
```

### 5) Manual run via the scheduled wrapper (Notification Center banners)

```bash
VENTAS_JOB_TAG=manual-gmail ./scripts/run_job.sh --fetch-gmail
VENTAS_JOB_TAG=manual-inbox ./scripts/run_job.sh --no-fetch-gmail
```

### CLI quick reference

| Flag | Purpose |
|---|---|
| `--auth-gmail` | One-time OAuth login; saves `secrets/token.json` |
| `--fetch-gmail` | Download LMF PDFs from Gmail, then process inbox |
| `--no-fetch-gmail` | Process `inbox/` only (cierres `.xlsx` + LMF PDFs) |
| `--import-cierres FILE` | Import DF/SJM from a cierres `.xlsx` path |
| `--dry-run` | With `--import-cierres`: parse/summarize only |
| `--excel FILE` | Override workbook path (testing a copy) |
| `--parse-only PDF` | Print parsed LMF totals as JSON |
| `--config PATH` | Alternate config (default: `config.json`) |

---

## Automatic schedule (macOS)

| Job | Times (Mac local) | Command |
|---|---|---|
| Gmail / LMF | **02:00** | `--fetch-gmail` |
| DF / SJM inbox | **11:00** | `--no-fetch-gmail` (reads cierres `.xlsx` in `inbox/`) |

```bash
./schedule/macos/install_schedule.sh
./schedule/macos/uninstall_schedule.sh
```

**DF/SJM habit:** export AdControl cierres Excel and drop it in `inbox/` **before 11:00**. Include both stores / both shifts in that export.

**Sleep:** if the Mac is fully asleep at a scheduled time, that run may be skipped. Keep it plugged in or allow wake for scheduled tasks.

### Notifications (macOS)

Scheduled jobs (and `./scripts/run_job.sh …`) show a **Notification Center** banner when the run finishes:

| Result | Banner |
|---|---|
| Success (exit 0) | `La Cata Ventas` / `OK` |
| Excel locked (exit 2) | `La Cata Ventas` / `WARN` |
| Failure (exit 1+) | `La Cata Ventas` / `FAILED` |

Disable for one run: `VENTAS_NOTIFY=0 ./scripts/run_job.sh --fetch-gmail`

### Sanity report (DF/SJM)

Every run that imports a cierres `.xlsx` writes a validation block to logs and CLI output:

- log lines start with `SANITY ...`
- one line per `fecha + ubicacion` showing `ef`, `tj`, `xf`, `py`

This is intended for quick reconciliation before/after the Excel write.

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

### Cierres Excel (DF / SJM — default)

| Excel column | Cierres column |
|---|---|
| FECHA | Date from `Hora de apertura` |
| UBICACION | `(SAN JUAN)` → `La Cata SJM`, `(DEFILLO)` → `La Cata DF` |
| EFECTIVO | `Cantidad de cierre` |
| TARJETA (BRUTA) | `Total en pago con tarjeta` |
| TRANSFERENCIAS | `Transeferencia bancaria` + `Total en otros pagos` |
| PEDIDOS YA | `Total en PedidosYA` |

Same-day shifts per store are summed. LMF rows are left alone.

Register-report PDFs for DF/SJM are **disabled by default** (`register_pdfs_enabled: false`). Set `true` in `config.json` only if you need the old PDF path.

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
| `inbox/` | Drop cierres `.xlsx` (DF/SJM) and optional LMF PDFs |
| `processed/` | Successfully handled files |
| `failed/` | Files that failed to parse/write |
| `pending/` | Queued Excel updates when the file was locked |
| `logs/` | Run logs |
| `secrets/` | `credentials.json` + `token.json` (do not commit) |

---

## Troubleshooting

| Symptom | What to do |
|---|---|
| Excel locked / OneDrive “in use” | Close Excel; re-run. Updates may sit in `pending/` until the next successful run. |
| Wrong DF/SJM totals | Re-export cierres Excel, drop in `inbox/`, run `--no-fetch-gmail`. Check log lines `Cierres row …`. |
| Register PDFs ignored | Expected — use cierres `.xlsx`. Or set `register_pdfs_enabled: true`. |
| Gmail 403 / consent | Add your Google account as a test user; re-run `--auth-gmail`. |
| Schedule did not fire | Check Mac sleep; re-run `install_schedule.sh`. |
| `can't open file .../tmp/run_ventas_sync.py` | `cd ~/Cata-Ventas-Auto` before running Python. |

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

cd ~/Cata-Ventas-Auto && ./schedule/macos/install_schedule.sh
```
