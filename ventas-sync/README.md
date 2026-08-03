# Cata Ventas Auto

Syncs daily sales into **`Ventas Diarias`** in `LaCata_Contabilidad_v3.xlsx` (OneDrive).

| Store | Source | How it gets in |
|---|---|---|
| **La Cata LMF** | Gmail PDF `RESUMEN DE VENTAS Y COBROS` | Auto fetch 01:00 & 11:00, or drop PDF in `inbox/` |
| **La Cata DF** | AdControl **Informe avanzado de cierres de caja** `.xlsx` | Drop export in `inbox/` (01:00 & 11:00 job) |
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

### All stores at once (same as the schedule)

One command processes **LMF (Gmail PDF) + DF/SJM (cierres Excel) + inbox PDFs**:

```bash
cd ~/Cata-Ventas-Auto
source .venv/bin/activate
python run_ventas_sync.py --fetch-gmail
```

Or with the same Notification Center banner as the scheduled job:

```bash
cd ~/Cata-Ventas-Auto
VENTAS_JOB_TAG=manual-all ./scripts/run_job.sh --fetch-gmail
```

`--fetch-gmail` does this in order:

1. Download matching Gmail PDFs into `inbox/` (LMF)
2. Process AdControl **cierres** `.xlsx` in `inbox/` (DF / SJM)
3. Process inbox PDFs
4. Retention cleanup (`logs/`, `processed/`, `failed/` older than `retention_days`)

**Habit:** drop the AdControl cierres `.xlsx` into `inbox/` before you run (or before 01:00 / 11:00).

### DF / SJM only — import a cierres Excel path

**Quote the path** — AdControl filenames have spaces:

```bash
cd ~/Cata-Ventas-Auto
source .venv/bin/activate

# Option A: file already in inbox/ (processes all cierres .xlsx there)
python run_ventas_sync.py --no-fetch-gmail

# Option B: import by name or full path (quotes required)
python run_ventas_sync.py --import-cierres "inbox/Informe avanzado de cierres de caja 3_8_2026-2.xlsx"
python run_ventas_sync.py --import-cierres "Informe avanzado de cierres de caja 3_8_2026-2.xlsx"
```

Bare filenames are also looked up under `inbox/`.

Dry-run first if you want:

```bash
python run_ventas_sync.py --import-cierres "inbox/Informe avanzado de cierres de caja 3_8_2026-2.xlsx" --dry-run
```

**Aug 2 example from your export (Pago en efectivo; fondo excluded):**

| Store | Shifts | Fondo/shift | EFECTIVO | TARJETA | PEDIDOS YA |
|---|---|---|---|---|---|
| La Cata SJM | 2 | 5,000 | **43770.00** | 5750.00 | 0 |
| La Cata DF | 2 | 2,500 | **17140.00** | 20880.00 | 3380.00 |

EFECTIVO = `Pago en efectivo` (= `Cantidad de cierre` − `Fondo de Caja`). Location from `Ubicación` column `(SAN JUAN)` / `(DEFILLO)`.

### Inbox only (no Gmail fetch)

Drop cierres `.xlsx` and/or LMF PDFs into `inbox/`, then:

```bash
python run_ventas_sync.py --no-fetch-gmail
# or with banner:
VENTAS_JOB_TAG=manual-inbox ./scripts/run_job.sh --no-fetch-gmail
```

### Parse one LMF PDF (debug, no Excel write)

```bash
python run_ventas_sync.py --parse-only "/full/path/to/file.pdf"
```

### CLI quick reference

| Flag | Purpose |
|---|---|
| `--auth-gmail` | One-time OAuth login; saves `secrets/token.json` |
| `--fetch-gmail` | Gmail PDFs + cierres `.xlsx` + inbox PDFs (all stores) |
| `--no-fetch-gmail` | Process `inbox/` only (cierres `.xlsx` + LMF PDFs) |
| `--import-cierres FILE` | Import DF/SJM from a cierres `.xlsx` path |
| `--cleanup` | Delete `logs/` `processed/` `failed/` files older than `retention_days` |
| `--dry-run` | With `--import-cierres` or `--cleanup`: preview only |
| `--excel FILE` | Override workbook path (testing a copy) |
| `--parse-only PDF` | Print parsed LMF totals as JSON |
| `--config PATH` | Alternate config (default: `config.json`) |

---

## Automatic schedule (macOS)

One LaunchAgent processes **PDF + Excel (all stores) in the same run**:

| Job | Times (Mac local) | What it does |
|---|---|---|
| `com.lacata.ventas.sync` | **01:00** and **11:00** | `--fetch-gmail` → Gmail LMF PDFs + cierres Excel + inbox PDFs + cleanup |

```bash
./schedule/macos/install_schedule.sh
./schedule/macos/uninstall_schedule.sh
```

Re-running install also removes older separate agents (`com.lacata.ventas.gmail` / `com.lacata.ventas.inbox`) if present.

**DF/SJM habit:** export AdControl cierres Excel and drop it in `inbox/` **before 01:00 or 11:00**. Include both stores / both shifts in that export.

**Sleep:** if the Mac is fully asleep at a scheduled time, that run may be skipped. Keep it plugged in or allow wake for scheduled tasks.

Manual equivalent of the scheduled job:

```bash
VENTAS_JOB_TAG=manual-all ./scripts/run_job.sh --fetch-gmail
```

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
| `logs/scheduled_sync_*.log` | Timed combined runs (Gmail + Excel + inbox) |
| `logs/scheduled_manual-all_*.log` | Manual wrapper runs with `VENTAS_JOB_TAG=manual-all` |
| `logs/launchd_sync.*.log` | launchd stdout/stderr for the combined agent |
| `logs/ventas_sync_YYYYMMDD.log` | App log for the day |

---

## Field mapping

### Cierres Excel (DF / SJM — default)

| Excel column | Cierres column |
|---|---|
| FECHA | Date from `Hora de apertura` |
| UBICACION | `(SAN JUAN)` → `La Cata SJM`, `(DEFILLO)` → `La Cata DF` |
| EFECTIVO | `Pago en efectivo` (sales cash; excludes `Fondo de Caja`) |
| TARJETA (BRUTA) | `Total en pago con tarjeta` |
| TRANSFERENCIAS | `Transeferencia bancaria` + `Total en otros pagos` |
| PEDIDOS YA | `Total en PedidosYA` |

`Fondo de Caja` is the POS float (typically **5,000** SJM / **2,500** DF per shift). It is **not** written to Ventas Diarias. `Cantidad de cierre` includes that float; we use `Pago en efectivo` instead so EFECTIVO is sales-only. If `Pago en efectivo` is missing on an older export, the importer falls back to `Cantidad de cierre − Fondo de Caja` (or the default floats above).

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
| `processed/` | Successfully handled files (auto-deleted after **7 days**) |
| `failed/` | Files that failed to parse/write (auto-deleted after **7 days**) |
| `pending/` | Queued Excel updates when the file was locked |
| `logs/` | Run logs (auto-deleted after **7 days**) |
| `secrets/` | `credentials.json` + `token.json` (do not commit) |

### Retention / cleanup
Keeps at most **7 days** of `logs/`, `processed/`, and `failed/` (configurable via `retention_days` in `config.json`). Inbox is never auto-deleted.

Cleanup runs automatically at the end of every sync. Manual:

```bash
python run_ventas_sync.py --cleanup --dry-run   # preview
python run_ventas_sync.py --cleanup             # delete now
```

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
