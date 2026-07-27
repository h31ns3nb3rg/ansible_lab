# Phase 1+2 — Ventas PDF → Excel (La Cata LMF / DF / SJM)

## Sources
- **La Cata LMF:** Gmail OAuth (`RESUMEN DE VENTAS Y COBROS`) or drop `VENTAS Y COBROS` PDF in `inbox/`
- **La Cata DF / SJM:** manually download AdControl **register-report** PDFs and drop in `inbox/`

The script auto-detects PDF type.

## Setup (Mac)

```bash
cd ~/Cata-Ventas-Auto
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## LMF (Gmail)
```bash
python run_ventas_sync.py --auth-gmail   # once
python run_ventas_sync.py --fetch-gmail
```

## DF / SJM (manual register PDFs)
1. In AdControl, open each cash-register detail report and save/print as PDF
2. Drop into `inbox/` (one PDF per shift is OK; same day+store are summed)
3. Run:
```bash
python run_ventas_sync.py --no-fetch-gmail
```

### Register PDF mapping
| Excel | PDF field |
|---|---|
| FECHA | start date from `Detalles de caja (...)` |
| UBICACION | `(SAN JUAN)` → `La Cata SJM`, `(DEFILLO)` → `La Cata DF` |
| EFECTIVO | `Efectivo del Dia` |
| TARJETA (BRUTA) | `Pago con tarjeta` |
| TRANSFERENCIAS | `Transferencia bancaria` + `Otros pagos` |
| PEDIDOS YA | `PedidosYA` |
| ITBS COBRADO | `Impuesto` |

## DF / SJM bulk backfill (cierres Excel)
From AdControl, export **Informe avanzado de cierres de caja** (`.xlsx`), then:

```bash
# Preview only
python run_ventas_sync.py --import-cierres "/path/to/Informe_avanzado_de_cierres_de_caja.xlsx" --dry-run

# Write into OneDrive workbook (close Excel first)
python run_ventas_sync.py --import-cierres "/path/to/Informe_avanzado_de_cierres_de_caja.xlsx"
```

| Excel | Cierres column |
|---|---|
| FECHA | date from `Hora de apertura` |
| UBICACION | `(SAN JUAN)` → `La Cata SJM`, `(DEFILLO)` → `La Cata DF` |
| EFECTIVO | `Cantidad de cierre` (= PDF `Efectivo del Dia`) |
| TARJETA (BRUTA) | `Total en pago con tarjeta` |
| TRANSFERENCIAS | `Transeferencia bancaria` + `Total en otros pagos` |
| PEDIDOS YA | `Total en PedidosYA` |

Same-day shifts per store are summed. LMF rows are left alone.

## Schedule (macOS LaunchAgents)

| Job | Time (Mac local) | What it does |
|---|---|---|
| Gmail / LMF | **02:00** | `python run_ventas_sync.py --fetch-gmail` |
| DF / SJM inbox | **11:00** | `python run_ventas_sync.py --no-fetch-gmail` (reads `inbox/`) |

### Install (once)
```bash
cd ~/Cata-Ventas-Auto
chmod +x scripts/run_job.sh schedule/macos/*.sh
./schedule/macos/install_schedule.sh
```

### Daily DF/SJM habit
Drop AdControl register-report PDFs into `~/Cata-Ventas-Auto/inbox/` **before 11:00**. The 11am job processes whatever is there (same-day shifts are summed).

### Logs
- Per-run: `logs/scheduled_gmail-2am_*.log`, `logs/scheduled_inbox-11am_*.log`
- launchd: `logs/launchd_gmail.*.log`, `logs/launchd_inbox.*.log`

### Manual / uninstall
```bash
VENTAS_JOB_TAG=manual-gmail ./scripts/run_job.sh --fetch-gmail
VENTAS_JOB_TAG=manual-inbox ./scripts/run_job.sh --no-fetch-gmail
./schedule/macos/uninstall_schedule.sh
```

**Sleep note:** if the Mac is fully asleep at 2am/11am, that run may be skipped until the next day. Keep it plugged in, or enable waking for scheduled tasks in Energy settings.

## Notes
- Close Excel before running when possible (OneDrive). If locked, updates go to `pending/` and flush on the next successful run.
- Never commit `secrets/`
- Multiple register shifts for the same store/day in one inbox run are **aggregated**
