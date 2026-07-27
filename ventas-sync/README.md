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

## Notes
- Close Excel before running (OneDrive)
- Never commit `secrets/`
- Multiple register shifts for the same store/day in one inbox run are **aggregated**
