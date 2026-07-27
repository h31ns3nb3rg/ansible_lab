#!/usr/bin/env python3
"""CLI: process ventas PDFs from inbox into Ventas Diarias."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Allow running without installing the package
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ventas_sync.pipeline import process_inbox  # noqa: E402
from ventas_sync.pdf_parser import parse_totales_generales  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Sync ventas PDF Totales Generales into Excel")
    parser.add_argument(
        "--config",
        default=str(ROOT / "config.json"),
        help="Path to config.json",
    )
    parser.add_argument(
        "--parse-only",
        metavar="PDF",
        help="Only parse a PDF and print JSON (no Excel write)",
    )
    args = parser.parse_args()

    if args.parse_only:
        totales = parse_totales_generales(args.parse_only)
        print(
            json.dumps(
                {
                    "fecha": totales.fecha.strftime("%Y-%m-%d"),
                    "efectivo": totales.efectivo,
                    "tarjeta": totales.tarjeta,
                    "otros_pagos_transferencias": totales.otros_pagos,
                    "nota_credito": totales.nota_credito,
                    "valor_neto": totales.valor_neto,
                    "itbis": totales.itbis,
                    "cantidad": totales.cantidad,
                },
                indent=2,
            )
        )
        return 0

    results = process_inbox(args.config)
    print(json.dumps(results, indent=2, default=str))
    if any(r.get("action") == "failed" for r in results):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
