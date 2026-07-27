#!/usr/bin/env python3
"""CLI: process ventas PDFs from inbox (optional Gmail OAuth fetch) into Ventas Diarias."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Allow running without installing the package
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ventas_sync.gmail_oauth import (  # noqa: E402
    describe_auth_status,
    get_gmail_service,
    gmail_paths_from_config,
)
from ventas_sync.pdf_parser import parse_totales_generales  # noqa: E402
from ventas_sync.pipeline import load_config, process_inbox  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Sync ventas PDF Totales Generales into Excel (optional Gmail OAuth fetch)"
    )
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
    parser.add_argument(
        "--auth-gmail",
        action="store_true",
        help="Run Google OAuth once and save token (browser login)",
    )
    parser.add_argument(
        "--fetch-gmail",
        action="store_true",
        help="Fetch matching PDFs from Gmail into inbox, then process",
    )
    parser.add_argument(
        "--no-fetch-gmail",
        action="store_true",
        help="Skip Gmail even if gmail.enabled=true in config",
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

    if args.auth_gmail:
        config = load_config(args.config)
        base_dir = Path(args.config).resolve().parent
        credentials_path, token_path = gmail_paths_from_config(config, base_dir)
        print(json.dumps(describe_auth_status(credentials_path, token_path), indent=2))
        get_gmail_service(credentials_path, token_path, open_browser=True)
        print(
            json.dumps(
                {
                    "action": "authenticated",
                    **describe_auth_status(credentials_path, token_path),
                },
                indent=2,
            )
        )
        return 0

    fetch: bool | None
    if args.fetch_gmail:
        fetch = True
    elif args.no_fetch_gmail:
        fetch = False
    else:
        fetch = None  # honor config gmail.enabled

    results = process_inbox(args.config, fetch_gmail=fetch)
    print(json.dumps(results, indent=2, default=str))
    if any(r.get("action") == "failed" for r in results):
        return 1
    if any(r.get("action") == "queued" for r in results):
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
