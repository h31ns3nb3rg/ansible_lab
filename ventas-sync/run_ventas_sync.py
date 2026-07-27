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
from ventas_sync.pipeline import import_cierres_excel, load_config, process_inbox  # noqa: E402


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
        "--import-cierres",
        metavar="XLSX",
        help="Bulk-import AdControl 'Informe avanzado de cierres de caja' into Ventas Diarias (DF/SJM)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="With --import-cierres: parse and summarize only (no Excel write)",
    )
    parser.add_argument(
        "--excel",
        metavar="XLSX",
        help="Override config excel_path (useful for testing a local workbook copy)",
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

    if args.import_cierres:
        result = import_cierres_excel(
            args.import_cierres,
            args.config,
            dry_run=args.dry_run,
            excel_path_override=args.excel,
        )
        # Keep JSON readable: drop per-row detail unless dry-run
        printable = dict(result)
        if not args.dry_run and "results" in printable:
            printable["results_sample"] = printable["results"][:5]
            printable["results_omitted"] = max(0, len(result.get("results", [])) - 5)
            del printable["results"]
        print(json.dumps(printable, indent=2, default=str))
        print("\n=== Cierres import resume ===")
        print(f"Shifts: {result.get('shifts')}")
        print(f"Day/location rows: {result.get('days_locations')} ({result.get('by_ubicacion')})")
        print(f"Range: {result.get('date_from')} → {result.get('date_to')}")
        if args.dry_run:
            print("Dry run — Excel not modified")
        else:
            print(
                f"Written: {result.get('written')} "
                f"(inserted={result.get('inserted')}, updated={result.get('updated')})"
            )
        print("=============================\n")
        return 0

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

    # Human-readable resume
    gmail_summary = next((r for r in results if r.get("action") == "gmail_summary"), None)
    downloaded = [r for r in results if r.get("action") == "downloaded"]
    written = [r for r in results if r.get("action") in {"inserted", "updated"}]
    failed = [r for r in results if r.get("action") == "failed"]
    queued = [r for r in results if r.get("action") == "queued"]

    print("\n=== Run resume ===")
    if gmail_summary:
        print(
            "Gmail: "
            f"{gmail_summary.get('emails_found', 0)} emails found, "
            f"{gmail_summary.get('pdfs_seen', 0)} PDFs seen, "
            f"{gmail_summary.get('pdfs_downloaded', 0)} downloaded, "
            f"{gmail_summary.get('pdfs_skipped', 0)} skipped"
        )
        print(f"Gmail query: {gmail_summary.get('query')}")
    else:
        print("Gmail: not fetched this run")
    print(f"Excel rows written: {len(written)} (inserted/updated)")
    if failed:
        print(f"Failed: {len(failed)}")
    if queued:
        print(f"Queued (Excel locked): {len(queued)}")
    print("==================\n")

    print(json.dumps(results, indent=2, default=str))
    if failed:
        return 1
    if queued:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
