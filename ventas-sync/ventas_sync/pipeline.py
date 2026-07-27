"""Process inbox PDFs into the Ventas Diarias Excel sheet."""

from __future__ import annotations

import json
import logging
import shutil
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .excel_updater import DailySaleRow, flush_pending, totales_to_row, upsert_with_retry
from .gmail_fetch import fetch_gmail_pdfs
from .pdf_parser import parse_totales_generales
from .register_parser import detect_pdf_kind, parse_register_report, register_to_row


def load_config(config_path: str | Path) -> dict[str, Any]:
    path = Path(config_path)
    return json.loads(path.read_text(encoding="utf-8"))


def setup_logging(logs_dir: str | Path) -> Path:
    logs = Path(logs_dir)
    logs.mkdir(parents=True, exist_ok=True)
    log_file = logs / f"ventas_sync_{datetime.now():%Y%m%d}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )
    return log_file


def _resolve(base: Path, maybe_relative: str) -> Path:
    path = Path(maybe_relative).expanduser()
    return path if path.is_absolute() else (base / path).resolve()


def _archive(pdf: Path, folder: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = folder / f"{pdf.stem}_{stamp}{pdf.suffix}"
    shutil.move(str(pdf), str(dest))
    return dest


def _write_row(
    row: DailySaleRow,
    config: dict[str, Any],
    base_dir: Path,
    source_pdfs: list[str],
) -> dict[str, Any]:
    excel_path = _resolve(base_dir, config["excel_path"])
    ubicaciones = config.get("ubicaciones") or [
        "La Cata LMF",
        "La Cata DF",
        "La Cata SJM",
    ]
    result = upsert_with_retry(
        excel_path=excel_path,
        sheet_name=config["sheet_name"],
        data=row,
        fee_rate_cell=config.get("fee_rate_cell", "Setup!$B$6"),
        retries=int(config.get("lock_retries", 5)),
        retry_seconds=float(config.get("lock_retry_seconds", 2)),
        pending_path=_resolve(base_dir, config.get("pending_path", "pending/updates.json")),
        ubicaciones=ubicaciones,
    )
    result["source_pdfs"] = source_pdfs
    result["efectivo"] = row.efectivo
    result["tarjeta_bruta"] = row.tarjeta_bruta
    result["transferencias"] = row.transferencias
    result["pedidos_ya"] = row.pedidos_ya
    result["itbs_cobrado"] = row.itbs_cobrado
    return result


def process_pdf(pdf_path: Path, config: dict[str, Any], base_dir: Path) -> dict[str, Any]:
    """Process a single LMF ventas PDF (kept for CLI compatibility)."""
    excel_path = _resolve(base_dir, config["excel_path"])
    if not excel_path.exists():
        raise FileNotFoundError(
            f"Excel file not found at excel_path. Check OneDrive sync and config.json:\n{excel_path}"
        )
    kind = detect_pdf_kind(pdf_path)
    if kind == "adcontrol_register":
        row = register_to_row(parse_register_report(pdf_path))
    elif kind == "lmf_ventas":
        row = totales_to_row(parse_totales_generales(pdf_path), config["ubicacion"])
    else:
        raise ValueError(f"Unrecognized PDF type: {pdf_path.name}")
    result = _write_row(row, config, base_dir, [str(pdf_path)])
    result["source_pdf"] = str(pdf_path)
    result["pdf_kind"] = kind
    return result


def process_inbox(
    config_path: str | Path = "config.json",
    *,
    fetch_gmail: bool | None = None,
) -> list[dict[str, Any]]:
    config_file = Path(config_path).resolve()
    base_dir = config_file.parent
    config = load_config(config_file)
    setup_logging(_resolve(base_dir, config.get("logs_dir", "logs")))

    inbox = _resolve(base_dir, config.get("inbox_dir", "inbox"))
    processed = _resolve(base_dir, config.get("processed_dir", "processed"))
    failed = _resolve(base_dir, config.get("failed_dir", "failed"))
    for folder in (inbox, processed, failed):
        folder.mkdir(parents=True, exist_ok=True)

    excel_path = _resolve(base_dir, config["excel_path"])
    pending_path = _resolve(base_dir, config.get("pending_path", "pending/updates.json"))

    logging.info("Config: %s", config_file)
    logging.info("Excel target: %s (exists=%s)", excel_path, excel_path.exists())
    if not excel_path.exists():
        raise FileNotFoundError(
            f"Excel file not found. Update excel_path or wait for OneDrive to download:\n{excel_path}"
        )

    should_fetch = (
        fetch_gmail
        if fetch_gmail is not None
        else bool((config.get("gmail") or {}).get("enabled", False))
    )
    gmail_results: list[dict[str, Any]] = []
    if should_fetch:
        logging.info("Fetching PDFs from Gmail (OAuth read-only)")
        gmail_results = fetch_gmail_pdfs(config, base_dir, force=True)

    logging.info("Flushing pending updates (if any)")
    ubicaciones = config.get("ubicaciones") or [
        "La Cata LMF",
        "La Cata DF",
        "La Cata SJM",
    ]
    flushed = flush_pending(
        pending_path=pending_path,
        excel_path=excel_path,
        sheet_name=config["sheet_name"],
        fee_rate_cell=config.get("fee_rate_cell", "Setup!$B$6"),
        retries=int(config.get("lock_retries", 5)),
        retry_seconds=float(config.get("lock_retry_seconds", 2)),
        ubicaciones=ubicaciones,
    )
    for item in flushed:
        logging.info("Flushed pending: %s", item)

    results: list[dict[str, Any]] = list(gmail_results)
    pdfs = sorted(
        {
            *inbox.glob("*.pdf"),
            *inbox.glob("*.PDF"),
            *inbox.glob("*.Pdf"),
        }
    )
    if not pdfs:
        logging.info("No PDFs in inbox: %s", inbox)
        return results

    lmf_jobs: list[tuple[date, Path, DailySaleRow]] = []
    # Aggregate register shifts by (fecha, ubicacion)
    register_agg: dict[tuple[date, str], DailySaleRow] = {}
    register_sources: dict[tuple[date, str], list[Path]] = defaultdict(list)

    for pdf in pdfs:
        try:
            kind = detect_pdf_kind(pdf)
            logging.info("Detected %s as %s", pdf.name, kind)
            if kind == "lmf_ventas":
                totales = parse_totales_generales(pdf)
                row = totales_to_row(totales, config["ubicacion"])
                lmf_jobs.append((row.fecha, pdf, row))
            elif kind == "adcontrol_register":
                report = parse_register_report(pdf)
                row = register_to_row(report)
                key = (row.fecha, row.ubicacion)
                register_sources[key].append(pdf)
                if key not in register_agg:
                    register_agg[key] = row
                else:
                    existing = register_agg[key]
                    register_agg[key] = DailySaleRow(
                        fecha=existing.fecha,
                        ubicacion=existing.ubicacion,
                        efectivo=existing.efectivo + row.efectivo,
                        tarjeta_bruta=existing.tarjeta_bruta + row.tarjeta_bruta,
                        transferencias=existing.transferencias + row.transferencias,
                        notas_credito=existing.notas_credito + row.notas_credito,
                        itbs_cobrado=(existing.itbs_cobrado or 0) + (row.itbs_cobrado or 0),
                        pedidos_ya=(existing.pedidos_ya or 0) + (row.pedidos_ya or 0),
                    )
            else:
                raise ValueError("Unrecognized PDF type (expected LMF ventas or AdControl register report)")
        except Exception as exc:  # noqa: BLE001
            logging.exception("Failed to parse %s: %s", pdf.name, exc)
            dest = _archive(pdf, failed)
            results.append(
                {
                    "action": "failed",
                    "source_pdf": str(pdf),
                    "archived_to": str(dest),
                    "error": str(exc),
                }
            )

    # Process LMF PDFs oldest→newest (one row each)
    lmf_jobs.sort(key=lambda item: (item[0], item[1].name))
    for _fecha, pdf, row in lmf_jobs:
        logging.info("Processing LMF PDF %s", pdf.name)
        try:
            result = _write_row(row, config, base_dir, [str(pdf)])
            dest = _archive(pdf, processed)
            result["archived_to"] = str(dest)
            result["pdf_kind"] = "lmf_ventas"
            result["source_pdf"] = str(pdf)
            if result.get("action") == "queued":
                logging.error("Excel locked/unavailable — update queued: %s", result)
            else:
                logging.info("OK wrote %s", result)
            results.append(result)
        except Exception as exc:  # noqa: BLE001
            logging.exception("Failed %s: %s", pdf.name, exc)
            dest = _archive(pdf, failed)
            results.append(
                {
                    "action": "failed",
                    "source_pdf": str(pdf),
                    "archived_to": str(dest),
                    "error": str(exc),
                }
            )

    # Process aggregated register reports oldest→newest
    for key in sorted(register_agg.keys(), key=lambda k: (k[0], k[1])):
        row = register_agg[key]
        sources = register_sources[key]
        logging.info(
            "Processing register aggregate %s %s from %s PDF(s)",
            row.fecha,
            row.ubicacion,
            len(sources),
        )
        try:
            result = _write_row(row, config, base_dir, [str(p) for p in sources])
            archived = []
            for pdf in sources:
                archived.append(str(_archive(pdf, processed)))
            result["archived_to"] = archived
            result["pdf_kind"] = "adcontrol_register"
            result["shifts_aggregated"] = len(sources)
            if result.get("action") == "queued":
                logging.error("Excel locked/unavailable — update queued: %s", result)
            else:
                logging.info("OK wrote %s", result)
            results.append(result)
        except Exception as exc:  # noqa: BLE001
            logging.exception("Failed register aggregate %s: %s", key, exc)
            archived = []
            for pdf in sources:
                if pdf.exists():
                    archived.append(str(_archive(pdf, failed)))
            results.append(
                {
                    "action": "failed",
                    "source_pdfs": [str(p) for p in sources],
                    "archived_to": archived,
                    "error": str(exc),
                }
            )

    return results
