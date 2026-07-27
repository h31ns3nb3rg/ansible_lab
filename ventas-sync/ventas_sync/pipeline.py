"""Process inbox PDFs into the Ventas Diarias Excel sheet."""

from __future__ import annotations

import json
import logging
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from .excel_updater import flush_pending, totales_to_row, upsert_with_retry
from .pdf_parser import parse_totales_generales


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


def process_pdf(pdf_path: Path, config: dict[str, Any], base_dir: Path) -> dict[str, Any]:
    excel_path = _resolve(base_dir, config["excel_path"])
    if not excel_path.exists():
        raise FileNotFoundError(
            f"Excel file not found at excel_path. Check OneDrive sync and config.json:\n{excel_path}"
        )
    totales = parse_totales_generales(pdf_path)
    row = totales_to_row(totales, config["ubicacion"])
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
    result["source_pdf"] = str(pdf_path)
    result["efectivo"] = row.efectivo
    result["tarjeta_bruta"] = row.tarjeta_bruta
    result["transferencias"] = row.transferencias
    result["itbs_cobrado"] = row.itbs_cobrado
    return result


def process_inbox(config_path: str | Path = "config.json") -> list[dict[str, Any]]:
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

    results: list[dict[str, Any]] = []
    # Accept any PDF name (ventas.pdf, ventas-2.pdf, Ventas (1).pdf, etc.)
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

    # Parse first, then process oldest→newest so non-consecutive days land in order.
    parsed: list[tuple[Any, Path]] = []
    for pdf in pdfs:
        try:
            totales = parse_totales_generales(pdf)
            parsed.append((totales.fecha, pdf))
        except Exception as exc:  # noqa: BLE001
            logging.exception("Failed to parse %s: %s", pdf.name, exc)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = failed / f"{pdf.stem}_{stamp}{pdf.suffix}"
            shutil.move(str(pdf), str(dest))
            results.append(
                {
                    "action": "failed",
                    "source_pdf": str(pdf),
                    "archived_to": str(dest),
                    "error": str(exc),
                }
            )
    parsed.sort(key=lambda item: (item[0], item[1].name))

    for _fecha, pdf in parsed:
        logging.info("Processing %s", pdf.name)
        try:
            result = process_pdf(pdf, config, base_dir)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = processed / f"{pdf.stem}_{stamp}{pdf.suffix}"
            shutil.move(str(pdf), str(dest))
            result["archived_to"] = str(dest)
            if result.get("action") == "queued":
                logging.error(
                    "Excel locked/unavailable — update queued for later: %s",
                    result,
                )
            else:
                logging.info("OK wrote %s", result)
            results.append(result)
        except Exception as exc:  # noqa: BLE001 - top-level per-file handler
            logging.exception("Failed %s: %s", pdf.name, exc)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = failed / f"{pdf.stem}_{stamp}{pdf.suffix}"
            shutil.move(str(pdf), str(dest))
            results.append(
                {
                    "action": "failed",
                    "source_pdf": str(pdf),
                    "archived_to": str(dest),
                    "error": str(exc),
                }
            )
    return results
