"""Fetch ventas PDF attachments from Gmail via OAuth."""

from __future__ import annotations

import base64
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from .gmail_oauth import get_gmail_service, gmail_paths_from_config


def _safe_filename(name: str) -> str:
    name = Path(name).name
    name = re.sub(r"[^\w.\- ()+]+", "_", name, flags=re.UNICODE).strip(" ._")
    if not name.lower().endswith(".pdf"):
        name = f"{name}.pdf" if name else "ventas.pdf"
    return name or "ventas.pdf"


def _unique_path(directory: Path, filename: str) -> Path:
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for i in range(2, 1000):
        alt = directory / f"{stem}-{i}{suffix}"
        if not alt.exists():
            return alt
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return directory / f"{stem}-{stamp}{suffix}"


def _walk_parts(payload: dict[str, Any]) -> list[dict[str, Any]]:
    parts = payload.get("parts") or []
    found: list[dict[str, Any]] = []
    if payload.get("filename") and payload.get("body", {}).get("attachmentId"):
        found.append(payload)
    for part in parts:
        found.extend(_walk_parts(part))
    return found


def _header_map(payload: dict[str, Any]) -> dict[str, str]:
    headers = payload.get("headers") or []
    return {h.get("name", "").lower(): h.get("value", "") for h in headers}


def fetch_gmail_pdfs(
    config: dict[str, Any],
    base_dir: Path,
    *,
    force: bool = False,
) -> list[dict[str, Any]]:
    """Search Gmail and save matching PDF attachments into inbox_dir.

    If force=True (CLI --fetch-gmail), run even when gmail.enabled is false.
    Returns download records plus a final summary dict (action=gmail_summary).
    """
    gmail_cfg = config.get("gmail") or {}
    if not force and not gmail_cfg.get("enabled", False):
        logging.info("Gmail fetch disabled in config (gmail.enabled=false)")
        return []

    credentials_path, token_path = gmail_paths_from_config(config, base_dir)
    inbox = Path(config.get("inbox_dir", "inbox"))
    inbox = inbox if inbox.is_absolute() else (base_dir / inbox)
    inbox.mkdir(parents=True, exist_ok=True)

    query = gmail_cfg.get(
        "query",
        'subject:"RESUMEN DE VENTAS Y COBROS" has:attachment filename:pdf newer_than:30d',
    )
    max_messages = int(gmail_cfg.get("max_messages", 50))
    filename_regex = gmail_cfg.get("filename_regex", r"(?i).*\.pdf$")
    pattern = re.compile(filename_regex)

    service = get_gmail_service(credentials_path, token_path)
    logging.info("Gmail search: %s", query)

    response = (
        service.users()
        .messages()
        .list(userId="me", q=query, maxResults=max_messages)
        .execute()
    )
    messages = response.get("messages") or []
    results: list[dict[str, Any]] = []

    emails_found = len(messages)
    pdfs_seen = 0
    pdfs_skipped = 0
    pdfs_downloaded = 0
    emails_with_downloads = 0

    for item in messages:
        msg_id = item["id"]
        message = (
            service.users()
            .messages()
            .get(userId="me", id=msg_id, format="full")
            .execute()
        )
        payload = message.get("payload") or {}
        headers = _header_map(payload)
        subject = headers.get("subject", "")
        sender = headers.get("from", "")
        parts = _walk_parts(payload)
        downloaded_this_email = 0

        for part in parts:
            filename = part.get("filename") or ""
            if not filename.lower().endswith(".pdf"):
                continue
            pdfs_seen += 1
            if filename_regex and not pattern.search(filename):
                pdfs_skipped += 1
                logging.info(
                    "Skipping attachment %r (filename_regex mismatch) from %s",
                    filename,
                    subject,
                )
                continue

            att_id = (part.get("body") or {}).get("attachmentId")
            if not att_id:
                pdfs_skipped += 1
                continue

            attachment = (
                service.users()
                .messages()
                .attachments()
                .get(userId="me", messageId=msg_id, id=att_id)
                .execute()
            )
            raw = base64.urlsafe_b64decode(attachment.get("data", ""))
            dest = _unique_path(inbox, _safe_filename(filename))
            dest.write_bytes(raw)
            info = {
                "action": "downloaded",
                "message_id": msg_id,
                "subject": subject,
                "from": sender,
                "attachment": filename,
                "saved_to": str(dest),
                "bytes": len(raw),
            }
            logging.info("Downloaded Gmail PDF: %s", info)
            results.append(info)
            pdfs_downloaded += 1
            downloaded_this_email += 1

        if downloaded_this_email:
            emails_with_downloads += 1

    summary = {
        "action": "gmail_summary",
        "query": query,
        "emails_found": emails_found,
        "emails_with_downloads": emails_with_downloads,
        "pdfs_seen": pdfs_seen,
        "pdfs_skipped": pdfs_skipped,
        "pdfs_downloaded": pdfs_downloaded,
        "inbox": str(inbox),
    }
    logging.info(
        "Gmail resume: emails_found=%s, pdfs_seen=%s, pdfs_downloaded=%s, pdfs_skipped=%s",
        emails_found,
        pdfs_seen,
        pdfs_downloaded,
        pdfs_skipped,
    )
    if not pdfs_downloaded:
        logging.info("No matching Gmail PDF attachments downloaded")
    results.append(summary)
    return results
