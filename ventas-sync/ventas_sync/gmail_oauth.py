"""Gmail OAuth (read-only) helpers for desktop apps."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Read-only is enough to list messages and download attachments.
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def _resolve(base: Path, maybe_relative: str) -> Path:
    path = Path(maybe_relative).expanduser()
    return path if path.is_absolute() else (base / path).resolve()


def _run_consent_flow(credentials_path: Path, *, open_browser: bool) -> Credentials:
    if not credentials_path.exists():
        raise FileNotFoundError(
            "Missing Google OAuth client file.\n"
            f"Expected: {credentials_path}\n"
            "Create a Desktop OAuth client in Google Cloud, enable Gmail API, "
            "download JSON, and save it as credentials.json (see README)."
        )
    flow = InstalledAppFlow.from_client_secrets_file(
        str(credentials_path), GMAIL_SCOPES
    )
    if open_browser:
        return flow.run_local_server(port=0, prompt="consent")
    return flow.run_console()


def get_gmail_service(
    credentials_path: Path,
    token_path: Path,
    *,
    open_browser: bool = True,
):
    """Return an authenticated Gmail API service.

    First run opens a browser for Google consent and writes token_path.
    Later runs refresh the token automatically. If the refresh token was
    revoked/expired, deletes token_path and opens the browser again.
    """
    creds: Credentials | None = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), GMAIL_SCOPES)

    if not creds or not creds.valid:
        need_consent = True
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                need_consent = False
            except RefreshError as exc:
                logging.warning(
                    "Gmail token expired or revoked (%s). "
                    "Removing %s — re-authorize in the browser.",
                    exc,
                    token_path,
                )
                if token_path.exists():
                    token_path.unlink()
                creds = None

        if need_consent:
            creds = _run_consent_flow(credentials_path, open_browser=open_browser)

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def gmail_paths_from_config(config: dict[str, Any], base_dir: Path) -> tuple[Path, Path]:
    gmail = config.get("gmail") or {}
    credentials = _resolve(base_dir, gmail.get("credentials_path", "secrets/credentials.json"))
    token = _resolve(base_dir, gmail.get("token_path", "secrets/token.json"))
    return credentials, token


def describe_auth_status(credentials_path: Path, token_path: Path) -> dict[str, Any]:
    status = {
        "credentials_exists": credentials_path.exists(),
        "credentials_path": str(credentials_path),
        "token_exists": token_path.exists(),
        "token_path": str(token_path),
        "scopes": GMAIL_SCOPES,
    }
    if token_path.exists():
        try:
            data = json.loads(token_path.read_text(encoding="utf-8"))
            status["has_refresh_token"] = bool(data.get("refresh_token"))
        except json.JSONDecodeError:
            status["has_refresh_token"] = False
    return status
