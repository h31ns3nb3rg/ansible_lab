"""Delete aged logs and archived inbox files (processed / failed)."""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any

# Keep marker files used by git / empty dirs
_KEEP_NAMES = {".gitkeep", ".DS_Store"}


def cleanup_old_files(
    directories: list[Path],
    *,
    retention_days: int = 7,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Remove files older than retention_days from the given directories.

    Age is based on filesystem mtime. Directories and keep-marker files are skipped.
    """
    if retention_days < 1:
        raise ValueError("retention_days must be >= 1")

    cutoff = time.time() - (retention_days * 24 * 60 * 60)
    deleted: list[str] = []
    skipped: list[str] = []
    errors: list[dict[str, str]] = []

    for directory in directories:
        if not directory.exists() or not directory.is_dir():
            continue
        for path in directory.iterdir():
            if path.name in _KEEP_NAMES:
                continue
            if path.is_dir():
                skipped.append(str(path))
                continue
            try:
                mtime = path.stat().st_mtime
            except OSError as exc:
                errors.append({"path": str(path), "error": str(exc)})
                continue
            if mtime >= cutoff:
                continue
            if dry_run:
                deleted.append(str(path))
                continue
            try:
                path.unlink()
                deleted.append(str(path))
            except OSError as exc:
                errors.append({"path": str(path), "error": str(exc)})

    result = {
        "action": "cleanup_dry_run" if dry_run else "cleanup",
        "retention_days": retention_days,
        "deleted_count": len(deleted),
        "deleted": deleted,
        "skipped_dirs": skipped,
        "errors": errors,
    }
    logging.info(
        "Cleanup retention=%sdays deleted=%s errors=%s%s",
        retention_days,
        len(deleted),
        len(errors),
        " (dry-run)" if dry_run else "",
    )
    for path in deleted[:20]:
        logging.info("Cleanup removed: %s", path)
    if len(deleted) > 20:
        logging.info("Cleanup removed %s more file(s)…", len(deleted) - 20)
    for err in errors:
        logging.warning("Cleanup failed %s: %s", err["path"], err["error"])
    return result
