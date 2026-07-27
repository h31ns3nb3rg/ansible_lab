#!/bin/bash
# Shared entrypoint for scheduled ventas sync jobs (launchd / cron).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

LOG_DIR="$ROOT/logs"
mkdir -p "$LOG_DIR"

if [[ -x "$ROOT/.venv/bin/python" ]]; then
  PYTHON="$ROOT/.venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON="$(command -v python3)"
else
  echo "ERROR: no Python found (expected $ROOT/.venv/bin/python)" >&2
  exit 127
fi

STAMP="$(date '+%Y%m%d_%H%M%S')"
JOB_TAG="${VENTAS_JOB_TAG:-manual}"
LOG_FILE="$LOG_DIR/scheduled_${JOB_TAG}_${STAMP}.log"

{
  echo "=== ventas sync start ==="
  echo "time: $(date '+%Y-%m-%d %H:%M:%S %Z')"
  echo "root: $ROOT"
  echo "python: $PYTHON"
  echo "args: $*"
  echo "========================="
  "$PYTHON" "$ROOT/run_ventas_sync.py" --config "$ROOT/config.json" "$@"
  status=$?
  echo "=== ventas sync end (exit=$status) ==="
  exit "$status"
} >>"$LOG_FILE" 2>&1
