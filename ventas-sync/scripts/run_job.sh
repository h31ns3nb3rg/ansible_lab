#!/bin/bash
# Shared entrypoint for scheduled ventas sync jobs (launchd / cron).
# Sends a macOS Notification Center banner on success and failure.
set -u

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
  if [[ "${VENTAS_NOTIFY:-1}" != "0" ]]; then
    osascript -e 'display notification "Python / .venv missing — job did not run" with title "La Cata Ventas" subtitle "setup error"' >/dev/null 2>&1 || true
  fi
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
} >>"$LOG_FILE" 2>&1

set +e
"$PYTHON" "$ROOT/run_ventas_sync.py" --config "$ROOT/config.json" "$@" >>"$LOG_FILE" 2>&1
status=$?
set -e

{
  echo "=== ventas sync end (exit=$status) ==="
} >>"$LOG_FILE" 2>&1

# --- macOS Notification Center ---
# Set VENTAS_NOTIFY=0 to disable.
notify_mac() {
  local title="$1"
  local subtitle="$2"
  local body="$3"
  if [[ "${VENTAS_NOTIFY:-1}" == "0" ]]; then
    return 0
  fi
  if ! command -v osascript >/dev/null 2>&1; then
    return 0
  fi
  # Escape for AppleScript string literals
  local esc_title esc_sub esc_body
  esc_title=$(printf '%s' "$title" | sed 's/\\/\\\\/g; s/"/\\"/g')
  esc_sub=$(printf '%s' "$subtitle" | sed 's/\\/\\\\/g; s/"/\\"/g')
  esc_body=$(printf '%s' "$body" | sed 's/\\/\\\\/g; s/"/\\"/g')
  osascript -e "display notification \"${esc_body}\" with title \"${esc_title}\" subtitle \"${esc_sub}\"" >/dev/null 2>&1 || true
}

summary_line="$(grep -E '^Excel rows written:' "$LOG_FILE" | tail -1 || true)"
failed_line="$(grep -E '^Failed:' "$LOG_FILE" | tail -1 || true)"
queued_line="$(grep -E '^Queued \(Excel locked\):' "$LOG_FILE" | tail -1 || true)"
gmail_line="$(grep -E '^Gmail:' "$LOG_FILE" | tail -1 || true)"

detail=""
[[ -n "$summary_line" ]] && detail="$summary_line"
[[ -n "$failed_line" ]] && detail="${detail:+$detail · }$failed_line"
[[ -n "$queued_line" ]] && detail="${detail:+$detail · }$queued_line"
[[ -z "$detail" && -n "$gmail_line" ]] && detail="$gmail_line"
[[ -z "$detail" ]] && detail="see $(basename "$LOG_FILE")"

case "$status" in
  0)
    notify_mac "La Cata Ventas" "$JOB_TAG OK" "${detail}"
    ;;
  2)
    notify_mac "La Cata Ventas" "$JOB_TAG WARN" "Excel locked / queued — ${detail}"
    ;;
  *)
    notify_mac "La Cata Ventas" "$JOB_TAG FAILED" "exit ${status} — ${detail}"
    ;;
esac

exit "$status"
