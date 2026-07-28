#!/bin/bash
# Install LaunchAgents for La Cata ventas sync (Gmail + inbox at 1am and 11am).
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
# schedule/macos -> project root
ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
AGENTS_DIR="$HOME/Library/LaunchAgents"
LABEL_GMAIL="com.lacata.ventas.gmail"
LABEL_INBOX="com.lacata.ventas.inbox"

mkdir -p "$AGENTS_DIR" "$ROOT/logs"
chmod +x "$ROOT/scripts/run_job.sh"

if [[ ! -x "$ROOT/.venv/bin/python" ]]; then
  echo "WARNING: $ROOT/.venv/bin/python not found."
  echo "Create the venv first:"
  echo "  cd \"$ROOT\" && python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt"
fi

if [[ ! -f "$ROOT/config.json" ]]; then
  echo "ERROR: missing $ROOT/config.json"
  exit 1
fi

if [[ ! -f "$ROOT/secrets/token.json" ]]; then
  echo "WARNING: Gmail token missing ($ROOT/secrets/token.json)."
  echo "Run once before relying on the Gmail jobs:"
  echo "  cd \"$ROOT\" && source .venv/bin/activate && python run_ventas_sync.py --auth-gmail"
fi

render_plist() {
  local template="$1"
  local dest="$2"
  sed "s|__INSTALL_ROOT__|${ROOT}|g" "$template" >"$dest"
}

# Unload existing if present (ignore errors)
launchctl bootout "gui/$(id -u)/${LABEL_GMAIL}" 2>/dev/null || true
launchctl bootout "gui/$(id -u)/${LABEL_INBOX}" 2>/dev/null || true
# Older macOS fallback
launchctl unload "$AGENTS_DIR/${LABEL_GMAIL}.plist" 2>/dev/null || true
launchctl unload "$AGENTS_DIR/${LABEL_INBOX}.plist" 2>/dev/null || true

render_plist "$SCRIPT_DIR/${LABEL_GMAIL}.plist.template" "$AGENTS_DIR/${LABEL_GMAIL}.plist"
render_plist "$SCRIPT_DIR/${LABEL_INBOX}.plist.template" "$AGENTS_DIR/${LABEL_INBOX}.plist"

# Load (prefer bootstrap on modern macOS)
if launchctl bootstrap "gui/$(id -u)" "$AGENTS_DIR/${LABEL_GMAIL}.plist" 2>/dev/null; then
  launchctl bootstrap "gui/$(id -u)" "$AGENTS_DIR/${LABEL_INBOX}.plist"
else
  launchctl load "$AGENTS_DIR/${LABEL_GMAIL}.plist"
  launchctl load "$AGENTS_DIR/${LABEL_INBOX}.plist"
fi

echo
echo "Installed LaunchAgents:"
echo "  $AGENTS_DIR/${LABEL_GMAIL}.plist   → daily 01:00 & 11:00  (--fetch-gmail)"
echo "  $AGENTS_DIR/${LABEL_INBOX}.plist   → daily 01:00 & 11:00  (--no-fetch-gmail / cierres .xlsx in inbox)"
echo
echo "Project root: $ROOT"
echo "Logs:         $ROOT/logs/"
echo
echo "Verify:"
echo "  launchctl print gui/\$(id -u)/${LABEL_GMAIL} | head"
echo "  launchctl print gui/\$(id -u)/${LABEL_INBOX} | head"
echo
echo "Manual test (does not wait for schedule):"
echo "  VENTAS_JOB_TAG=manual-gmail \"$ROOT/scripts/run_job.sh\" --fetch-gmail"
echo "  VENTAS_JOB_TAG=manual-inbox \"$ROOT/scripts/run_job.sh\" --no-fetch-gmail"
echo
echo "Notes:"
echo "  • Mac must be awake (or wake) at 1am / 11am — sleep can skip a run."
echo "  • Close Excel when possible; locked files are queued to pending/."
echo "  • Drop AdControl 'Informe avanzado de cierres de caja' .xlsx into $ROOT/inbox/ before a run."
echo "  • Success and failure show a Notification Center banner (disable: VENTAS_NOTIFY=0)."
