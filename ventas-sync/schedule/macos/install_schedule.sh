#!/bin/bash
# Install one LaunchAgent: Gmail PDFs + inbox PDFs + cierres Excel (1am & 11am).
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
# schedule/macos -> project root
ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
AGENTS_DIR="$HOME/Library/LaunchAgents"
LABEL_SYNC="com.lacata.ventas.sync"
# Legacy agents removed on install
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
  echo "Run once before relying on the scheduled job:"
  echo "  cd \"$ROOT\" && source .venv/bin/activate && python run_ventas_sync.py --auth-gmail"
fi

render_plist() {
  local template="$1"
  local dest="$2"
  sed "s|__INSTALL_ROOT__|${ROOT}|g" "$template" >"$dest"
}

unload_label() {
  local label="$1"
  launchctl bootout "gui/$(id -u)/${label}" 2>/dev/null || true
  launchctl unload "$AGENTS_DIR/${label}.plist" 2>/dev/null || true
  rm -f "$AGENTS_DIR/${label}.plist"
}

# Unload combined + legacy agents
unload_label "$LABEL_SYNC"
unload_label "$LABEL_GMAIL"
unload_label "$LABEL_INBOX"

render_plist "$SCRIPT_DIR/${LABEL_SYNC}.plist.template" "$AGENTS_DIR/${LABEL_SYNC}.plist"

if launchctl bootstrap "gui/$(id -u)" "$AGENTS_DIR/${LABEL_SYNC}.plist" 2>/dev/null; then
  :
else
  launchctl load "$AGENTS_DIR/${LABEL_SYNC}.plist"
fi

echo
echo "Installed LaunchAgent:"
echo "  $AGENTS_DIR/${LABEL_SYNC}.plist   → daily 01:00 & 11:00"
echo "  One run does: Gmail LMF PDFs + inbox PDFs + cierres .xlsx"
echo
echo "Project root: $ROOT"
echo "Logs:         $ROOT/logs/"
echo
echo "Verify:"
echo "  launchctl print gui/\$(id -u)/${LABEL_SYNC} | head"
echo
echo "Manual run (same as schedule — processes Gmail + Excel + inbox PDFs):"
echo "  cd \"$ROOT\" && source .venv/bin/activate && python run_ventas_sync.py --fetch-gmail"
echo "  # or with Notification Center banner:"
echo "  VENTAS_JOB_TAG=manual-all \"$ROOT/scripts/run_job.sh\" --fetch-gmail"
echo
echo "Notes:"
echo "  • Mac must be awake (or wake) at 1am / 11am — sleep can skip a run."
echo "  • Close Excel when possible; locked files are queued to pending/."
echo "  • Drop AdControl cierres .xlsx into $ROOT/inbox/ before a run (DF/SJM)."
echo "  • Success and failure show a Notification Center banner (disable: VENTAS_NOTIFY=0)."
echo "  • Old separate gmail/inbox agents were removed if present."
