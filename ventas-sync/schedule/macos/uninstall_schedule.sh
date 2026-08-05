#!/bin/bash
# Remove La Cata ventas sync LaunchAgents (combined + legacy).
set -euo pipefail

AGENTS_DIR="$HOME/Library/LaunchAgents"
LABEL_SYNC="com.lacata.ventas.sync"
LABEL_GMAIL="com.lacata.ventas.gmail"
LABEL_INBOX="com.lacata.ventas.inbox"
UID_NUM="$(id -u)"

for label in "$LABEL_SYNC" "$LABEL_GMAIL" "$LABEL_INBOX"; do
  launchctl bootout "gui/${UID_NUM}/${label}" 2>/dev/null || true
  launchctl unload "$AGENTS_DIR/${label}.plist" 2>/dev/null || true
  rm -f "$AGENTS_DIR/${label}.plist"
done

echo "Removed LaunchAgents: ${LABEL_SYNC}, ${LABEL_GMAIL}, ${LABEL_INBOX}"
