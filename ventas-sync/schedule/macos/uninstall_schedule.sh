#!/bin/bash
# Remove La Cata ventas sync LaunchAgents.
set -euo pipefail

AGENTS_DIR="$HOME/Library/LaunchAgents"
LABEL_GMAIL="com.lacata.ventas.gmail"
LABEL_INBOX="com.lacata.ventas.inbox"
UID_NUM="$(id -u)"

launchctl bootout "gui/${UID_NUM}/${LABEL_GMAIL}" 2>/dev/null || true
launchctl bootout "gui/${UID_NUM}/${LABEL_INBOX}" 2>/dev/null || true
launchctl unload "$AGENTS_DIR/${LABEL_GMAIL}.plist" 2>/dev/null || true
launchctl unload "$AGENTS_DIR/${LABEL_INBOX}.plist" 2>/dev/null || true

rm -f "$AGENTS_DIR/${LABEL_GMAIL}.plist" "$AGENTS_DIR/${LABEL_INBOX}.plist"

echo "Removed LaunchAgents: ${LABEL_GMAIL}, ${LABEL_INBOX}"
