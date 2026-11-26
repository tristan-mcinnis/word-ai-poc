#!/usr/bin/env bash
set -euo pipefail

PLIST_LABEL="local.word-ai-poc"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"
WEF_FILE="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/word-ai-poc.xml"

echo "Stopping background server..."
launchctl bootout "gui/$(id -u)/${PLIST_LABEL}" 2>/dev/null || true
rm -f "$PLIST_PATH"

echo "Removing sideloaded Word manifest..."
rm -f "$WEF_FILE"

echo "Done. Your local config/keys in ~/.config/word-poc/ were left untouched."
echo "Remove ~/.config/word-poc/ manually if you also want to delete saved settings."
