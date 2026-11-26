#!/usr/bin/env bash
set -euo pipefail

PLIST_LABEL="local.word-ai-poc"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"
WEF_FILES=(
  "$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/word-ai-poc.xml"
  "$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/word-ai-poc.xml"
  "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/word-ai-poc.xml"
)

echo "Stopping background server..."
launchctl bootout "gui/$(id -u)/${PLIST_LABEL}" 2>/dev/null || true
rm -f "$PLIST_PATH"

echo "Removing sideloaded Word, Excel, and PowerPoint manifests..."
for f in "${WEF_FILES[@]}"; do
  rm -f "$f"
done

echo "Done. Your local config/keys in ~/.config/word-poc/ were left untouched."
echo "Remove ~/.config/word-poc/ manually if you also want to delete saved settings."
