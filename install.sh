#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG_DIR="$HOME/.config/word-poc"
ENV_FILE="$CONFIG_DIR/.env"
CONFIG_FILE="$CONFIG_DIR/config.json"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
PLIST_LABEL="local.word-ai-poc"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"

echo ""
echo "Word AI Editor — macOS Installer"
echo "================================"
echo ""

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "This installer currently supports macOS Microsoft Word."
  echo "On other platforms, run 'npm install && npx office-addin-dev-certs install && npm start' and sideload manifest.xml."
  exit 1
fi

NODE_BIN="$(command -v node || true)"
if [[ -z "$NODE_BIN" ]]; then
  echo "Error: Node.js (>= 18) is required. Install it from https://nodejs.org or via 'brew install node'."
  exit 1
fi

NODE_MAJOR="$("$NODE_BIN" -p 'process.versions.node.split(".")[0]')"
if (( NODE_MAJOR < 18 )); then
  echo "Error: Node.js 18+ is required (found $("$NODE_BIN" -v))."
  exit 1
fi

echo "1/5  Installing dependencies..."
(cd "$REPO_DIR" && npm install --no-audit --no-fund)

echo ""
echo "2/5  Installing local HTTPS dev certificate for Word..."
echo "     (macOS may ask for your login password to trust localhost.)"
(cd "$REPO_DIR" && npx office-addin-dev-certs install)

echo ""
echo "3/5  Sideloading Word add-in manifest..."
mkdir -p "$WEF_DIR"
cp "$REPO_DIR/manifest.xml" "$WEF_DIR/word-ai-poc.xml"
echo "     Copied to $WEF_DIR/word-ai-poc.xml"

echo ""
echo "4/5  Configuring your AI provider & API key (stored locally in $ENV_FILE, chmod 600)..."
mkdir -p "$CONFIG_DIR"
chmod 700 "$CONFIG_DIR"
touch "$ENV_FILE"
chmod 600 "$ENV_FILE"

echo "     Choose your default provider:"
echo "       1) Claude (Anthropic)      [ANTHROPIC_API_KEY]"
echo "       2) OpenAI                  [OPENAI_API_KEY]"
echo "       3) Gemini (Google)         [GEMINI_API_KEY]"
echo "       4) DeepSeek                [DEEPSEEK_API_KEY]"
echo "       5) Qwen (DashScope)        [DASHSCOPE_API_KEY]"
echo "       6) Kimi (Moonshot)         [KIMI_CN_API_KEY]"
echo "       7) GLM (Zhipu)             [ZHIPU_API_KEY]"
echo "       8) SiliconFlow             [SILICONFLOW_API_KEY]"
echo "       0) Skip (configure later)"
echo ""
read -r -p "     Select [1-8, or 0 to skip]: " CHOICE || CHOICE="0"

PROVIDER=""
ENV_KEY=""
case "${CHOICE:-0}" in
  1) PROVIDER="anthropic";   ENV_KEY="ANTHROPIC_API_KEY" ;;
  2) PROVIDER="openai";      ENV_KEY="OPENAI_API_KEY" ;;
  3) PROVIDER="gemini";      ENV_KEY="GEMINI_API_KEY" ;;
  4) PROVIDER="deepseek";    ENV_KEY="DEEPSEEK_API_KEY" ;;
  5) PROVIDER="dashscope";   ENV_KEY="DASHSCOPE_API_KEY" ;;
  6) PROVIDER="moonshot";    ENV_KEY="KIMI_CN_API_KEY" ;;
  7) PROVIDER="zhipu";       ENV_KEY="ZHIPU_API_KEY" ;;
  8) PROVIDER="siliconflow"; ENV_KEY="SILICONFLOW_API_KEY" ;;
  *) PROVIDER="";            ENV_KEY="" ;;
esac

if [[ -n "$PROVIDER" ]]; then
  cat > "$CONFIG_FILE" <<EOF
{
  "provider": "$PROVIDER"
}
EOF
  chmod 600 "$CONFIG_FILE"

  read -r -s -p "     Paste your $ENV_KEY (hidden, or press Enter to skip): " API_KEY || API_KEY=""
  echo ""
  if [[ -n "$API_KEY" ]]; then
    # Replace existing entry for ENV_KEY if present, then append
    grep -v "^${ENV_KEY}=" "$ENV_FILE" > "${ENV_FILE}.tmp" || true
    printf '%s=%s\n' "$ENV_KEY" "$API_KEY" >> "${ENV_FILE}.tmp"
    mv "${ENV_FILE}.tmp" "$ENV_FILE"
    chmod 600 "$ENV_FILE"
    echo "     Saved $ENV_KEY to $ENV_FILE."
  else
    echo "     Skipped key entry. Add ${ENV_KEY}=... to $ENV_FILE whenever you are ready."
  fi
fi

echo ""
echo "5/5  Installing background login agent ($PLIST_LABEL)..."
mkdir -p "$HOME/Library/LaunchAgents"
cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${PLIST_LABEL}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${NODE_BIN}</string>
    <string>${REPO_DIR}/server.mjs</string>
  </array>
  <key>WorkingDirectory</key>
  <string>${REPO_DIR}</string>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>/tmp/word-ai-poc.log</string>
  <key>StandardErrorPath</key>
  <string>/tmp/word-ai-poc.log</string>
</dict>
</plist>
EOF

launchctl bootout "gui/$(id -u)/${PLIST_LABEL}" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST_PATH"
launchctl kickstart -k "gui/$(id -u)/${PLIST_LABEL}"

sleep 1
if curl -skf https://localhost:3443/api/config >/dev/null 2>&1; then
  echo "     Server is live at https://localhost:3443"
else
  echo "     Server started. Check /tmp/word-ai-poc.log if it does not respond."
fi

echo ""
echo "Done! Next steps:"
echo "  1. Open (or restart) Microsoft Word."
echo "  2. Go to Home -> Add-ins -> Developer Add-ins -> Local Editor POC."
echo "  3. Chat about your document, or select text to edit it in place."
echo ""
echo "To add more API keys later, edit: $ENV_FILE"
echo "To uninstall, run: $REPO_DIR/uninstall.sh"
echo ""
