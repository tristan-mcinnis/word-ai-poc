# Word AI Editor

A local, model-agnostic AI chat and editing sidebar for **Microsoft Word on macOS**, designed to feel like Claude in Word or Gemini in Google Docs.

Chat naturally about your open document, select text to edit in place, inspect word-level diffs, and apply edits or Word comments in one click. Bring your own API key from **Claude, OpenAI, Gemini, DeepSeek, Qwen, Kimi, GLM, or SiliconFlow**.

## Features

- **Ongoing chat with full-document awareness**: Reads both the open Word document and your current selection automatically. Switch scope between **Auto (doc + selection)**, **Selection only**, and **Whole doc** with one click.
- **Structured edit cards with Diff / Preview**: Proposed edits appear as clean cards with word/CJK diffing and one-click apply:
  - Replace selection (with smart recovery if your cursor moved)
  - Insert clean paragraphs directly below or above
  - Targeted search-and-replace anywhere in the document
  - Append to the end of the document
  - Add native Word margin comments
  - **Apply all** when a reply proposes multiple edits
- **Quick actions without clutter**: Press `/` in the composer or click **Actions** for one-click transforms (`/tighten`, `/shorten`, `/rewrite`, `/plain`, `/bullets`, `/zh`, `/en`, `/summarise`, `/critique`, `/comment`).
- **Private & local**: Runs on `https://localhost:3443` on your Mac. API keys stay in `~/.config/word-poc/.env` (or your environment) and are never sent to the browser or stored in git.

## Quick Install (macOS)

Requires **Node.js 18+** and **Microsoft Word for Mac**.

```bash
git clone https://github.com/tristan-mcinnis/word-ai-poc.git
cd word-ai-poc
./install.sh
```

The installer will:
1. Install Node dependencies
2. Set up a trusted local HTTPS dev certificate for Word (`office-addin-dev-certs`)
3. Sideload `manifest.xml` into Word's local add-ins folder
4. Prompt you to pick a default provider and save your API key locally in `~/.config/word-poc/.env` (`chmod 600`)
5. Install and start a macOS LaunchAgent so the local server runs automatically in the background

Then open (or restart) Microsoft Word and go to:
**Home → Add-ins → Developer Add-ins → Local Editor POC**

## API Keys

Keys are loaded at runtime on your machine in this order:
1. Process environment variables
2. `~/.config/word-poc/.env` (created by `./install.sh`)
3. `.env` in the repository directory (gitignored)

| Provider | Environment variable | Default model |
|---|---|---|
| Claude | `ANTHROPIC_API_KEY` | `claude-sonnet-4-5` |
| OpenAI | `OPENAI_API_KEY` | `gpt-4.1-mini` |
| Gemini | `GEMINI_API_KEY` | `gemini-2.5-flash` |
| DeepSeek | `DEEPSEEK_API_KEY` | `deepseek-chat` |
| Qwen | `DASHSCOPE_API_KEY` | `qwen-plus` |
| Kimi | `KIMI_CN_API_KEY` | `kimi-k2-0905-preview` |
| GLM | `ZHIPU_API_KEY` | `glm-4.6` |
| SiliconFlow | `SILICONFLOW_API_KEY` | `deepseek-ai/DeepSeek-V3` |

Example `~/.config/word-poc/.env`:

```env
ANTHROPIC_API_KEY=sk-ant-...
DEEPSEEK_API_KEY=sk-...
```

## Development & Tests

```bash
npm test      # run parser, markdown, and diff unit tests
npm start     # run the HTTPS server in the foreground
```

## Uninstall

```bash
./uninstall.sh
```

## Limits & Privacy Note

- Plain-text paragraph replacements do not yet preserve mixed inline character formatting (bold/italics inside a replaced paragraph). Word Undo (`Cmd+Z`) reverts any edit, and Word's native Track Changes works when enabled in the Review tab.
- Document text you chat about is sent directly from your Mac to whichever model provider you select.

## License

MIT
