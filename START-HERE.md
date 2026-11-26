# Word pane (local POC)

A Word task pane that rewrites the selected passage with whatever model you point it at. The model list lives in `server.mjs`; the key is read from 1Password at call time through `~/skills/_lib/auth.py`, so no key is stored in the browser, in the config, or in git.

Current files: `manifest.xml` is the Word add-in; `taskpane.html`, `style.css` and `app.js` are the pane; `server.mjs` serves them on `https://localhost:3443` (both `127.0.0.1` and `::1`) and proxies model calls.

## Run

The server runs as the login agent `~/Library/LaunchAgents/co.innerchapter.word-poc.plist` (label `co.innerchapter.word-poc`). It starts at login and restarts on crash. Log: `/tmp/word-poc-server.log`.

```sh
launchctl kickstart -k gui/$(id -u)/co.innerchapter.word-poc   # restart
launchctl bootout gui/$(id -u)/co.innerchapter.word-poc         # stop
```

First-time setup, already done on this Mac: `npm install`, `npx office-addin-dev-certs install`, and a copy of `manifest.xml` in `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`.

## Use

Add-ins → Developer Add-ins → Local Editor POC. Select a paragraph in the document, press **Read selection**, pick an action or type an instruction, press **Run**. The answer streams into Result. Then **Replace selection** (refuses if the document changed underneath) or **Insert below**. Word Undo (Cmd+Z) reverts.

The provider dropdown switches the model. It writes `~/.config/word-poc/config.json` (provider and model only, never a key).

## Providers and keys

| Provider | 1Password ref |
|---|---|
| DeepSeek (default) | `op://Automation/DEEPSEEK_API_KEY/credential` |
| Kimi (CN) | `op://Automation/KIMI_CN_API_KEY/credential` |
| Zhipu GLM | `op://Automation/ZHIPU_API_KEY/credential` |
| SiliconFlow | `op://Automation/SILICONFLOW_API_KEY/credential` |
| Qwen | `op://Automation/DASHSCOPE_API_KEY/credential` |
| Gemini | `op://Automation/GEMINI_API_KEY/credential` |
| Claude | `op://Automation/ANTHROPIC_API_KEY/credential` |
| OpenAI | `op://Automation/OPENAI_API_KEY/credential` |

A provider with no item in 1Password returns a clear error naming the missing ref; nothing else breaks. An env var with the item name (`DEEPSEEK_API_KEY`) also works.

## Limits

Selection only, plain text replacement: inline formatting is lost. No tracked changes, no full-document access, no chat history. Test on a scratch document first. Client text sent to a provider leaves the Mac, so pick the provider with that in mind.
