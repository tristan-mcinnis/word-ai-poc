# Office AI Editor (Word, Excel & PowerPoint)

A local, model-agnostic AI chat and editing sidebar for **Microsoft Word, Excel, and PowerPoint on macOS**, designed to feel like **Claude in Word**, **Claude in Excel**, and **Claude in PowerPoint**.

Chat naturally about your open document, workbook, or presentation, inspect diffs and structured previews, jump directly to cited cells or slides, and apply native edits in one click. Bring your own API key from **Claude, OpenAI, Gemini, DeepSeek, Qwen, Kimi, GLM, or SiliconFlow**.

## Features

### Microsoft Word
- **Full-document & selection awareness**: Reads the open Word document, current selection, and open Word comments automatically. Switch scope between **Auto**, **Selection only**, and **Whole doc**.
- **Native Track Changes & Ask / Auto mode**: Toggle **Track** in the composer to write edits as native Word redlines (`WordApi 1.5+`). Review and click **Allow** or switch to **Auto** for non-destructive auto-apply.
- **Structured Word edits**: Replace selection (including multi-paragraph spans), insert below/above, targeted search-and-replace (`replace_match`), append to document, add margin comments, and reply in open comment threads.
- **Slash workflows (`/`)**: `/copy-edit`, `/check-doc`, `/work-comments`, `/tighten`, `/shorten`, `/rewrite`, `/plain`, `/bullets`, `/zh`, `/en`, `/summarise`, `/comment`.

### Microsoft Excel (Claude in Excel-style)
- **Multi-sheet formula & value awareness**: Reads all worksheet tabs, active & secondary sheet grids with **both formulas and calculated values** side by side (`=SUM(B2:B5) [420]`), current selection, open cell comments, and automatically flags formula errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`).
- **Clickable cell citations**: Cell and range references in chat (e.g. `` `B4` ``, `` `Sheet2!C12` ``, `` `'P&L'!D5:F5` ``) render as clickable chips that jump straight to those cells in Excel.
- **Edit highlighting**: Toggle **Highlight** in the composer to subtly tint AI-modified cells light yellow for quick auditing.
- **Structured Excel edits with mini-grid previews**: Write cell grids & dynamic formulas (`set_cells`, auto-resizing cleanly from the anchor cell), set formulas (`set_formula`), style & number-format ranges (`format_cells`), create populated worksheets (`add_sheet`), create native Excel charts (`add_chart`), sort ranges (`sort_range`), add cell comments (`excel_comment`), and clear ranges (`clear_range`, always confirmation-gated).
- **Slash workflows (`/`)**: `/explain-formula`, `/audit-model`, `/fix-errors`, `/write-formula`, `/summarize-data`, `/clean-data`, `/build-table`, `/add-chart`, `/format-table`, `/comment`.

### Microsoft PowerPoint (Claude in PowerPoint-style)
- **Full deck & Slide Master awareness**: Reads every slide, shape, and text frame across the presentation, plus the deck's **Slide Master layouts** so new slides stay on-brand with your template's fonts, colors, and placeholders.
- **Clickable slide citations**: `Slide 3` references in chat render as clickable pills that jump directly to that slide.
- **Structured PowerPoint edits with slide & diff previews**: Update shape text with word-level diff (`update_shape`), targeted text search-and-replace across a slide or deck (`replace_slide_text`), insert on-brand slides using matched template layouts (`add_slide`), add text boxes (`add_textbox`), insert native tables (`add_table`), and delete slides (`delete_slide`, always confirmation-gated).
- **Slash workflows (`/`)**: `/review-deck`, `/tighten-slide`, `/rewrite-title`, `/add-slide`, `/outline-deck`, `/speaker-notes`, `/add-table`, `/polish-copy`, `/zh`, `/en`.

### Shared UX & Privacy
- **Markdown tables & inline diffs**: Renders clean Markdown tables in chat and word/CJK inline diffs on text edits.
- **Safety guardrails**: Destructive edits (`delete_slide`, `clear_range`) are never auto-applied even in Auto mode.
- **Private & local**: Runs on `https://localhost:3443` on your Mac. API keys stay in `~/.config/word-poc/.env` (`chmod 600`) or environment variables and are never sent back to the browser or stored in git.

## Quick Install (macOS)

Requires **Node.js 18+** and **Microsoft Word, Excel, and/or PowerPoint for Mac**.

```bash
git clone https://github.com/tristan-mcinnis/word-ai-poc.git
cd word-ai-poc
./install.sh
```

The installer will:
1. Install Node dependencies
2. Set up a trusted local HTTPS dev certificate (`office-addin-dev-certs`)
3. Sideload `manifest.xml` into Word, Excel, and PowerPoint's macOS add-in folders (`wef`)
4. Prompt you to pick a default provider and save your API key locally in `~/.config/word-poc/.env` (`chmod 600`)
5. Install and start a macOS LaunchAgent so the local server runs automatically in the background

Then open (or restart) Microsoft Word, Excel, or PowerPoint and go to:
**Home → Add-ins → Developer Add-ins → AI Editor**

## API Keys

Add or update keys at any time by clicking **Keys** in the sidebar header (writes locally to `~/.config/word-poc/.env` with `chmod 600`; saved keys are never sent back to the browser), during `./install.sh`, or by editing `.env` directly.

Keys are loaded at runtime on your machine in this order:
1. Process environment variables
2. `~/.config/word-poc/.env` (created by `./install.sh` or the in-pane **Keys** button)
3. `.env` in the repository directory (gitignored)

| Provider | Environment variable | Default model |
|---|---|---|
| Claude | `ANTHROPIC_API_KEY` | `claude-sonnet-4-5` |
| OpenAI | `OPENAI_API_KEY` | `gpt-4.1-mini` |
| Gemini | `GEMINI_API_KEY` | `gemini-2.5-flash` |
| DeepSeek | `DEEPSEEK_API_KEY` | `deepseek-flash` |
| Qwen | `DASHSCOPE_API_KEY` | `qwen-plus` |
| Kimi | `KIMI_CN_API_KEY` | `kimi-k2-0905-preview` |
| GLM | `ZHIPU_API_KEY` | `glm-4.6` |
| SiliconFlow | `SILICONFLOW_API_KEY` | `deepseek-ai/DeepSeek-V3` |

## Architecture

- [`server.mjs`](server.mjs) — Thin HTTPS bootstrap (IPv4 + IPv6 loopback)
- [`lib/server-core.mjs`](lib/server-core.mjs) — Testable HTTP API handler & allowlisted static file router
- [`lib/providers.mjs`](lib/providers.mjs) — Provider registry, local `.env`/config store, and OpenAI/Anthropic SSE stream adapters
- [`lib/prompts.mjs`](lib/prompts.mjs) — Host-specific system prompts & context builders for Word, Excel, and PowerPoint
- [`lib/format.js`](lib/format.js) — Multi-host `<edit>` parser, XSS-safe Markdown (with tables & citations), and word/CJK diff engine
- [`lib/excel-utils.js`](lib/excel-utils.js) — Pure Excel A1 parsing, grid/formula coercion, workbook snapshot formatting, and error detection
- [`lib/ppt-utils.js`](lib/ppt-utils.js) — Pure PowerPoint slide spec parsing, Slide Master layout matching, placeholder classification, and deck serialization
- [`lib/hosts/`](lib/hosts/) — Clean Office.js adapters for [`word.js`](lib/hosts/word.js), [`excel.js`](lib/hosts/excel.js), and [`powerpoint.js`](lib/hosts/powerpoint.js)
- [`app.js`](app.js) — Host-agnostic sidebar UI orchestrator

## Development, Tests & Evaluations

```bash
npm test      # run unit & integration tests across Word, Excel, PowerPoint, and HTTP server
npm run eval  # run end-to-end multi-host evaluation suite (18 scenarios & quality rubrics)
npm start     # run the HTTPS server in the foreground
```

## Uninstall

```bash
./uninstall.sh
```

## License

MIT
