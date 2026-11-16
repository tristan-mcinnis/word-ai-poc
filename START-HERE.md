# Local Word editor POC

Current files: manifest.xml is the Word add-in; taskpane.html and app.js are the sidebar; server.mjs serves it and forwards optional rewrites to local Ollama. No cloud AI, telemetry, document storage or API keys. Office.js itself loads from Microsoft's CDN.

## Run

From this folder:

```sh
npm install
npx office-addin-dev-certs install
npm start
```

The certificate command asks to trust a local development HTTPS certificate. Review that prompt yourself. It has NOT been run by the agent. Do not change organisation policy if blocked.

## Load in Word for Mac

With Word closed, place a copy of manifest.xml in:

`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`

Create the folder if needed. Start Word yourself, open a scratch document, then look in Insert > Add-ins > My Add-ins (or the equivalent Add-ins menu in your Word version) for Local Editor POC. A restart or refresh may be needed. Managed Word can block this. If absent or blocked, ask IT rather than bypassing policy.

Select a short paragraph, Read selection, Demo, then Apply. Demo only tidies whitespace, it is NOT AI. Apply replaces the captured range, even if the cursor moves; it refuses if the original text changed. Inline formatting may be lost. Use Word Undo to revert.

Optional real AI: run your own Ollama separately on 127.0.0.1:11434, enter an installed model name, then Rewrite with local AI. No models are downloaded by this project.

## Limits

Not yet tested inside Word. No tracked changes, rich-format preservation, background edits or full-document access. Test only on a scratch document. Host policy and certificate trust remain unverified. Stop the server with Ctrl+C. Remove the copied manifest to uninstall the add-in; no Word policy changes are needed.
