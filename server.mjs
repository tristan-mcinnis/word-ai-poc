import https from 'node:https';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { createRequestHandler } from './lib/server-core.mjs';
import { createMcpManager } from './lib/mcp.mjs';

const PORT = 3443;
const certDir = path.join(os.homedir(), '.office-addin-dev-certs');

let options;
try {
  options = {
    key: fs.readFileSync(path.join(certDir, 'localhost.key')),
    cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
  };
} catch {
  console.error('Missing development certificate. Run ./install.sh or see README.md.');
  process.exit(1);
}

const mcpManager = createMcpManager();
const handler = createRequestHandler({ allowedHost: `localhost:${PORT}`, mcpManager });

// Listen on both loopback stacks: WebKit resolves localhost to ::1 first.
const s4 = https.createServer(options, handler).listen(PORT, '127.0.0.1', () => {
  console.log(`Office AI Editor pane: https://localhost:${PORT} (IPv4).`);
});
const s6 = https.createServer(options, handler).listen(PORT, '::1', () => {
  console.log(`Office AI Editor pane: https://localhost:${PORT} (IPv6).`);
});

function shutdown() {
  mcpManager.closeAll();
  s4.close();
  s6.close();
  process.exit(0);
}
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
process.on('exit', () => mcpManager.closeAll());
