import https from 'node:https';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { createRequestHandler } from './lib/server-core.mjs';

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

const handler = createRequestHandler({ allowedHost: `localhost:${PORT}` });

// Listen on both loopback stacks: WebKit resolves localhost to ::1 first.
https.createServer(options, handler).listen(PORT, '127.0.0.1', () => {
  console.log(`Office AI Editor pane: https://localhost:${PORT} (IPv4).`);
});
https.createServer(options, handler).listen(PORT, '::1', () => {
  console.log(`Office AI Editor pane: https://localhost:${PORT} (IPv6).`);
});
