import https from 'node:https';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
let options;
try {
  options = {
    key: fs.readFileSync(path.join(certDir, 'localhost.key')),
    cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
  };
} catch {
  console.error('Missing development certificate. See START-HERE.md.');
  process.exit(1);
}

const files = {
  '/': 'taskpane.html',
  '/taskpane.html': 'taskpane.html',
  '/app.js': 'app.js',
  '/manifest.xml': 'manifest.xml',
};

const handler = (req, res) => {
  console.log(new Date().toISOString(), req.method, req.url);
  if (req.headers.host !== 'localhost:3443') { res.writeHead(403).end(); return; }
  const file = files[req.url];
  if (!file || req.method !== 'GET') { res.writeHead(404).end(); return; }
  res.writeHead(200, {
    'Content-Type': file.endsWith('.js') ? 'text/javascript' : file.endsWith('.xml') ? 'application/xml' : 'text/html',
    'Cache-Control': 'no-store',
  });
  res.end(fs.readFileSync(new URL(file, import.meta.url)));
};

// Listen on both loopback stacks: WebKit resolves localhost to ::1 first.
https.createServer(options, handler).listen(3443, '127.0.0.1', () => console.log('Word POC: https://localhost:3443 (IPv4).'));
https.createServer(options, handler).listen(3443, '::1', () => console.log('Word POC: https://localhost:3443 (IPv6).'));
