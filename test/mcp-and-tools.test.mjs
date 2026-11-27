import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import { parseJsonc, discoverMcpServers, createMcpManager } from '../lib/mcp.mjs';
import { evaluateMathExpression, getBuiltinTools, executeToolCall } from '../lib/tools.mjs';

function makeTempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-mcp-'));
}

test('parseJsonc strips comments and trailing commas while preserving URLs inside strings', () => {
  const raw = `{
    // OpenCode MCP config
    "mcp": {
      "docs": {
        "url": "https://example.com/mcp", // inline comment
      },
    },
    /* block comment */
  }`;
  const parsed = parseJsonc(raw);
  assert.equal(parsed.mcp.docs.url, 'https://example.com/mcp');
});

test('discoverMcpServers merges Word-POC, Pi, Claude, and OpenCode configs with proper priority', () => {
  const home = makeTempDir();
  try {
    fs.mkdirSync(path.join(home, '.config', 'word-poc'), { recursive: true });
    fs.mkdirSync(path.join(home, '.pi', 'agent'), { recursive: true });
    fs.mkdirSync(path.join(home, '.config', 'opencode'), { recursive: true });

    fs.writeFileSync(
      path.join(home, '.config', 'word-poc', 'mcp.json'),
      JSON.stringify({ mcpServers: { shared: { url: 'http://localhost:9001/mcp' } } }),
    );
    fs.writeFileSync(
      path.join(home, '.pi', 'agent', 'mcp.json'),
      JSON.stringify({
        mcpServers: {
          shared: { url: 'http://localhost:9999/shadowed' },
          pi_tools: { command: 'node', args: ['server.js'] },
        },
      }),
    );
    fs.writeFileSync(
      path.join(home, '.claude.json'),
      JSON.stringify({ mcpServers: { claude_fs: { command: 'npx', args: ['-y', 'fs-mcp'] } } }),
    );
    fs.writeFileSync(
      path.join(home, '.config', 'opencode', 'opencode.jsonc'),
      `{ // OpenCode config
        "mcp": {
          "oc_search": { "type": "local", "command": ["node", "search-mcp.mjs"] },
          "disabled_one": { "command": "node", "enabled": false },
        }
      }`,
    );

    const servers = discoverMcpServers({ homeDir: home, env: {} });
    assert.equal(servers.length, 4);
    assert.equal(servers.find(s => s.name === 'shared').url, 'http://localhost:9001/mcp');
    assert.equal(servers.find(s => s.name === 'pi_tools').transport, 'stdio');
    assert.equal(servers.find(s => s.name === 'claude_fs').command, 'npx');
    assert.deepEqual(servers.find(s => s.name === 'oc_search').args, ['search-mcp.mjs']);
  } finally {
    fs.rmSync(home, { recursive: true, force: true });
  }
});

test('MCP stdio client connects to real subprocess server, lists tools, calls tool, and closes cleanly', async () => {
  const dir = makeTempDir();
  const mockServerScript = path.join(dir, 'mock-mcp.mjs');
  fs.writeFileSync(
    mockServerScript,
    `
    import readline from 'node:readline';
    const rl = readline.createInterface({ input: process.stdin });
    rl.on('line', line => {
      if (!line.trim()) return;
      const msg = JSON.parse(line);
      if (msg.method === 'initialize') {
        process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: msg.id, result: { protocolVersion: '2025-03-26', capabilities: { tools: {} } } }) + '\\n');
      } else if (msg.method === 'tools/list') {
        process.stdout.write(JSON.stringify({
          jsonrpc: '2.0',
          id: msg.id,
          result: {
            tools: [{
              name: 'lookup_fx',
              description: 'Return FX rate',
              annotations: { readOnlyHint: true },
              inputSchema: { type: 'object', properties: { pair: { type: 'string' } } }
            }]
          }
        }) + '\\n');
      } else if (msg.method === 'tools/call') {
        process.stdout.write(JSON.stringify({
          jsonrpc: '2.0',
          id: msg.id,
          result: { content: [{ type: 'text', text: 'FX ' + msg.params.arguments.pair + ' = 1.08' }] }
        }) + '\\n');
      }
    });
    `,
  );

  const cfgPath = path.join(dir, 'mcp.json');
  fs.writeFileSync(cfgPath, JSON.stringify({
    mcpServers: {
      fx: { command: process.execPath, args: [mockServerScript] },
    },
  }));

  const mgr = createMcpManager({ homeDir: dir, configDir: dir, configPaths: [cfgPath] });
  try {
    const tools = await mgr.getActiveTools();
    assert.equal(tools.length, 1);
    assert.equal(tools[0].name, 'mcp__fx__lookup_fx');
    assert.equal(tools[0].readOnly, true);

    const callRes = await mgr.callTool('fx', 'lookup_fx', { pair: 'EUR/USD' });
    assert.equal(callRes.isError, false);
    assert.equal(callRes.text, 'FX EUR/USD = 1.08');
  } finally {
    mgr.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('MCP Streamable HTTP client handles JSON-RPC + SSE responses cleanly', async () => {
  const dir = makeTempDir();
  const httpSrv = http.createServer(async (req, res) => {
    const chunks = [];
    for await (const c of req) chunks.push(c);
    const msg = JSON.parse(Buffer.concat(chunks).toString('utf8'));
    res.setHeader('Mcp-Session-Id', 'sess-xyz');
    if (msg.method === 'initialize') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', id: msg.id, result: { protocolVersion: '2025-03-26', capabilities: {} } }));
    } else if (msg.method === 'notifications/initialized') {
      res.writeHead(202).end();
    } else if (msg.method === 'tools/list') {
      res.writeHead(200, { 'Content-Type': 'text/event-stream' });
      res.end(`data: ${JSON.stringify({
        jsonrpc: '2.0',
        id: msg.id,
        result: { tools: [{ name: 'ping', description: 'Ping', annotations: { readOnlyHint: true } }] },
      })}\n\n`);
    } else if (msg.method === 'tools/call') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', id: msg.id, result: { content: [{ type: 'text', text: 'pong:' + (req.headers['mcp-session-id'] || '') }] } }));
    }
  });
  await new Promise(r => httpSrv.listen(0, '127.0.0.1', r));
  const { port } = httpSrv.address();

  const cfgPath = path.join(dir, 'mcp.json');
  fs.writeFileSync(cfgPath, JSON.stringify({
    mcpServers: { remote: { url: `http://127.0.0.1:${port}/mcp` } },
  }));

  const mgr = createMcpManager({ homeDir: dir, configDir: dir, configPaths: [cfgPath] });
  try {
    const tools = await mgr.getActiveTools();
    assert.equal(tools.length, 1);
    assert.equal(tools[0].name, 'mcp__remote__ping');

    const res = await mgr.callTool('remote', 'ping', {});
    assert.equal(res.text, 'pong:sess-xyz');
  } finally {
    mgr.closeAll();
    await new Promise(r => httpSrv.close(r));
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('evaluateMathExpression and built-in tools execute safely without eval and block traversal', async () => {
  assert.equal(evaluateMathExpression('(145000 - 120000) / 120000'), 25000 / 120000);
  assert.equal(evaluateMathExpression('round(avg(10, 20, 35), 2)'), 21.67);
  assert.equal(evaluateMathExpression('2^10 + sqrt(81)'), 1033);
  assert.throws(() => evaluateMathExpression('process.exit(1)'));
  assert.throws(() => evaluateMathExpression('1 / 0'));

  const dir = makeTempDir();
  try {
    fs.mkdirSync(path.join(dir, 'references'), { recursive: true });
    fs.writeFileSync(path.join(dir, 'references', 'rubric.md'), 'Always cite exact cells.');

    const enabledSkills = [{ name: 'audit', dirPath: dir, resources: ['references/rubric.md'] }];
    const builtins = getBuiltinTools({ enabledSkills });
    assert.equal(builtins.length, 2);

    const calcRes = await executeToolCall({ toolDef: builtins[0], args: { expression: 'sum(10, 20, 30)' } });
    assert.equal(calcRes.ok, true);
    assert.equal(calcRes.text, '60');

    const readOk = await executeToolCall({
      toolDef: builtins[1],
      args: { skill: 'audit', path: 'references/rubric.md' },
      enabledSkills,
    });
    assert.equal(readOk.ok, true);
    assert.match(readOk.text, /Always cite exact cells/);

    const readTraversal = await executeToolCall({
      toolDef: builtins[1],
      args: { skill: 'audit', path: '../../etc/passwd' },
      enabledSkills,
    });
    assert.equal(readTraversal.ok, false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
