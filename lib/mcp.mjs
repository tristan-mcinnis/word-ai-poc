/* Pure Node.js Model Context Protocol (MCP) client & multi-harness config discovery.
   Supports stdio and Streamable HTTP JSON-RPC 2.0 transports with zero npm dependencies. */

import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawn } from 'node:child_process';
import { toDisplayPath } from './skills.mjs';

const MCP_PROTOCOL_VERSION = '2025-03-26';
const DEFAULT_CALL_TIMEOUT_MS = 30000;
const MAX_TOOL_OUTPUT_CHARS = 16000;

function expandHome(rawPath, homeDir) {
  return path.resolve(String(rawPath || '').trim().replace(/^~(?=\/|$)/, homeDir));
}

/**
 * Strip // and /* *\/ comments & trailing commas safely (respecting quoted strings)
 * so OpenCode .jsonc / .json configs parse cleanly.
 */
export function parseJsonc(raw) {
  const src = String(raw || '');
  let out = '';
  let inStr = false;
  let escaped = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    const next = src[i + 1];
    if (inStr) {
      out += ch;
      if (escaped) {
        escaped = false;
      } else if (ch === '\\') {
        escaped = true;
      } else if (ch === '"') {
        inStr = false;
      }
      continue;
    }
    if (ch === '"') {
      inStr = true;
      out += ch;
      continue;
    }
    if (ch === '/' && next === '/') {
      while (i < src.length && src[i] !== '\n') i++;
      out += '\n';
      continue;
    }
    if (ch === '/' && next === '*') {
      i += 2;
      while (i < src.length && !(src[i] === '*' && src[i + 1] === '/')) i++;
      i++;
      continue;
    }
    out += ch;
  }
  const noTrailing = out.replace(/,\s*([}\]])/g, '$1');
  return JSON.parse(noTrailing);
}

export function normalizeServerName(raw) {
  return String(raw || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 48);
}

/**
 * Normalize a server entry from standard MCP json, Claude json, Pi mcp.json, or OpenCode opencode.json.
 */
function normalizeServerEntry(name, raw, source) {
  if (!raw || typeof raw !== 'object') return null;
  if (raw.enabled === false || raw.disabled === true) return null;

  const cleanName = normalizeServerName(name);
  if (!cleanName) return null;

  // Remote HTTP / SSE server
  const url = typeof raw.url === 'string' && raw.url.trim()
    ? raw.url.trim()
    : typeof raw.serverUrl === 'string' && raw.serverUrl.trim()
      ? raw.serverUrl.trim()
      : '';
  if (url && /^https?:\/\//i.test(url)) {
    return {
      name: cleanName,
      transport: 'http',
      url,
      headers: raw.headers && typeof raw.headers === 'object' ? { ...raw.headers } : {},
      source,
    };
  }

  // Stdio server: either { command: "npx", args: [...] } or OpenCode { command: ["npx", ...] }
  let command = '';
  let args = [];
  if (Array.isArray(raw.command) && raw.command.length > 0) {
    command = String(raw.command[0]).trim();
    args = raw.command.slice(1).map(String);
  } else if (typeof raw.command === 'string' && raw.command.trim()) {
    command = raw.command.trim();
    args = Array.isArray(raw.args) ? raw.args.map(String) : [];
  }

  if (!command) return null;
  const envObj = raw.env || raw.environment;
  return {
    name: cleanName,
    transport: 'stdio',
    command,
    args,
    env: envObj && typeof envObj === 'object' ? { ...envObj } : {},
    source,
  };
}

export function defaultMcpConfigPaths({
  homeDir = os.homedir(),
  configDir = path.join(homeDir, '.config', 'word-poc'),
  env = process.env,
} = {}) {
  const piDir = env.PI_CODING_AGENT_DIR && String(env.PI_CODING_AGENT_DIR).trim()
    ? expandHome(env.PI_CODING_AGENT_DIR, homeDir)
    : path.join(homeDir, '.pi', 'agent');
  const openCodeDir = env.OPENCODE_CONFIG_DIR && String(env.OPENCODE_CONFIG_DIR).trim()
    ? expandHome(env.OPENCODE_CONFIG_DIR, homeDir)
    : path.join(homeDir, '.config', 'opencode');

  const paths = [
    path.join(configDir, 'mcp.json'),
    path.join(piDir, 'mcp.json'),
    path.join(homeDir, '.claude', 'mcp.json'),
    path.join(homeDir, '.claude.json'),
    path.join(openCodeDir, 'opencode.jsonc'),
    path.join(openCodeDir, 'opencode.json'),
    path.join(homeDir, '.config', 'opencode', 'opencode.jsonc'),
    path.join(homeDir, '.config', 'opencode', 'opencode.json'),
  ];
  return [...new Set(paths)];
}

export function discoverMcpServers({
  homeDir = os.homedir(),
  configDir = path.join(homeDir, '.config', 'word-poc'),
  env = process.env,
  configPaths,
} = {}) {
  const paths = configPaths || defaultMcpConfigPaths({ homeDir, configDir, env });
  const seen = new Map();

  for (const cfgPath of paths) {
    let parsed;
    try {
      if (!fs.statSync(cfgPath).isFile()) continue;
      parsed = parseJsonc(fs.readFileSync(cfgPath, 'utf8'));
    } catch {
      continue;
    }
    if (!parsed || typeof parsed !== 'object') continue;

    const map = parsed.mcpServers || parsed.mcp || null;
    if (!map || typeof map !== 'object') continue;

    const displaySrc = toDisplayPath(cfgPath, homeDir);
    for (const [name, entry] of Object.entries(map)) {
      const normalized = normalizeServerEntry(name, entry, displaySrc);
      if (normalized && !seen.has(normalized.name)) {
        seen.set(normalized.name, normalized);
      }
    }
  }

  return [...seen.values()];
}

function formatMcpToolResult(result) {
  if (!result) return '';
  if (Array.isArray(result.content)) {
    const text = result.content
      .map(block => {
        if (typeof block === 'string') return block;
        if (block?.type === 'text' && typeof block.text === 'string') return block.text;
        if (block?.type === 'resource' && block.resource?.text) return block.resource.text;
        return JSON.stringify(block);
      })
      .join('\n');
    return text.slice(0, MAX_TOOL_OUTPUT_CHARS);
  }
  if (typeof result === 'string') return result.slice(0, MAX_TOOL_OUTPUT_CHARS);
  return JSON.stringify(result).slice(0, MAX_TOOL_OUTPUT_CHARS);
}

/**
 * Stdio JSON-RPC 2.0 MCP connection.
 */
class StdioMcpClient {
  constructor(serverConfig) {
    this.config = serverConfig;
    this.proc = null;
    this.nextId = 1;
    this.pending = new Map();
    this.buf = '';
    this.initialized = false;
    this.tools = [];
  }

  async connect(timeoutMs = DEFAULT_CALL_TIMEOUT_MS) {
    if (this.initialized && this.proc && !this.proc.killed) return;
    this.close();

    this.proc = spawn(this.config.command, this.config.args || [], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: { ...process.env, ...(this.config.env || {}) },
    });

    this.proc.stdout.setEncoding('utf8');
    this.proc.stdout.on('data', chunk => {
      this.buf += chunk;
      let idx;
      while ((idx = this.buf.indexOf('\n')) >= 0) {
        const line = this.buf.slice(0, idx).replace(/\r$/, '').trim();
        this.buf = this.buf.slice(idx + 1);
        if (!line) continue;
        try {
          const msg = JSON.parse(line);
          if (msg && msg.id !== undefined && this.pending.has(msg.id)) {
            const p = this.pending.get(msg.id);
            this.pending.delete(msg.id);
            clearTimeout(p.timer);
            if (msg.error) p.reject(new Error(msg.error.message || JSON.stringify(msg.error)));
            else p.resolve(msg.result);
          }
        } catch {
          // Ignore non-JSON log lines written to stdout by noisy servers
        }
      }
    });

    const failAll = err => {
      for (const p of this.pending.values()) {
        clearTimeout(p.timer);
        p.reject(err);
      }
      this.pending.clear();
      this.initialized = false;
    };

    this.proc.on('error', err => failAll(err));
    this.proc.on('exit', code => failAll(new Error(`MCP server "${this.config.name}" exited (code ${code})`)));

    await this.request('initialize', {
      protocolVersion: MCP_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: 'office-ai-editor', version: '1.2.0' },
    }, timeoutMs);

    this.notify('notifications/initialized', {});
    this.initialized = true;
  }

  notify(method, params) {
    if (!this.proc || !this.proc.stdin.writable) return;
    this.proc.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, params }) + '\n');
  }

  request(method, params, timeoutMs = DEFAULT_CALL_TIMEOUT_MS, signal = null) {
    return new Promise((resolve, reject) => {
      if (!this.proc || !this.proc.stdin.writable) {
        reject(new Error(`MCP server "${this.config.name}" is not running.`));
        return;
      }
      const id = this.nextId++;
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`MCP request "${method}" timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      const onAbort = () => {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(new Error('Aborted'));
      };
      if (signal) {
        if (signal.aborted) {
          onAbort();
          return;
        }
        signal.addEventListener('abort', onAbort, { once: true });
      }

      this.pending.set(id, {
        timer,
        resolve: val => {
          if (signal) signal.removeEventListener('abort', onAbort);
          resolve(val);
        },
        reject: err => {
          if (signal) signal.removeEventListener('abort', onAbort);
          reject(err);
        },
      });

      this.proc.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    });
  }

  async listTools(timeoutMs = DEFAULT_CALL_TIMEOUT_MS) {
    await this.connect(timeoutMs);
    const res = await this.request('tools/list', {}, timeoutMs);
    this.tools = Array.isArray(res?.tools) ? res.tools : [];
    return this.tools;
  }

  async callTool(name, args, { timeoutMs = DEFAULT_CALL_TIMEOUT_MS, signal = null } = {}) {
    await this.connect(timeoutMs);
    const res = await this.request('tools/call', { name, arguments: args || {} }, timeoutMs, signal);
    return {
      isError: Boolean(res?.isError),
      text: formatMcpToolResult(res),
    };
  }

  close() {
    for (const p of this.pending.values()) {
      clearTimeout(p.timer);
      p.reject(new Error('MCP client closed'));
    }
    this.pending.clear();
    if (this.proc) {
      try { this.proc.kill('SIGTERM'); } catch {}
      this.proc = null;
    }
    this.initialized = false;
  }
}

/**
 * Streamable HTTP JSON-RPC 2.0 MCP client.
 */
class HttpMcpClient {
  constructor(serverConfig, fetchImpl = fetch) {
    this.config = serverConfig;
    this.fetchImpl = fetchImpl;
    this.nextId = 1;
    this.sessionId = '';
    this.initialized = false;
    this.tools = [];
  }

  async rpc(method, params, { timeoutMs = DEFAULT_CALL_TIMEOUT_MS, signal = null, isNotification = false } = {}) {
    const headers = {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
      ...(this.config.headers || {}),
    };
    if (this.sessionId) headers['Mcp-Session-Id'] = this.sessionId;

    const payload = isNotification
      ? { jsonrpc: '2.0', method, params }
      : { jsonrpc: '2.0', id: this.nextId++, method, params };

    const signals = [AbortSignal.timeout(timeoutMs)];
    if (signal) signals.push(signal);
    const combinedSignal = signals.length > 1 && AbortSignal.any ? AbortSignal.any(signals) : signals[0];

    const res = await this.fetchImpl(this.config.url, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
      signal: combinedSignal,
    });

    const sid = res.headers?.get?.('mcp-session-id');
    if (sid) this.sessionId = sid;

    if (isNotification) return null;
    if (!res.ok) {
      throw new Error(`HTTP MCP server "${this.config.name}" returned ${res.status}`);
    }

    const ct = String(res.headers?.get?.('content-type') || '').toLowerCase();
    const text = await res.text();

    let msg;
    if (ct.includes('text/event-stream')) {
      // Extract last JSON-RPC data frame from SSE response
      const dataLines = text
        .split(/\r?\n/)
        .filter(l => l.startsWith('data:'))
        .map(l => l.slice(5).trim())
        .filter(Boolean);
      for (const line of dataLines) {
        try {
          const parsed = JSON.parse(line);
          if (parsed && (parsed.result !== undefined || parsed.error !== undefined)) {
            msg = parsed;
          }
        } catch {}
      }
    } else {
      msg = JSON.parse(text);
    }

    if (!msg) throw new Error(`Empty JSON-RPC response from "${this.config.name}"`);
    if (msg.error) throw new Error(msg.error.message || JSON.stringify(msg.error));
    return msg.result;
  }

  async connect(timeoutMs = DEFAULT_CALL_TIMEOUT_MS) {
    if (this.initialized) return;
    await this.rpc('initialize', {
      protocolVersion: MCP_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: 'office-ai-editor', version: '1.2.0' },
    }, { timeoutMs });
    await this.rpc('notifications/initialized', {}, { timeoutMs, isNotification: true }).catch(() => {});
    this.initialized = true;
  }

  async listTools(timeoutMs = DEFAULT_CALL_TIMEOUT_MS) {
    await this.connect(timeoutMs);
    const res = await this.rpc('tools/list', {}, { timeoutMs });
    this.tools = Array.isArray(res?.tools) ? res.tools : [];
    return this.tools;
  }

  async callTool(name, args, { timeoutMs = DEFAULT_CALL_TIMEOUT_MS, signal = null } = {}) {
    await this.connect(timeoutMs);
    const res = await this.rpc('tools/call', { name, arguments: args || {} }, { timeoutMs, signal });
    return {
      isError: Boolean(res?.isError),
      text: formatMcpToolResult(res),
    };
  }

  close() {
    this.initialized = false;
    this.sessionId = '';
  }
}

export function createMcpManager({
  homeDir = os.homedir(),
  configDir = path.join(homeDir, '.config', 'word-poc'),
  env = process.env,
  configPaths,
  fetchImpl = fetch,
} = {}) {
  const localMcpFile = path.join(configDir, 'mcp.json');
  const clients = new Map();
  const statusMap = new Map(); // serverName -> { connected, error, toolCount, tools }

  function listConfiguredServers() {
    return discoverMcpServers({ homeDir, configDir, env, configPaths });
  }

  function getClient(serverConfig) {
    const existing = clients.get(serverConfig.name);
    if (existing) return existing;
    const client = serverConfig.transport === 'http'
      ? new HttpMcpClient(serverConfig, fetchImpl)
      : new StdioMcpClient(serverConfig);
    clients.set(serverConfig.name, client);
    return client;
  }

  async function refreshServer(serverConfig, timeoutMs = 15000) {
    const client = getClient(serverConfig);
    try {
      const tools = await client.listTools(timeoutMs);
      statusMap.set(serverConfig.name, {
        connected: true,
        error: null,
        toolCount: tools.length,
        tools,
      });
      return { ok: true, tools };
    } catch (e) {
      statusMap.set(serverConfig.name, {
        connected: false,
        error: String(e.message || e).slice(0, 200),
        toolCount: 0,
        tools: [],
      });
      return { ok: false, error: String(e.message || e), tools: [] };
    }
  }

  async function getActiveTools(disabledServers = []) {
    const disabledSet = new Set((disabledServers || []).map(normalizeServerName));
    const servers = listConfiguredServers().filter(s => !disabledSet.has(s.name));
    const allTools = [];

    for (const srv of servers) {
      let info = statusMap.get(srv.name);
      if (!info || (!info.connected && !info.error)) {
        await refreshServer(srv, 10000);
        info = statusMap.get(srv.name);
      }
      if (!info?.connected) continue;

      for (const tool of info.tools || []) {
        const cleanTool = String(tool.name || '').replace(/[^a-zA-Z0-9_-]/g, '_');
        if (!cleanTool) continue;
        allTools.push({
          name: `mcp__${srv.name}__${cleanTool}`,
          originalName: tool.name,
          serverName: srv.name,
          description: `[MCP: ${srv.name}] ${tool.description || tool.name}`,
          inputSchema: tool.inputSchema && typeof tool.inputSchema === 'object'
            ? tool.inputSchema
            : { type: 'object', properties: {} },
          readOnly: Boolean(tool.annotations?.readOnlyHint),
        });
      }
    }
    return allTools;
  }

  async function callTool(serverName, originalToolName, args, opts = {}) {
    const servers = listConfiguredServers();
    const srv = servers.find(s => s.name === serverName);
    if (!srv) throw new Error(`MCP server "${serverName}" is not configured.`);
    const client = getClient(srv);
    return client.callTool(originalToolName, args, opts);
  }

  function saveLocalServer(name, spec) {
    const cleanName = normalizeServerName(name);
    if (!cleanName) throw new Error('Invalid MCP server name');
    let current = { mcpServers: {} };
    try {
      const existing = parseJsonc(fs.readFileSync(localMcpFile, 'utf8'));
      if (existing && typeof existing === 'object') {
        current = { ...existing, mcpServers: { ...(existing.mcpServers || {}) } };
      }
    } catch {}

    if (spec === null) {
      delete current.mcpServers[cleanName];
      const c = clients.get(cleanName);
      if (c) { c.close(); clients.delete(cleanName); }
      statusMap.delete(cleanName);
    } else {
      const normalized = normalizeServerEntry(cleanName, spec, 'local');
      if (!normalized) throw new Error('Provide either a valid http(s) URL or a stdio command.');
      current.mcpServers[cleanName] = normalized.transport === 'http'
        ? { url: normalized.url, ...(Object.keys(normalized.headers).length ? { headers: normalized.headers } : {}) }
        : { command: normalized.command, args: normalized.args, ...(Object.keys(normalized.env).length ? { env: normalized.env } : {}) };
      const c = clients.get(cleanName);
      if (c) { c.close(); clients.delete(cleanName); }
      statusMap.delete(cleanName);
    }

    fs.mkdirSync(configDir, { recursive: true, mode: 0o700 });
    fs.writeFileSync(localMcpFile, JSON.stringify(current, null, 2) + '\n', { mode: 0o600 });
  }

  function summary(disabledServers = []) {
    const disabledSet = new Set((disabledServers || []).map(normalizeServerName));
    const servers = listConfiguredServers();
    return {
      localConfigPath: toDisplayPath(localMcpFile, homeDir),
      servers: servers.map(s => {
        const st = statusMap.get(s.name) || { connected: false, error: null, toolCount: 0, tools: [] };
        return {
          name: s.name,
          transport: s.transport,
          target: s.transport === 'http' ? s.url : [s.command, ...(s.args || [])].join(' '),
          source: s.source,
          enabled: !disabledSet.has(s.name),
          connected: st.connected,
          error: st.error,
          toolCount: st.toolCount,
          toolNames: (st.tools || []).map(t => t.name).slice(0, 20),
        };
      }),
    };
  }

  function closeAll() {
    for (const client of clients.values()) {
      try { client.close(); } catch {}
    }
    clients.clear();
    statusMap.clear();
  }

  return {
    listConfiguredServers,
    refreshServer,
    getActiveTools,
    callTool,
    saveLocalServer,
    summary,
    closeAll,
  };
}
