/* Testable HTTP request handler for the local AI Editor add-in server.
   Supports Word/Excel/PowerPoint chat with streaming tool loops, MCP servers,
   global AGENTS.md & SKILL.md discovery, and persistent local chat sessions. */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  PROVIDERS,
  normalizeTheme,
  createConfigStore,
  streamOpenAITurn,
  streamAnthropicTurn,
} from './providers.mjs';
import { normalizeHost, systemPromptFor, buildContextPrompt } from './prompts.mjs';
import {
  createSkillsLoader,
  resolveActiveSkills,
  buildSkillsPromptSection,
  isSkillCompatibleWithHost,
  normalizeSkillSlug,
} from './skills.mjs';
import { createSessionStore, compactHistory } from './sessions.mjs';
import { createMcpManager, normalizeServerName } from './mcp.mjs';
import { getBuiltinTools, executeToolCall } from './tools.mjs';

export const STATIC_FILES = {
  '/': 'taskpane.html',
  '/taskpane.html': 'taskpane.html',
  '/app.js': 'app.js',
  '/style.css': 'style.css',
  '/manifest.xml': 'manifest.xml',
  '/lib/format.js': 'lib/format.js',
  '/lib/excel-utils.js': 'lib/excel-utils.js',
  '/lib/ppt-utils.js': 'lib/ppt-utils.js',
  '/lib/hosts/index.js': 'lib/hosts/index.js',
  '/lib/hosts/word.js': 'lib/hosts/word.js',
  '/lib/hosts/excel.js': 'lib/hosts/excel.js',
  '/lib/hosts/powerpoint.js': 'lib/hosts/powerpoint.js',
};

const DEFAULT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MAX_TOOL_ROUNDS = 12;

function contentTypeFor(file) {
  if (file.endsWith('.js')) return 'text/javascript; charset=utf-8';
  if (file.endsWith('.css')) return 'text/css; charset=utf-8';
  if (file.endsWith('.xml')) return 'application/xml; charset=utf-8';
  return 'text/html; charset=utf-8';
}

const sendJson = (res, code, body) =>
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' }).end(JSON.stringify(body));

/**
 * UTF-8 safe HTTP body reader using Buffer.concat and true byte limits so
 * multi-byte CJK characters split across TCP packets are never corrupted.
 */
async function readBody(req, limitBytes) {
  const chunks = [];
  let totalBytes = 0;
  for await (const chunk of req) {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk, 'utf8');
    totalBytes += buf.length;
    if (totalBytes > limitBytes) {
      const err = new Error('Payload too large');
      err.statusCode = 413;
      throw err;
    }
    chunks.push(buf);
  }
  return Buffer.concat(chunks).toString('utf8');
}

export function createRequestHandler({
  allowedHost = 'localhost:3443',
  rootDir = DEFAULT_ROOT,
  store = createConfigStore(),
  skillsLoader = createSkillsLoader(),
  sessionStore = createSessionStore(),
  mcpManager = createMcpManager(),
  fetchImpl = fetch,
  logger = (...args) => console.log(...args),
} = {}) {
  function skillsSummary(hostFilter = '', { force = false } = {}) {
    const cfg = store.readConfig();
    const disabledSet = new Set((cfg.disabledSkills || []).map(normalizeSkillSlug));
    const { agentsFiles, skills } = skillsLoader.load({ force });
    const filteredSkills = hostFilter
      ? skills.filter(s => isSkillCompatibleWithHost(s, hostFilter))
      : skills;

    return {
      agentsEnabled: cfg.agentsEnabled,
      toolsEnabled: cfg.toolsEnabled,
      agentsFiles: agentsFiles.map(f => ({
        name: f.name,
        path: f.displayPath,
        chars: f.chars,
      })),
      skills: filteredSkills.map(s => ({
        name: s.name,
        cmd: s.cmd,
        description: s.description,
        hosts: s.hosts,
        resources: s.resources || [],
        source: s.source,
        path: s.path,
        bodyChars: s.bodyChars,
        enabled: !disabledSet.has(s.name),
      })),
      disabledSkills: cfg.disabledSkills || [],
      searchPaths: skillsLoader.getSearchPaths(),
    };
  }

  function mcpSummary() {
    const cfg = store.readConfig();
    const { skills } = skillsLoader.load();
    const disabledSkillSet = new Set((cfg.disabledSkills || []).map(normalizeSkillSlug));
    const enabledSkills = skills.filter(s => !disabledSkillSet.has(s.name));
    const builtins = getBuiltinTools({ enabledSkills }).map(t => ({
      name: t.name,
      description: t.description,
      readOnly: t.readOnly,
    }));
    return {
      toolsEnabled: cfg.toolsEnabled,
      builtinTools: builtins,
      ...mcpManager.summary(cfg.disabledMcpServers),
    };
  }

  async function handleChat(req, res) {
    let body;
    try {
      body = await readBody(req, 600000);
    } catch (e) {
      return sendJson(res, e.statusCode || 413, { error: 'Conversation too long' });
    }

    let input;
    try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }

    const compacted = compactHistory(input.messages, { maxMessages: 24, maxTotalChars: 48000 });
    if (!compacted.messages.length) return sendJson(res, 400, { error: 'Nothing to send' });

    const cfg = store.readConfig();
    const host = normalizeHost(input.host);
    const messages = [{ role: 'system', content: systemPromptFor(host) }];

    // Inject global AGENTS.md instructions + global skills catalog / active skill bodies
    const { agentsFiles, skills } = skillsLoader.load();
    const { enabledSkills, activeSkills } = resolveActiveSkills({
      availableSkills: skills,
      disabledSkills: cfg.disabledSkills,
      explicitSkills: Array.isArray(input.activeSkills) ? input.activeSkills : [],
      messages: compacted.messages,
      host,
    });
    const skillsBlock = buildSkillsPromptSection({
      agentsFiles,
      agentsEnabled: cfg.agentsEnabled,
      enabledSkills,
      activeSkills,
    });
    if (skillsBlock) {
      messages.push({ role: 'system', content: skillsBlock });
    }

    const contextBlock = buildContextPrompt({
      host,
      selection: input.selection,
      documentText: input.documentText,
      contextMode: input.contextMode || 'auto',
      comments: input.comments,
      hostContext: input.hostContext,
    });
    if (contextBlock) {
      messages.push({ role: 'system', content: contextBlock });
    }

    // Separate compacted summary (system role) from user/assistant conversation turns
    for (const m of compacted.messages) {
      messages.push(m);
    }

    let key;
    try {
      key = store.secretFor(cfg.provider);
    } catch (e) {
      return sendJson(res, 502, { error: `No API key for ${PROVIDERS[cfg.provider].label}: ${e.message}` });
    }

    // Prepare available tools (built-in safe tools + enabled MCP tools)
    const toolsEnabled = cfg.toolsEnabled && input.toolsEnabled !== false;
    const allowExternalTools = Boolean(input.allowExternalTools);
    let activeTools = [];
    if (toolsEnabled) {
      const builtins = getBuiltinTools({ enabledSkills, host });
      let mcpTools = [];
      try {
        mcpTools = await mcpManager.getActiveTools(cfg.disabledMcpServers);
      } catch {
        mcpTools = [];
      }
      activeTools = [...builtins, ...mcpTools];
    }
    const toolMap = new Map(activeTools.map(t => [t.name, t]));

    // Abort upstream LLM fetch + MCP tool calls immediately when client disconnects / clicks Stop
    const upstreamAbort = new AbortController();
    const onClose = () => {
      if (!res.writableEnded) upstreamAbort.abort();
    };
    res.on('close', onClose);

    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-store',
      Connection: 'keep-alive',
    });
    const send = obj => {
      if (!res.writableEnded && !upstreamAbort.signal.aborted) {
        res.write(`data: ${JSON.stringify(obj)}\n\n`);
      }
    };

    const isAnthropic = PROVIDERS[cfg.provider].api === 'anthropic';
    const turnMessages = messages.slice();
    const toolsUsed = [];
    const seenToolSignatures = new Set();
    let forceFinalAnswer = false;

    try {
      for (let round = 0; round < MAX_TOOL_ROUNDS; round++) {
        if (upstreamAbort.signal.aborted) break;

        let roundText = '';
        const roundToolCalls = [];
        // On the final allowed round (or after a duplicate-call loop is detected), disable tools
        // so the model always produces a complete final answer instead of looping.
        const roundTools = (round === MAX_TOOL_ROUNDS - 1 || forceFinalAnswer) ? [] : activeTools;

        const stream = isAnthropic
          ? streamAnthropicTurn(cfg, turnMessages, key, { tools: roundTools, signal: upstreamAbort.signal, fetchImpl })
          : streamOpenAITurn(cfg, turnMessages, key, { tools: roundTools, signal: upstreamAbort.signal, fetchImpl });

        for await (const ev of stream) {
          if (upstreamAbort.signal.aborted) break;
          if (ev.type === 'text' && ev.delta) {
            roundText += ev.delta;
            send({ delta: ev.delta });
          } else if (ev.type === 'tool_call') {
            roundToolCalls.push(ev);
          }
        }

        if (upstreamAbort.signal.aborted || roundToolCalls.length === 0) {
          break;
        }

        // Execute requested tool calls and append native provider tool results for next round
        const executed = [];
        let repeatedInRound = false;
        for (const call of roundToolCalls) {
          if (upstreamAbort.signal.aborted) break;
          const sig = `${call.name}:${JSON.stringify(call.arguments || {})}`;
          if (seenToolSignatures.has(sig)) repeatedInRound = true;
          seenToolSignatures.add(sig);
          const toolDef = toolMap.get(call.name);
          const serverLabel = toolDef?.serverName || 'unknown';
          send({ toolStart: { id: call.id, name: call.name, server: serverLabel, args: call.arguments } });

          let outcome;
          if (!toolDef) {
            const isEditHallucination = /^(?:edit|write|insert|replace|update|append|modify)/i.test(call.name || '');
            outcome = {
              ok: false,
              text: isEditHallucination
                ? `There is no "${call.name}" function tool. To write to the document, emit an inline <edit action="append_doc">...</edit> (or replace_selection, insert_after, replace_match, replace_doc) XML block directly in your text response so the Office sidebar applies it.`
                : `Tool "${call.name}" is not available.`,
            };
          } else if (!toolDef.readOnly && serverLabel !== 'built-in' && !allowExternalTools) {
            outcome = {
              ok: false,
              text: `Permission required: external MCP tool "${call.name}" is not read-only. Switch mode to Auto or enable Auto-apply to allow external MCP writes.`,
            };
          } else {
            outcome = await executeToolCall({
              toolDef,
              args: call.arguments,
              enabledSkills,
              mcpManager,
              signal: upstreamAbort.signal,
            });
          }

          toolsUsed.push(call.name);
          const preview = String(outcome.text || '').slice(0, 320);
          send({ toolResult: { id: call.id, name: call.name, server: serverLabel, ok: outcome.ok, preview } });
          executed.push({ call, outcome });
        }
        if (repeatedInRound || toolsUsed.filter(n => n === 'calculate').length >= 4) {
          forceFinalAnswer = true;
        }

        if (isAnthropic) {
          const assistantBlocks = [];
          if (roundText) assistantBlocks.push({ type: 'text', text: roundText });
          for (const { call } of executed) {
            assistantBlocks.push({ type: 'tool_use', id: call.id, name: call.name, input: call.arguments || {} });
          }
          turnMessages.push({ role: 'assistant', content: assistantBlocks });
          turnMessages.push({
            role: 'user',
            content: executed.map(({ call, outcome }) => ({
              type: 'tool_result',
              tool_use_id: call.id,
              content: outcome.text || '',
              is_error: !outcome.ok,
            })),
          });
        } else {
          turnMessages.push({
            role: 'assistant',
            content: roundText || null,
            tool_calls: executed.map(({ call }) => ({
              id: call.id,
              type: 'function',
              function: { name: call.name, arguments: JSON.stringify(call.arguments || {}) },
            })),
          });
          for (const { call, outcome } of executed) {
            turnMessages.push({
              role: 'tool',
              tool_call_id: call.id,
              content: outcome.text || '',
            });
          }
        }
      }

      send({
        done: true,
        host,
        model: cfg.model,
        provider: cfg.provider,
        providerLabel: PROVIDERS[cfg.provider].label,
        activeSkills: activeSkills.map(s => s.name),
        agentsLoaded: cfg.agentsEnabled && agentsFiles.length > 0,
        compacted: compacted.compacted,
        toolsUsed: [...new Set(toolsUsed)],
      });
    } catch (e) {
      if (!upstreamAbort.signal.aborted && e?.name !== 'AbortError') {
        send({ error: String(e.message || e).slice(0, 300) });
      }
    } finally {
      res.off('close', onClose);
      if (!res.writableEnded) res.end();
    }
  }

  return async function handler(req, res) {
    if (logger) logger(new Date().toISOString(), req.method, req.url);
    if (allowedHost && req.headers.host !== allowedHost) {
      res.writeHead(403).end();
      return;
    }

    let parsedUrl;
    try {
      parsedUrl = new URL(req.url, `https://${allowedHost || 'localhost:3443'}`);
    } catch {
      parsedUrl = { pathname: req.url, searchParams: new URLSearchParams() };
    }
    const pathname = parsedUrl.pathname;

    if (pathname === '/api/config' && req.method === 'GET') {
      return sendJson(res, 200, {
        active: store.readConfig(),
        providers: Object.entries(PROVIDERS).map(([id, p]) => ({ id, label: p.label, model: p.model })),
      });
    }

    if (pathname === '/api/config' && req.method === 'POST') {
      let body;
      try { body = await readBody(req, 16384); } catch { return sendJson(res, 413, { error: 'Payload too large' }); }
      let input = {};
      try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }
      const patch = {};
      if (input.provider !== undefined) {
        if (!PROVIDERS[input.provider]) return sendJson(res, 400, { error: 'Unknown provider' });
        patch.provider = input.provider;
        patch.model = String(input.model || PROVIDERS[input.provider].model).trim().slice(0, 80);
        patch.base = String(input.base || PROVIDERS[input.provider].base).trim().slice(0, 200);
      }
      if (input.theme && typeof input.theme === 'object') {
        patch.theme = normalizeTheme(input.theme);
      }
      if (Object.keys(patch).length === 0) return sendJson(res, 400, { error: 'Nothing to update' });
      store.writeConfig(patch);
      return sendJson(res, 200, { active: store.readConfig() });
    }

    if (pathname === '/api/keys' && req.method === 'GET') {
      return sendJson(res, 200, store.keyStatusSummary());
    }

    if (pathname === '/api/keys' && req.method === 'POST') {
      let body;
      try { body = await readBody(req, 16384); } catch { return sendJson(res, 413, { error: 'Payload too large' }); }
      let input = {};
      try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }
      const p = PROVIDERS[input.provider];
      if (!p) return sendJson(res, 400, { error: 'Unknown provider' });
      const apiKey = typeof input.apiKey === 'string' ? input.apiKey.trim() : '';
      store.writeEnvKey(p.envKey, apiKey);
      return sendJson(res, 200, { ok: true, ...store.keyStatusSummary() });
    }

    if (pathname === '/api/skills' && req.method === 'GET') {
      const hostFilter = parsedUrl.searchParams.get('host') || '';
      const force = parsedUrl.searchParams.get('refresh') === '1';
      return sendJson(res, 200, skillsSummary(hostFilter, { force }));
    }

    if (pathname === '/api/skills' && req.method === 'POST') {
      let body;
      try { body = await readBody(req, 16384); } catch { return sendJson(res, 413, { error: 'Payload too large' }); }
      let input = {};
      try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }

      const current = store.readConfig();
      const patch = {};
      if (typeof input.agentsEnabled === 'boolean') patch.agentsEnabled = input.agentsEnabled;
      if (typeof input.toolsEnabled === 'boolean') patch.toolsEnabled = input.toolsEnabled;
      if (typeof input.toggleSkill === 'string' && typeof input.enabled === 'boolean') {
        const slug = normalizeSkillSlug(input.toggleSkill);
        const set = new Set((current.disabledSkills || []).map(normalizeSkillSlug));
        if (input.enabled) set.delete(slug);
        else if (slug) set.add(slug);
        patch.disabledSkills = [...set];
      } else if (Array.isArray(input.disabledSkills)) {
        patch.disabledSkills = [...new Set(input.disabledSkills.map(normalizeSkillSlug).filter(Boolean))];
      }
      store.writeConfig(patch);
      return sendJson(res, 200, { ok: true, ...skillsSummary(input.host || '', { force: true }) });
    }

    if (pathname === '/api/mcp' && req.method === 'GET') {
      return sendJson(res, 200, mcpSummary());
    }

    if (pathname === '/api/mcp' && req.method === 'POST') {
      let body;
      try { body = await readBody(req, 32768); } catch { return sendJson(res, 413, { error: 'Payload too large' }); }
      let input = {};
      try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }

      const current = store.readConfig();
      const patch = {};

      if (typeof input.toolsEnabled === 'boolean') {
        patch.toolsEnabled = input.toolsEnabled;
      }

      if (typeof input.toggleServer === 'string' && typeof input.enabled === 'boolean') {
        const name = normalizeServerName(input.toggleServer);
        const set = new Set((current.disabledMcpServers || []).map(normalizeServerName));
        if (input.enabled) set.delete(name);
        else if (name) set.add(name);
        patch.disabledMcpServers = [...set];
      }

      if (Object.keys(patch).length > 0) {
        store.writeConfig(patch);
      }

      if (input.addServer && typeof input.addServer === 'object') {
        try {
          mcpManager.saveLocalServer(input.addServer.name, input.addServer);
        } catch (err) {
          return sendJson(res, 400, { error: err.message || 'Invalid MCP server config' });
        }
      }

      if (typeof input.removeServer === 'string') {
        try {
          mcpManager.saveLocalServer(input.removeServer, null);
        } catch (err) {
          return sendJson(res, 400, { error: err.message || 'Invalid MCP server name' });
        }
      }

      if (typeof input.testServer === 'string') {
        const target = mcpManager.listConfiguredServers().find(s => s.name === normalizeServerName(input.testServer));
        if (!target) return sendJson(res, 404, { error: 'MCP server not found' });
        const probe = await mcpManager.refreshServer(target, 12000);
        return sendJson(res, 200, { ok: probe.ok, error: probe.error || null, ...mcpSummary() });
      }

      return sendJson(res, 200, { ok: true, ...mcpSummary() });
    }

    if (pathname === '/api/sessions' && req.method === 'GET') {
      const host = parsedUrl.searchParams.get('host') || '';
      return sendJson(res, 200, { sessions: sessionStore.list({ host }) });
    }

    if (pathname === '/api/sessions' && req.method === 'POST') {
      let body;
      try { body = await readBody(req, 600000); } catch { return sendJson(res, 413, { error: 'Session too large' }); }
      let input = {};
      try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }
      const saved = sessionStore.save({
        id: input.id,
        host: normalizeHost(input.host),
        title: input.title,
        messages: input.messages,
      });
      return sendJson(res, 200, { ok: true, session: saved });
    }

    const sessionMatch = /^\/api\/sessions\/([a-zA-Z0-9_-]{4,64})$/.exec(pathname);
    if (sessionMatch && req.method === 'GET') {
      const found = sessionStore.get(sessionMatch[1]);
      if (!found) return sendJson(res, 404, { error: 'Session not found' });
      return sendJson(res, 200, { session: found });
    }
    if (sessionMatch && req.method === 'DELETE') {
      const removed = sessionStore.remove(sessionMatch[1]);
      return sendJson(res, removed ? 200 : 404, { ok: removed });
    }

    if (pathname === '/api/chat' && req.method === 'POST') {
      return handleChat(req, res);
    }

    const relFile = STATIC_FILES[pathname];
    if (!relFile || req.method !== 'GET') {
      res.writeHead(404).end();
      return;
    }

    const absPath = path.join(rootDir, relFile);
    try {
      const data = fs.readFileSync(absPath);
      res.writeHead(200, {
        'Content-Type': contentTypeFor(relFile),
        'Cache-Control': 'no-store',
      });
      res.end(data);
    } catch {
      res.writeHead(404).end();
    }
  };
}
