/* Testable HTTP request handler for the local AI Editor add-in server. */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PROVIDERS, createConfigStore, streamOpenAI, streamAnthropic } from './providers.mjs';
import { normalizeHost, systemPromptFor, buildContextPrompt, historyFrom } from './prompts.mjs';
import {
  createSkillsLoader,
  resolveActiveSkills,
  buildSkillsPromptSection,
  isSkillCompatibleWithHost,
  normalizeSkillSlug,
} from './skills.mjs';

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

function contentTypeFor(file) {
  if (file.endsWith('.js')) return 'text/javascript; charset=utf-8';
  if (file.endsWith('.css')) return 'text/css; charset=utf-8';
  if (file.endsWith('.xml')) return 'application/xml; charset=utf-8';
  return 'text/html; charset=utf-8';
}

const sendJson = (res, code, body) =>
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' }).end(JSON.stringify(body));

async function readBody(req, limitBytes) {
  let body = '';
  for await (const chunk of req) {
    body += chunk;
    if (body.length > limitBytes) {
      const err = new Error('Payload too large');
      err.statusCode = 413;
      throw err;
    }
  }
  return body;
}

export function createRequestHandler({
  allowedHost = 'localhost:3443',
  rootDir = DEFAULT_ROOT,
  store = createConfigStore(),
  skillsLoader = createSkillsLoader(),
  fetchImpl = fetch,
  logger = (...args) => console.log(...args),
} = {}) {
  function skillsSummary(hostFilter = '') {
    const cfg = store.readConfig();
    const disabledSet = new Set((cfg.disabledSkills || []).map(normalizeSkillSlug));
    const { agentsFiles, skills } = skillsLoader.load();
    const filteredSkills = hostFilter
      ? skills.filter(s => isSkillCompatibleWithHost(s, hostFilter))
      : skills;

    return {
      agentsEnabled: cfg.agentsEnabled,
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
        source: s.source,
        path: s.path,
        bodyChars: s.bodyChars,
        enabled: !disabledSet.has(s.name),
      })),
      disabledSkills: cfg.disabledSkills || [],
      searchPaths: skillsLoader.getSearchPaths(),
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

    const history = historyFrom(input.messages);
    if (!history.length) return sendJson(res, 400, { error: 'Nothing to send' });

    const cfg = store.readConfig();
    const host = normalizeHost(input.host);
    const messages = [{ role: 'system', content: systemPromptFor(host) }];

    // Inject global AGENTS.md instructions + global skills catalog / active skill bodies
    const { agentsFiles, skills } = skillsLoader.load();
    const { enabledSkills, activeSkills } = resolveActiveSkills({
      availableSkills: skills,
      disabledSkills: cfg.disabledSkills,
      explicitSkills: Array.isArray(input.activeSkills) ? input.activeSkills : [],
      messages: history,
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
    messages.push(...history);

    let key;
    try {
      key = store.secretFor(cfg.provider);
    } catch (e) {
      return sendJson(res, 502, { error: `No API key for ${PROVIDERS[cfg.provider].label}: ${e.message}` });
    }

    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-store',
      Connection: 'keep-alive',
    });
    const send = obj => res.write(`data: ${JSON.stringify(obj)}\n\n`);

    try {
      const stream = PROVIDERS[cfg.provider].api === 'anthropic'
        ? streamAnthropic(cfg, messages, key, fetchImpl)
        : streamOpenAI(cfg, messages, key, fetchImpl);
      for await (const delta of stream) send({ delta });
      send({
        done: true,
        host,
        model: cfg.model,
        provider: cfg.provider,
        providerLabel: PROVIDERS[cfg.provider].label,
        activeSkills: activeSkills.map(s => s.name),
        agentsLoaded: cfg.agentsEnabled && agentsFiles.length > 0,
      });
    } catch (e) {
      send({ error: String(e.message || e).slice(0, 300) });
    }
    res.end();
  }

  return async function handler(req, res) {
    if (logger) logger(new Date().toISOString(), req.method, req.url);
    if (allowedHost && req.headers.host !== allowedHost) {
      res.writeHead(403).end();
      return;
    }

    // Office appends ?_host_Info=... to the SourceLocation, so route on the path only.
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
      if (!PROVIDERS[input.provider]) return sendJson(res, 400, { error: 'Unknown provider' });
      store.writeConfig({
        provider: input.provider,
        model: String(input.model || PROVIDERS[input.provider].model).trim().slice(0, 80),
        base: String(input.base || PROVIDERS[input.provider].base).trim().slice(0, 200),
      });
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
      return sendJson(res, 200, skillsSummary(hostFilter));
    }

    if (pathname === '/api/skills' && req.method === 'POST') {
      let body;
      try { body = await readBody(req, 16384); } catch { return sendJson(res, 413, { error: 'Payload too large' }); }
      let input = {};
      try { input = JSON.parse(body); } catch { return sendJson(res, 400, { error: 'Bad request' }); }

      const current = store.readConfig();
      const patch = {};
      if (typeof input.agentsEnabled === 'boolean') {
        patch.agentsEnabled = input.agentsEnabled;
      }
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
      return sendJson(res, 200, { ok: true, ...skillsSummary(input.host || '') });
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
