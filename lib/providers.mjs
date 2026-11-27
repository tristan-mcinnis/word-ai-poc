/* Provider registry, local config & API key management, and SSE streaming adapters
   with native tool calling, upstream AbortSignal propagation, and Anthropic prompt caching. */

import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { parseEnvContent, upsertEnvKey } from './env.mjs';

export const PROVIDERS = {
  deepseek:    { label: 'DeepSeek',    api: 'openai',    base: 'https://api.deepseek.com/v1',                             model: 'deepseek-flash',          envKey: 'DEEPSEEK_API_KEY' },
  anthropic:   { label: 'Claude',      api: 'anthropic', base: 'https://api.anthropic.com/v1',                            model: 'claude-sonnet-4-5',       envKey: 'ANTHROPIC_API_KEY' },
  openai:      { label: 'OpenAI',      api: 'openai',    base: 'https://api.openai.com/v1',                               model: 'gpt-4.1-mini',            envKey: 'OPENAI_API_KEY' },
  gemini:      { label: 'Gemini',      api: 'openai',    base: 'https://generativelanguage.googleapis.com/v1beta/openai', model: 'gemini-2.5-flash',        envKey: 'GEMINI_API_KEY' },
  moonshot:    { label: 'Kimi',        api: 'openai',    base: 'https://api.moonshot.cn/v1',                              model: 'kimi-k2-0905-preview',    envKey: 'KIMI_CN_API_KEY' },
  zhipu:       { label: 'GLM',         api: 'openai',    base: 'https://open.bigmodel.cn/api/paas/v4',                    model: 'glm-4.6',                 envKey: 'ZHIPU_API_KEY' },
  dashscope:   { label: 'Qwen',        api: 'openai',    base: 'https://dashscope.aliyuncs.com/compatible-mode/v1',       model: 'qwen-plus',               envKey: 'DASHSCOPE_API_KEY' },
  siliconflow: { label: 'SiliconFlow', api: 'openai',    base: 'https://api.siliconflow.cn/v1',                           model: 'deepseek-ai/DeepSeek-V3', envKey: 'SILICONFLOW_API_KEY' },
};

function combineSignal(signal, timeoutMs = 180000) {
  const timeoutSig = AbortSignal.timeout(timeoutMs);
  if (!signal) return timeoutSig;
  if (typeof AbortSignal.any === 'function') return AbortSignal.any([signal, timeoutSig]);
  return signal;
}

export function createConfigStore({
  configDir = path.join(os.homedir(), '.config', 'word-poc'),
  cwd = process.cwd(),
  env = process.env,
  enableExternalHelper = true,
} = {}) {
  const configFile = path.join(configDir, 'config.json');
  const envFile = path.join(configDir, '.env');

  function readEnvFile(filePath) {
    try { return parseEnvContent(fs.readFileSync(filePath, 'utf8')); } catch { return {}; }
  }

  function readConfig() {
    let saved = {};
    try { saved = JSON.parse(fs.readFileSync(configFile, 'utf8')); } catch {}
    const provider = PROVIDERS[saved.provider] ? saved.provider : 'deepseek';
    const p = PROVIDERS[provider];
    const rawModel = typeof saved.model === 'string' && saved.model.trim() ? saved.model.trim() : p.model;
    const model = provider === 'deepseek' && rawModel === 'deepseek-chat' ? p.model : rawModel;
    const agentsEnabled = typeof saved.agentsEnabled === 'boolean' ? saved.agentsEnabled : true;
    const toolsEnabled = typeof saved.toolsEnabled === 'boolean' ? saved.toolsEnabled : true;
    const disabledSkills = Array.isArray(saved.disabledSkills)
      ? [...new Set(saved.disabledSkills.map(s => String(s || '').trim().toLowerCase()).filter(Boolean))]
      : [];
    const disabledMcpServers = Array.isArray(saved.disabledMcpServers)
      ? [...new Set(saved.disabledMcpServers.map(s => String(s || '').trim().toLowerCase()).filter(Boolean))]
      : [];
    return {
      provider,
      model,
      base: typeof saved.base === 'string' && saved.base.trim() ? saved.base.trim() : p.base,
      agentsEnabled,
      toolsEnabled,
      disabledSkills,
      disabledMcpServers,
    };
  }

  function writeConfig(next) {
    const merged = { ...readConfig(), ...next };
    fs.mkdirSync(configDir, { recursive: true, mode: 0o700 });
    fs.writeFileSync(configFile, JSON.stringify(merged, null, 2) + '\n', { mode: 0o600 });
  }

  function writeEnvKey(envName, value) {
    fs.mkdirSync(configDir, { recursive: true, mode: 0o700 });
    let current = '';
    try { current = fs.readFileSync(envFile, 'utf8'); } catch {}
    const next = upsertEnvKey(current, envName, value);
    fs.writeFileSync(envFile, next, { mode: 0o600 });
  }

  function keyStatusSummary() {
    const userEnv = readEnvFile(envFile);
    const localEnv = readEnvFile(path.join(cwd, '.env'));
    const hasHelper = enableExternalHelper && fs.existsSync(path.join(os.homedir(), 'skills', '_lib', 'auth.py'));
    return {
      envPath: '~/.config/word-poc/.env',
      hasExternalHelper: hasHelper,
      providers: Object.entries(PROVIDERS).map(([id, p]) => {
        const envName = p.envKey;
        const source = env[envName]
          ? 'env'
          : userEnv[envName]
            ? 'config'
            : localEnv[envName]
              ? 'local'
              : null;
        return {
          id,
          label: p.label,
          envKey: envName,
          configured: Boolean(source),
          source,
        };
      }),
    };
  }

  function secretFor(provider) {
    const p = PROVIDERS[provider];
    if (!p) throw new Error(`Unknown provider: ${provider}`);
    const envName = p.envKey;

    if (env[envName]) return String(env[envName]).trim();

    const userEnv = readEnvFile(envFile);
    if (userEnv[envName]) return userEnv[envName];
    const localEnv = readEnvFile(path.join(cwd, '.env'));
    if (localEnv[envName]) return localEnv[envName];

    const localAuthPy = path.join(os.homedir(), 'skills', '_lib', 'auth.py');
    if (enableExternalHelper && fs.existsSync(localAuthPy)) {
      const vault = env.WORD_POC_OP_VAULT || 'Automation';
      const keyRef = `op://${vault}/${envName}/credential`;
      const helper = [
        'import sys',
        'sys.path.insert(0, ' + JSON.stringify(path.dirname(localAuthPy)) + ')',
        'from auth import resolve',
        'print(resolve(' + JSON.stringify(keyRef) + '), end="")',
      ].join('\n');
      try {
        const resolved = execFileSync('python3', ['-c', helper], {
          encoding: 'utf8',
          stdio: ['ignore', 'pipe', 'pipe'],
          timeout: 45000,
          env: { ...env, PATH: '/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin' },
        }).trim();
        if (resolved) return resolved;
      } catch {}
    }

    throw new Error(`Add ${envName}=your-key to ~/.config/word-poc/.env (or run ./install.sh)`);
  }

  return { readConfig, writeConfig, writeEnvKey, keyStatusSummary, secretFor };
}

export async function* linesOf(body) {
  const decoder = new TextDecoder();
  let buf = '';
  for await (const chunk of body) {
    buf += typeof chunk === 'string' ? chunk : decoder.decode(chunk, { stream: true });
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).replace(/\r$/, '');
      buf = buf.slice(i + 1);
      yield line;
    }
  }
  const flush = decoder.decode();
  if (flush) buf += flush;
  if (buf) yield buf;
}

function safeParseJsonObject(raw) {
  if (!raw || !String(raw).trim()) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : { value: parsed };
  } catch {
    return { raw: String(raw) };
  }
}

/**
 * Stream a single turn from an OpenAI-compatible Chat Completions endpoint.
 * Yields { type: 'text', delta } and { type: 'tool_call', id, name, arguments }.
 */
export async function* streamOpenAITurn(cfg, messages, key, { tools = [], signal = null, fetchImpl = fetch } = {}) {
  const body = { model: cfg.model, messages, stream: true, temperature: 0.4 };
  if (Array.isArray(tools) && tools.length > 0) {
    body.tools = tools.map(t => ({
      type: 'function',
      function: {
        name: t.name,
        description: t.description,
        parameters: t.inputSchema || { type: 'object', properties: {} },
      },
    }));
  }

  const res = await fetchImpl(`${cfg.base}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
    body: JSON.stringify(body),
    signal: combineSignal(signal),
  });
  if (!res.ok) throw new Error(`Provider returned ${res.status}: ${(await res.text()).slice(0, 200)}`);

  const toolAcc = new Map(); // index -> { id, name, argsRaw }
  for await (const line of linesOf(res.body)) {
    if (signal?.aborted) break;
    if (!line.startsWith('data:')) continue;
    const data = line.slice(5).trim();
    if (!data || data === '[DONE]') continue;
    try {
      const parsed = JSON.parse(data);
      const deltaObj = parsed.choices?.[0]?.delta;
      if (!deltaObj) continue;
      if (deltaObj.content) {
        yield { type: 'text', delta: deltaObj.content };
      }
      if (Array.isArray(deltaObj.tool_calls)) {
        for (const tc of deltaObj.tool_calls) {
          const idx = tc.index ?? 0;
          const entry = toolAcc.get(idx) || { id: '', name: '', argsRaw: '' };
          if (tc.id) entry.id = tc.id;
          if (tc.function?.name) entry.name += tc.function.name;
          if (tc.function?.arguments) entry.argsRaw += tc.function.arguments;
          toolAcc.set(idx, entry);
        }
      }
    } catch {}
  }

  const ordered = [...toolAcc.entries()].sort((a, b) => a[0] - b[0]);
  for (const [, tc] of ordered) {
    if (!tc.name) continue;
    yield {
      type: 'tool_call',
      id: tc.id || `call_${Math.random().toString(36).slice(2, 10)}`,
      name: tc.name,
      arguments: safeParseJsonObject(tc.argsRaw),
    };
  }
}

/**
 * Stream a single turn from the Anthropic Messages API with prompt caching & tool use.
 * Yields { type: 'text', delta } and { type: 'tool_call', id, name, arguments }.
 */
export async function* streamAnthropicTurn(cfg, messages, key, { tools = [], signal = null, fetchImpl = fetch } = {}) {
  const systemMsgs = messages.filter(m => m.role === 'system' && m.content);
  const systemBlocks = systemMsgs.map((m, idx) => {
    const block = { type: 'text', text: m.content };
    // Mark stable prefix blocks (host prompt + global AGENTS.md/skills) for ephemeral prompt caching
    if (idx === 0 || idx === systemMsgs.length - 1) {
      block.cache_control = { type: 'ephemeral' };
    }
    return block;
  });
  const rest = messages.filter(m => m.role !== 'system');

  const body = {
    model: cfg.model,
    system: systemBlocks,
    messages: rest,
    max_tokens: 8192,
    stream: true,
  };
  if (Array.isArray(tools) && tools.length > 0) {
    body.tools = tools.map((t, idx) => ({
      name: t.name,
      description: t.description,
      input_schema: t.inputSchema || { type: 'object', properties: {} },
      ...(idx === tools.length - 1 ? { cache_control: { type: 'ephemeral' } } : {}),
    }));
  }

  const res = await fetchImpl(`${cfg.base}/messages`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    body: JSON.stringify(body),
    signal: combineSignal(signal),
  });
  if (!res.ok) throw new Error(`Provider returned ${res.status}: ${(await res.text()).slice(0, 200)}`);

  const toolAcc = new Map(); // content block index -> { id, name, argsRaw }
  for await (const line of linesOf(res.body)) {
    if (signal?.aborted) break;
    if (!line.startsWith('data:')) continue;
    const data = line.slice(5).trim();
    if (!data) continue;
    try {
      const j = JSON.parse(data);
      if (j.type === 'content_block_start' && j.content_block?.type === 'tool_use') {
        toolAcc.set(j.index, {
          id: j.content_block.id || `toolu_${j.index}`,
          name: j.content_block.name || '',
          argsRaw: '',
        });
      } else if (j.type === 'content_block_delta') {
        if (j.delta?.type === 'text_delta' || j.delta?.text) {
          if (j.delta.text) yield { type: 'text', delta: j.delta.text };
        } else if (j.delta?.type === 'input_json_delta' && toolAcc.has(j.index)) {
          toolAcc.get(j.index).argsRaw += j.delta.partial_json || '';
        }
      }
    } catch {}
  }

  const ordered = [...toolAcc.entries()].sort((a, b) => a[0] - b[0]);
  for (const [, tc] of ordered) {
    if (!tc.name) continue;
    yield {
      type: 'tool_call',
      id: tc.id,
      name: tc.name,
      arguments: safeParseJsonObject(tc.argsRaw),
    };
  }
}

/* Backwards-compatible simple text-only stream wrappers */
export async function* streamOpenAI(cfg, messages, key, fetchImpl = fetch, signal = null) {
  for await (const ev of streamOpenAITurn(cfg, messages, key, { fetchImpl, signal })) {
    if (ev.type === 'text') yield ev.delta;
  }
}

export async function* streamAnthropic(cfg, messages, key, fetchImpl = fetch, signal = null) {
  for await (const ev of streamAnthropicTurn(cfg, messages, key, { fetchImpl, signal })) {
    if (ev.type === 'text') yield ev.delta;
  }
}
