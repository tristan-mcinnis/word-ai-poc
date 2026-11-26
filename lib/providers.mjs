/* Provider registry, local config & API key management, and SSE streaming adapters. */

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
    // Migrate deprecated DeepSeek model alias to current default
    const model = provider === 'deepseek' && rawModel === 'deepseek-chat' ? p.model : rawModel;
    return {
      provider,
      model,
      base: typeof saved.base === 'string' && saved.base.trim() ? saved.base.trim() : p.base,
    };
  }

  function writeConfig(next) {
    fs.mkdirSync(configDir, { recursive: true, mode: 0o700 });
    fs.writeFileSync(configFile, JSON.stringify(next, null, 2) + '\n', { mode: 0o600 });
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

    // 1. Process environment variable
    if (env[envName]) return String(env[envName]).trim();

    // 2. Local config .env (~/.config/word-poc/.env) or repo .env
    const userEnv = readEnvFile(envFile);
    if (userEnv[envName]) return userEnv[envName];
    const localEnv = readEnvFile(path.join(cwd, '.env'));
    if (localEnv[envName]) return localEnv[envName];

    // 3. Optional local 1Password helper if present on this machine
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
  if (buf) yield buf;
}

export async function* streamOpenAI(cfg, messages, key, fetchImpl = fetch) {
  const res = await fetchImpl(`${cfg.base}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
    body: JSON.stringify({ model: cfg.model, messages, stream: true, temperature: 0.4 }),
    signal: AbortSignal.timeout(180000),
  });
  if (!res.ok) throw new Error(`Provider returned ${res.status}: ${(await res.text()).slice(0, 200)}`);
  for await (const line of linesOf(res.body)) {
    if (!line.startsWith('data:')) continue;
    const data = line.slice(5).trim();
    if (!data || data === '[DONE]') continue;
    try {
      const delta = JSON.parse(data).choices?.[0]?.delta?.content;
      if (delta) yield delta;
    } catch {}
  }
}

export async function* streamAnthropic(cfg, messages, key, fetchImpl = fetch) {
  const system = messages.filter(m => m.role === 'system').map(m => m.content).join('\n\n');
  const rest = messages.filter(m => m.role !== 'system');
  const res = await fetchImpl(`${cfg.base}/messages`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    body: JSON.stringify({ model: cfg.model, system, messages: rest, max_tokens: 8192, stream: true }),
    signal: AbortSignal.timeout(180000),
  });
  if (!res.ok) throw new Error(`Provider returned ${res.status}: ${(await res.text()).slice(0, 200)}`);
  for await (const line of linesOf(res.body)) {
    if (!line.startsWith('data:')) continue;
    const data = line.slice(5).trim();
    if (!data) continue;
    try {
      const j = JSON.parse(data);
      if (j.type === 'content_block_delta' && j.delta?.text) yield j.delta.text;
    } catch {}
  }
}
