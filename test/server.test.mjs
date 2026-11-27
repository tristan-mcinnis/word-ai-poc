import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createConfigStore } from '../lib/providers.mjs';
import { createRequestHandler, STATIC_FILES } from '../lib/server-core.mjs';

function makeTempConfigDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-test-'));
}

async function startTestServer({ store, fetchImpl } = {}) {
  const handler = createRequestHandler({
    allowedHost: 'localhost:3443',
    store,
    fetchImpl,
    logger: null,
  });
  const server = http.createServer(handler);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();

  const request = (urlPath, { method = 'GET', headers = {}, body = null } = {}) =>
    new Promise((resolve, reject) => {
      const req = http.request(
        {
          hostname: '127.0.0.1',
          port,
          path: urlPath,
          method,
          headers: { Host: 'localhost:3443', ...headers },
        },
        res => {
          let data = '';
          res.on('data', chunk => { data += chunk; });
          res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: data }));
        },
      );
      req.on('error', reject);
      if (body) req.write(body);
      req.end();
    });

  return { server, request, close: () => new Promise(r => server.close(r)) };
}

test('server enforces Host allowlist, serves static allowlist, and blocks traversal/private files', async () => {
  const tmpDir = makeTempConfigDir();
  const store = createConfigStore({ configDir: tmpDir, cwd: tmpDir, env: {}, enableExternalHelper: false });
  const srv = await startTestServer({ store });
  try {
    // DNS rebinding Host rejection
    const badHost = await srv.request('/api/config', { headers: { Host: 'evil.com:3443' } });
    assert.equal(badHost.status, 403);

    // Every allowlisted static file returns 200
    for (const urlPath of Object.keys(STATIC_FILES)) {
      const res = await srv.request(urlPath);
      assert.equal(res.status, 200, `Expected 200 for ${urlPath}`);
    }

    // Server internals and path traversal are blocked
    assert.equal((await srv.request('/server.mjs')).status, 404);
    assert.equal((await srv.request('/lib/providers.mjs')).status, 404);
    assert.equal((await srv.request('/../package.json')).status, 404);
  } finally {
    await srv.close();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('server /api/config and /api/keys persist settings without ever echoing secrets back', async () => {
  const tmpDir = makeTempConfigDir();
  const store = createConfigStore({ configDir: tmpDir, cwd: tmpDir, env: {}, enableExternalHelper: false });
  const srv = await startTestServer({ store });
  try {
    const cfgPost = await srv.request('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: 'anthropic' }),
    });
    assert.equal(cfgPost.status, 200);
    assert.equal(JSON.parse(cfgPost.body).active.provider, 'anthropic');

    const secret = 'sk-ant-super-secret-value-999';
    const keySave = await srv.request('/api/keys', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: 'anthropic', apiKey: secret }),
    });
    assert.equal(keySave.status, 200);
    assert.doesNotMatch(keySave.body, new RegExp(secret));

    const keyGet = await srv.request('/api/keys');
    assert.equal(keyGet.status, 200);
    assert.doesNotMatch(keyGet.body, new RegExp(secret));
    const parsed = JSON.parse(keyGet.body);
    const claude = parsed.providers.find(p => p.id === 'anthropic');
    assert.equal(claude.configured, true);
    assert.equal(claude.source, 'config');
  } finally {
    await srv.close();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('server /api/chat routes host-specific prompts and streams OpenAI and Anthropic SSE cleanly', async () => {
  const tmpDir = makeTempConfigDir();
  const store = createConfigStore({
    configDir: tmpDir,
    cwd: tmpDir,
    env: { DEEPSEEK_API_KEY: 'ds-key', ANTHROPIC_API_KEY: 'ant-key' },
    enableExternalHelper: false,
  });

  const capturedBodies = [];
  const mockFetch = async (url, opts) => {
    const payload = JSON.parse(opts.body);
    capturedBodies.push({ url, payload });

    const isAnthropic = url.includes('anthropic.com');
    const chunks = isAnthropic
      ? [
          'data: {"type":"content_block_delta","delta":{"text":"Slide 2 looks crisp. "}}\n',
          'data: {"type":"content_block_delta","delta":{"text":"<edit action=\\"add_slide\\" layout=\\"Title and Content\\">Title: Next Steps</edit>"}}\n',
        ]
      : [
          'data: {"choices":[{"delta":{"content":"Check `C4`: "}}]}\n',
          'data: {"choices":[{"delta":{"content":"<edit action=\\"set_formula\\" cell=\\"C4\\">=IFERROR(C2/C3,0)</edit>"}}]}\n',
          'data: [DONE]\n',
        ];

    const body = (async function* () {
      for (const c of chunks) yield new TextEncoder().encode(c);
    })();
    return { ok: true, status: 200, body, text: async () => '' };
  };

  const srv = await startTestServer({ store, fetchImpl: mockFetch });
  try {
    // 1. Excel chat via OpenAI-compatible provider (DeepSeek default)
    const excelRes = await srv.request('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        host: 'excel',
        messages: [{ role: 'user', content: 'Fix C4' }],
        documentText: '| Row | A | B | C |\n| 4 | Margin | 0.4 | #DIV/0! |',
        hostContext: {
          sheets: [{ name: 'Model', active: true }],
          errors: ['Model!C4: #DIV/0!'],
        },
      }),
    });
    assert.equal(excelRes.status, 200);
    assert.match(excelRes.body, /set_formula/);
    assert.match(excelRes.body, /"host":"excel"/);
    assert.match(capturedBodies[0].payload.messages[0].content, /Microsoft Excel/);
    assert.match(capturedBodies[0].payload.messages[1].content, /<errors>[\s\S]*#DIV\/0!/);

    // 2. Switch to Claude and run PowerPoint chat
    await srv.request('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: 'anthropic' }),
    });
    const pptRes = await srv.request('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        host: 'powerpoint',
        messages: [{ role: 'user', content: 'Add a next steps slide' }],
        documentText: '--- Slide 1 ---\n[Title 1] Board Deck',
        hostContext: { layouts: [{ layoutName: 'Title and Content' }] },
      }),
    });
    assert.equal(pptRes.status, 200);
    assert.match(pptRes.body, /add_slide/);
    assert.match(pptRes.body, /"host":"powerpoint"/);
    const anthropicSysText = Array.isArray(capturedBodies[1].payload.system)
      ? capturedBodies[1].payload.system.map(b => b.text).join('\n')
      : String(capturedBodies[1].payload.system || '');
    assert.match(anthropicSysText, /Microsoft PowerPoint/);
    assert.match(anthropicSysText, /<layouts>/);
  } finally {
    await srv.close();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('server /api/skills discovers global AGENTS.md and skills, persists toggles, and injects them into /api/chat', async () => {
  const { createSkillsLoader } = await import('../lib/skills.mjs');
  const tmpDir = makeTempConfigDir();
  const fakeHome = makeTempConfigDir();
  fs.mkdirSync(path.join(fakeHome, '.agents', 'skills', 'exec-brief'), { recursive: true });
  fs.writeFileSync(path.join(fakeHome, '.agents', 'AGENTS.md'), 'Always lead with the bottom line.');
  fs.writeFileSync(
    path.join(fakeHome, '.agents', 'skills', 'exec-brief', 'SKILL.md'),
    '---\nname: exec-brief\ndescription: Executive brief format\n---\nFormat into BLUF + 3 bullets.',
  );

  const store = createConfigStore({
    configDir: tmpDir,
    cwd: tmpDir,
    env: { DEEPSEEK_API_KEY: 'ds-key' },
    enableExternalHelper: false,
  });
  const skillsLoader = createSkillsLoader({ homeDir: fakeHome, configDir: tmpDir, env: {} });

  const captured = [];
  const mockFetch = async (_url, opts) => {
    captured.push(JSON.parse(opts.body));
    const body = (async function* () {
      yield new TextEncoder().encode('data: {"choices":[{"delta":{"content":"BLUF: Ready."}}]}\n');
      yield new TextEncoder().encode('data: [DONE]\n');
    })();
    return { ok: true, status: 200, body, text: async () => '' };
  };

  const handler = createRequestHandler({
    allowedHost: 'localhost:3443',
    store,
    skillsLoader,
    fetchImpl: mockFetch,
    logger: null,
  });
  const server = http.createServer(handler);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();

  const request = (urlPath, { method = 'GET', headers = {}, body = null } = {}) =>
    new Promise((resolve, reject) => {
      const req = http.request(
        { hostname: '127.0.0.1', port, path: urlPath, method, headers: { Host: 'localhost:3443', ...headers } },
        res => {
          let data = '';
          res.on('data', chunk => { data += chunk; });
          res.on('end', () => resolve({ status: res.statusCode, body: data }));
        },
      );
      req.on('error', reject);
      if (body) req.write(body);
      req.end();
    });

  try {
    const listRes = await request('/api/skills?host=word');
    assert.equal(listRes.status, 200);
    const listData = JSON.parse(listRes.body);
    assert.equal(listData.agentsFiles.length, 1);
    assert.equal(listData.skills.length, 1);
    assert.equal(listData.skills[0].name, 'exec-brief');
    assert.equal(listData.skills[0].enabled, true);

    // Chat invoking /exec-brief injects both <global_instructions> and <active_skill>
    const chatRes = await request('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        host: 'word',
        messages: [{ role: 'user', content: '/exec-brief summarize this memo' }],
      }),
    });
    assert.equal(chatRes.status, 200);
    assert.match(chatRes.body, /"activeSkills":\["exec-brief"\]/);
    assert.match(chatRes.body, /"agentsLoaded":true/);
    const sysCombined = captured[0].messages.filter(m => m.role === 'system').map(m => m.content).join('\n');
    assert.match(sysCombined, /<global_instructions>[\s\S]*Always lead with the bottom line/);
    assert.match(sysCombined, /<active_skill name="exec-brief"[\s\S]*BLUF \+ 3 bullets/);

    // Disable AGENTS.md and exec-brief via POST /api/skills; verify persistence across provider switch
    const toggleRes = await request('/api/skills', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ agentsEnabled: false, toggleSkill: 'exec-brief', enabled: false }),
    });
    assert.equal(toggleRes.status, 200);
    const toggled = JSON.parse(toggleRes.body);
    assert.equal(toggled.agentsEnabled, false);
    assert.equal(toggled.skills[0].enabled, false);

    // Switching provider should NOT wipe agentsEnabled or disabledSkills
    await request('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: 'openai' }),
    });
    const afterSwitch = JSON.parse((await request('/api/skills')).body);
    assert.equal(afterSwitch.agentsEnabled, false);
    assert.equal(afterSwitch.skills[0].enabled, false);
  } finally {
    await new Promise(r => server.close(r));
    fs.rmSync(tmpDir, { recursive: true, force: true });
    fs.rmSync(fakeHome, { recursive: true, force: true });
  }
});

test('server executes multi-step OpenAI & Anthropic tool loops, prompt caching, sessions CRUD, UTF-8 chunking, and upstream abort', async () => {
  const { createSessionStore } = await import('../lib/sessions.mjs');
  const { createMcpManager } = await import('../lib/mcp.mjs');
  const tmpDir = makeTempConfigDir();
  const store = createConfigStore({
    configDir: tmpDir,
    cwd: tmpDir,
    env: { DEEPSEEK_API_KEY: 'ds-key', ANTHROPIC_API_KEY: 'ant-key' },
    enableExternalHelper: false,
  });
  const sessionStore = createSessionStore({ configDir: tmpDir });
  const mcpManager = createMcpManager({ homeDir: tmpDir, configDir: tmpDir, configPaths: [] });

  let callNum = 0;
  let capturedAnthropicBody = null;
  let abortObserved = false;

  const mockFetch = async (url, opts) => {
    const payload = JSON.parse(opts.body);
    const isAnthropic = url.includes('anthropic.com');

    // Special slow stream case to verify upstream AbortSignal propagation
    if (payload.messages?.some(m => typeof m.content === 'string' && m.content.includes('__SLOW_ABORT__'))) {
      const body = (async function* () {
        yield new TextEncoder().encode('data: {"choices":[{"delta":{"content":"Starting..."}}]}\n');
        await new Promise((resolve, reject) => {
          const t = setTimeout(resolve, 5000);
          opts.signal?.addEventListener('abort', () => {
            clearTimeout(t);
            abortObserved = true;
            reject(new Error('Aborted'));
          });
        });
      })();
      return { ok: true, status: 200, body, text: async () => '' };
    }

    if (isAnthropic) {
      capturedAnthropicBody = payload;
      const hasToolResult = payload.messages.some(m => Array.isArray(m.content) && m.content.some(b => b.type === 'tool_result'));
      const chunks = !hasToolResult
        ? [
            'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"toolu_1","name":"calculate"}}\n',
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"expression\\":\\"25*4\\"}"}}\n',
          ]
        : [
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"The verified result is 100."}}\n',
          ];
      const body = (async function* () {
        for (const c of chunks) yield new TextEncoder().encode(c);
      })();
      return { ok: true, status: 200, body, text: async () => '' };
    }

    callNum++;
    const hasToolMsg = payload.messages.some(m => m.role === 'tool');
    const chunks = !hasToolMsg
      ? [
          'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"calculate","arguments":"{\\"expression\\":\\"120*1.15\\"}"}}]}}]}\n',
          'data: [DONE]\n',
        ]
      : [
          `data: {"choices":[{"delta":{"content":"Computed value: ${payload.messages.find(m => m.role === 'tool').content}"}}]}\n`,
          'data: [DONE]\n',
        ];
    const body = (async function* () {
      for (const c of chunks) yield new TextEncoder().encode(c);
    })();
    return { ok: true, status: 200, body, text: async () => '' };
  };

  const handler = createRequestHandler({
    allowedHost: 'localhost:3443',
    store,
    sessionStore,
    mcpManager,
    fetchImpl: mockFetch,
    logger: null,
  });
  const server = http.createServer(handler);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();

  const request = (urlPath, { method = 'GET', headers = {}, body = null } = {}) =>
    new Promise((resolve, reject) => {
      const req = http.request(
        { hostname: '127.0.0.1', port, path: urlPath, method, headers: { Host: 'localhost:3443', ...headers } },
        res => {
          let data = '';
          res.on('data', chunk => { data += chunk; });
          res.on('end', () => resolve({ status: res.statusCode, body: data }));
        },
      );
      req.on('error', reject);
      if (body) req.write(body);
      req.end();
    });

  try {
    // 1. OpenAI multi-step tool loop with calculate tool
    const oaiRes = await request('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ host: 'excel', messages: [{ role: 'user', content: 'What is 120 * 1.15?' }] }),
    });
    assert.equal(oaiRes.status, 200);
    assert.match(oaiRes.body, /"toolStart"/);
    assert.match(oaiRes.body, /"toolResult"/);
    assert.match(oaiRes.body, /138/);
    assert.equal(callNum, 2);

    // 2. Anthropic multi-step tool loop + ephemeral prompt caching breakpoints
    await request('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: 'anthropic' }),
    });
    const antRes = await request('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ host: 'excel', messages: [{ role: 'user', content: 'Compute 25 * 4' }] }),
    });
    assert.equal(antRes.status, 200);
    assert.match(antRes.body, /The verified result is 100/);
    assert.ok(Array.isArray(capturedAnthropicBody.system));
    assert.equal(capturedAnthropicBody.system[0].cache_control?.type, 'ephemeral');

    // 3. /api/sessions CRUD + split multi-byte UTF-8 CJK body
    const cjkPayload = Buffer.from(JSON.stringify({
      host: 'word',
      messages: [{ role: 'user', content: '请润色这段季度财报摘要' }],
    }), 'utf8');
    const sessSave = await new Promise((resolve, reject) => {
      const req = http.request(
        { hostname: '127.0.0.1', port, path: '/api/sessions', method: 'POST', headers: { Host: 'localhost:3443', 'Content-Type': 'application/json' } },
        res => {
          let data = '';
          res.on('data', c => { data += c; });
          res.on('end', () => resolve({ status: res.statusCode, body: data }));
        },
      );
      req.on('error', reject);
      // Deliberately split mid-UTF8 CJK byte boundary!
      const mid = Math.floor(cjkPayload.length / 2);
      req.write(cjkPayload.subarray(0, mid));
      req.write(cjkPayload.subarray(mid));
      req.end();
    });
    assert.equal(sessSave.status, 200);
    const savedId = JSON.parse(sessSave.body).session.id;
    const loadedSess = JSON.parse((await request(`/api/sessions/${savedId}`)).body).session;
    assert.equal(loadedSess.messages[0].content, '请润色这段季度财报摘要');
    assert.equal((await request(`/api/sessions/${savedId}`, { method: 'DELETE' })).status, 200);

    // 4. Upstream abort propagation when client destroys socket mid-stream
    await request('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: 'deepseek' }),
    });
    await new Promise(resolve => {
      const req = http.request(
        { hostname: '127.0.0.1', port, path: '/api/chat', method: 'POST', headers: { Host: 'localhost:3443', 'Content-Type': 'application/json' } },
        res => {
          res.once('data', () => {
            req.destroy();
            setTimeout(resolve, 60);
          });
        },
      );
      req.on('error', () => {});
      req.write(JSON.stringify({ host: 'word', messages: [{ role: 'user', content: '__SLOW_ABORT__' }] }));
      req.end();
    });
    assert.equal(abortObserved, true, 'Upstream fetch AbortSignal should fire when client disconnects');
  } finally {
    mcpManager.closeAll();
    await new Promise(r => server.close(r));
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});
