/* Generic local .docx template renderer bridge for installed skills that ship
   a tools/Generate.ts renderer (and optional sibling <type>-lint.py + fix_fields.py).
   Rendered .docx files are held briefly in memory behind opaque tokens so the
   Word taskpane can insert them natively via body.insertFileFromBase64(). */

import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';

const MAX_CACHED_RENDERS = 20;
const RENDER_TTL_MS = 30 * 60 * 1000;

function buildToolEnv(baseEnv = process.env) {
  const home = os.homedir();
  const extraPaths = [
    path.join(home, '.bun', 'bin'),
    path.join(home, '.venvs', 'skills', 'bin'),
    '/opt/homebrew/bin',
    '/usr/local/bin',
    '/usr/bin',
    '/bin',
  ];
  const merged = [...extraPaths, ...(baseEnv.PATH ? baseEnv.PATH.split(':') : [])];
  return { ...baseEnv, PATH: [...new Set(merged.filter(Boolean))].join(':') };
}

export function findWordRendererSkills(enabledSkills = []) {
  return (enabledSkills || []).filter(s => {
    if (!s?.dirPath) return false;
    try {
      return fs.existsSync(path.join(s.dirPath, 'tools', 'Generate.ts'));
    } catch {
      return false;
    }
  });
}

function findMatchingLinter(enabledSkills, type) {
  const cleanType = String(type || '').toLowerCase();
  const candidateName = `${cleanType}-lint.py`;
  for (const s of enabledSkills || []) {
    if (!s?.dirPath) continue;
    const p = path.join(s.dirPath, 'scripts', candidateName);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

export function createRenderStore() {
  const cache = new Map(); // token -> { base64, createdAt, type, byteLength }

  function prune() {
    const now = Date.now();
    for (const [k, v] of cache.entries()) {
      if (now - v.createdAt > RENDER_TTL_MS) cache.delete(k);
    }
    while (cache.size > MAX_CACHED_RENDERS) {
      const oldestKey = cache.keys().next().value;
      cache.delete(oldestKey);
    }
  }

  function put({ base64, type, byteLength }) {
    prune();
    const token = `rnd_${crypto.randomBytes(10).toString('hex')}`;
    cache.set(token, { base64, type, byteLength, createdAt: Date.now() });
    return token;
  }

  function get(token) {
    prune();
    const clean = String(token || '').trim();
    if (!/^rnd_[a-f0-9]{20}$/.test(clean)) return null;
    return cache.get(clean) || null;
  }

  return { put, get };
}

export function runSkillDocxRender({
  rendererSkill,
  enabledSkills = [],
  type = 'screener',
  specJson,
  skipLint = false,
  renderStore,
}) {
  if (!rendererSkill?.dirPath) {
    return { ok: false, text: 'No installed skill with tools/Generate.ts was found.' };
  }

  const cleanType = String(type || '').trim().toLowerCase();
  if (!/^[a-z0-9_-]{1,32}$/.test(cleanType)) {
    return { ok: false, text: `Invalid document type "${type}".` };
  }

  let parsedSpec;
  try {
    parsedSpec = typeof specJson === 'string' ? JSON.parse(specJson) : specJson;
    if (!parsedSpec || typeof parsedSpec !== 'object') throw new Error('Spec must be a JSON object');
  } catch (e) {
    return { ok: false, text: `Invalid JSON spec: ${e.message}` };
  }

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'word-ai-render-'));
  const specPath = path.join(tmpDir, `${cleanType}-spec.json`);
  const outPath = path.join(tmpDir, `${cleanType}-output.docx`);
  const env = buildToolEnv();

  try {
    const serialized = JSON.stringify(parsedSpec, null, 2);
    fs.writeFileSync(specPath, serialized, 'utf8');

    // 1. Run sibling linter if available (e.g. screener-lint.py)
    let lintNote = 'no linter found';
    const linterPath = !skipLint ? findMatchingLinter(enabledSkills, cleanType) : null;
    if (linterPath) {
      try {
        const lintOut = execFileSync('python3', [linterPath, specPath], {
          encoding: 'utf8',
          stdio: ['ignore', 'pipe', 'pipe'],
          timeout: 30000,
          env,
        });
        lintNote = lintOut.trim().slice(0, 400) || 'lint passed';
      } catch (e) {
        const stderr = String(e.stderr || '').trim();
        const stdout = String(e.stdout || '').trim();
        const combined = [stdout, stderr].filter(Boolean).join('\n').slice(0, 2000);
        return {
          ok: false,
          text: `Linter (${path.basename(linterPath)}) failed. Fix these issues in spec_json and call render_word_template again:\n${combined}`,
        };
      }
    }

    // 2. Render .docx via bun run tools/Generate.ts --type <type> <outPath>
    const generatorPath = path.join(rendererSkill.dirPath, 'tools', 'Generate.ts');
    try {
      execFileSync('bun', ['run', generatorPath, '--type', cleanType, outPath], {
        input: serialized,
        cwd: rendererSkill.dirPath,
        encoding: 'utf8',
        stdio: ['pipe', 'pipe', 'pipe'],
        timeout: 45000,
        env,
      });
    } catch (e) {
      const detail = [String(e.stdout || '').trim(), String(e.stderr || '').trim(), String(e.message || '')]
        .filter(Boolean)
        .join('\n')
        .slice(0, 1200);
      return {
        ok: false,
        text: `Renderer failed (${rendererSkill.name}/tools/Generate.ts --type ${cleanType}):\n${detail}`,
      };
    }

    if (!fs.existsSync(outPath)) {
      return { ok: false, text: 'Renderer exited without producing a .docx file.' };
    }

    // 3. Patch Word fields if scripts/fix_fields.py is present
    const fixFieldsPath = path.join(rendererSkill.dirPath, 'scripts', 'fix_fields.py');
    if (fs.existsSync(fixFieldsPath)) {
      try {
        execFileSync('python3', [fixFieldsPath, outPath], {
          encoding: 'utf8',
          stdio: ['ignore', 'pipe', 'pipe'],
          timeout: 20000,
          env,
        });
      } catch {
        // Non-fatal; continue with rendered docx
      }
    }

    const buf = fs.readFileSync(outPath);
    const base64 = buf.toString('base64');
    const token = renderStore.put({ base64, type: cleanType, byteLength: buf.length });

    return {
      ok: true,
      text: JSON.stringify({
        rendered: true,
        token,
        type: cleanType,
        bytes: buf.length,
        lint: lintNote,
        next_step: `Emit <edit action="insert_docx" token="${token}" label="Insert rendered ${cleanType} (.docx)"></edit> in your reply so Word inserts the branded document directly.`,
      }),
    };
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }
}
