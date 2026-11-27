/* Built-in safe local tools (pure math evaluator + skill resource reader)
   and unified tool dispatcher combining built-in tools with MCP server tools. */

import { readSkillResourceFile, normalizeSkillSlug } from './skills.mjs';
import { findWordRendererSkills, runSkillDocxRender } from './render-bridge.mjs';

/**
 * Zero-eval recursive-descent mathematical expression evaluator.
 * Supports +, -, *, /, ^, %, parentheses, and safe Math helpers:
 * abs, round, floor, ceil, sqrt, pow, min, max, log, ln, exp, sum, avg.
 */
export function evaluateMathExpression(expr) {
  const src = String(expr || '').trim();
  if (!src || src.length > 500) throw new Error('Expression is empty or too long.');

  const tokens = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (/[0-9.]/.test(ch)) {
      let num = '';
      while (i < src.length && /[0-9.eE]/.test(src[i])) {
        if ((src[i] === 'e' || src[i] === 'E') && (src[i + 1] === '+' || src[i + 1] === '-')) {
          num += src[i] + src[i + 1];
          i += 2;
          continue;
        }
        num += src[i++];
      }
      const val = Number(num);
      if (!Number.isFinite(val)) throw new Error(`Invalid number: ${num}`);
      tokens.push({ type: 'num', value: val });
      continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let id = '';
      while (i < src.length && /[A-Za-z0-9_]/.test(src[i])) id += src[i++];
      tokens.push({ type: 'id', value: id.toLowerCase() });
      continue;
    }
    if ('+-*/^%(),'.includes(ch)) {
      tokens.push({ type: 'op', value: ch });
      i++;
      continue;
    }
    throw new Error(`Unsupported character: "${ch}"`);
  }

  let pos = 0;
  const peek = () => tokens[pos];
  const take = () => tokens[pos++];

  const FUNCS = {
    abs: ([x]) => Math.abs(x),
    round: ([x, d = 0]) => {
      const f = Math.pow(10, d);
      return Math.round(x * f) / f;
    },
    floor: ([x]) => Math.floor(x),
    ceil: ([x]) => Math.ceil(x),
    sqrt: ([x]) => Math.sqrt(x),
    pow: ([a, b]) => Math.pow(a, b),
    min: args => Math.min(...args),
    max: args => Math.max(...args),
    log: ([x]) => Math.log10(x),
    ln: ([x]) => Math.log(x),
    exp: ([x]) => Math.exp(x),
    sum: args => args.reduce((a, b) => a + b, 0),
    avg: args => (args.length ? args.reduce((a, b) => a + b, 0) / args.length : 0),
  };

  function parseAddSub() {
    let left = parseMulDiv();
    while (peek()?.type === 'op' && (peek().value === '+' || peek().value === '-')) {
      const op = take().value;
      const right = parseMulDiv();
      left = op === '+' ? left + right : left - right;
    }
    return left;
  }

  function parseMulDiv() {
    let left = parsePow();
    while (peek()?.type === 'op' && (peek().value === '*' || peek().value === '/' || peek().value === '%')) {
      const op = take().value;
      const right = parsePow();
      if (op === '*') left *= right;
      else if (op === '/') {
        if (right === 0) throw new Error('Division by zero');
        left /= right;
      } else {
        left %= right;
      }
    }
    return left;
  }

  function parsePow() {
    let base = parseUnary();
    if (peek()?.type === 'op' && peek().value === '^') {
      take();
      const exp = parsePow();
      base = Math.pow(base, exp);
    }
    return base;
  }

  function parseUnary() {
    if (peek()?.type === 'op' && (peek().value === '+' || peek().value === '-')) {
      const op = take().value;
      const val = parseUnary();
      return op === '-' ? -val : val;
    }
    return parsePrimary();
  }

  function parsePrimary() {
    const t = peek();
    if (!t) throw new Error('Unexpected end of expression');
    if (t.type === 'num') {
      take();
      return t.value;
    }
    if (t.type === 'id') {
      take();
      if (t.value === 'pi') return Math.PI;
      if (t.value === 'e') return Math.E;
      if (peek()?.type === 'op' && peek().value === '(') {
        take();
        const args = [];
        if (!(peek()?.type === 'op' && peek().value === ')')) {
          args.push(parseAddSub());
          while (peek()?.type === 'op' && peek().value === ',') {
            take();
            args.push(parseAddSub());
          }
        }
        if (!(peek()?.type === 'op' && peek().value === ')')) {
          throw new Error('Missing closing parenthesis');
        }
        take();
        const fn = FUNCS[t.value];
        if (!fn) throw new Error(`Unknown function: ${t.value}`);
        return fn(args);
      }
      throw new Error(`Unknown identifier: ${t.value}`);
    }
    if (t.type === 'op' && t.value === '(') {
      take();
      const val = parseAddSub();
      if (!(peek()?.type === 'op' && peek().value === ')')) {
        throw new Error('Missing closing parenthesis');
      }
      take();
      return val;
    }
    throw new Error(`Unexpected token: ${t.value}`);
  }

  const result = parseAddSub();
  if (pos < tokens.length) throw new Error(`Unexpected trailing token: ${tokens[pos].value}`);
  if (!Number.isFinite(result)) throw new Error('Result is not a finite number');
  return result;
}

export function getBuiltinTools({ enabledSkills = [], host = '' } = {}) {
  const tools = [];

  // Expose calculate on Excel (and default unset callers in tests); omit on Word and PowerPoint
  // where some models obsessively loop on trivial arithmetic instead of writing prose/slides.
  if (host !== 'word' && host !== 'powerpoint') {
    tools.push({
      name: 'calculate',
      serverName: 'built-in',
      readOnly: true,
      description: 'Safely evaluate an arithmetic or financial math expression (supports +, -, *, /, ^, %, abs, round, floor, ceil, sqrt, pow, min, max, log, ln, exp, sum, avg) to verify numbers before writing to Excel or PowerPoint.',
      inputSchema: {
        type: 'object',
        properties: {
          expression: { type: 'string', description: 'Mathematical expression, e.g. "(145000 - 120000) / 120000"' },
        },
        required: ['expression'],
      },
    });
  }

  if (enabledSkills.length > 0) {
    tools.push({
      name: 'load_skill',
      serverName: 'built-in',
      readOnly: true,
      description: 'Load the full SKILL.md instructions and list of bundled resource files for an enabled global skill.',
      inputSchema: {
        type: 'object',
        properties: {
          skill: { type: 'string', description: 'Skill name from <available_skills> (e.g. "ic-screener")' },
        },
        required: ['skill'],
      },
    });
  }

  const skillsWithResources = enabledSkills.filter(s => Array.isArray(s.resources) && s.resources.length > 0);
  if (skillsWithResources.length > 0) {
    tools.push({
      name: 'read_skill_resource',
      serverName: 'built-in',
      readOnly: true,
      description: 'Read a bundled reference file from an enabled global skill directory (call load_skill first to see exact resource paths).',
      inputSchema: {
        type: 'object',
        properties: {
          skill: { type: 'string', description: 'Skill name (e.g. "ic-screener")' },
          path: { type: 'string', description: 'Relative resource path inside the skill folder (e.g. "references/style.md" or "SKILL.md")' },
        },
        required: ['skill', 'path'],
      },
    });
  }

  const rendererSkills = host === 'word' ? findWordRendererSkills(enabledSkills) : [];
  if (rendererSkills.length > 0) {
    const names = rendererSkills.map(s => s.name).join(', ');
    tools.push({
      name: 'render_word_template',
      serverName: 'built-in',
      readOnly: true,
      description: `Render a fully branded .docx using an installed Word renderer skill (${names}) from a JSON spec, running any matching linter first and returning a render token you can insert into Word via <edit action="insert_docx" token="..."></edit>.`,
      inputSchema: {
        type: 'object',
        properties: {
          type: { type: 'string', description: 'Document type supported by the renderer, e.g. "screener", "dg", "report", "transcript"' },
          spec_json: { type: 'string', description: 'Complete JSON spec string matching the skill schema (e.g. ScreenerSpec in Generate.ts)' },
          renderer_skill: { type: 'string', description: `Optional renderer skill name (defaults to "${rendererSkills[0].name}")` },
          skip_lint: { type: 'boolean', description: 'Optional: set true only for rough draft renders when placeholder fields are intentionally incomplete' },
        },
        required: ['type', 'spec_json'],
      },
    });
  }

  return tools;
}

export async function executeToolCall({
  toolDef,
  args,
  enabledSkills = [],
  mcpManager = null,
  renderStore = null,
  signal = null,
}) {
  if (!toolDef) {
    return { ok: false, text: 'Unknown tool.' };
  }

  try {
    if (toolDef.name === 'calculate') {
      const value = evaluateMathExpression(args?.expression);
      return { ok: true, text: String(value) };
    }

    if (toolDef.name === 'load_skill') {
      const slug = normalizeSkillSlug(args?.skill);
      const skill = enabledSkills.find(s => s.name === slug);
      if (!skill) return { ok: false, text: `Skill "${args?.skill}" is not enabled or not found.` };
      const resList = Array.isArray(skill.resources) && skill.resources.length
        ? `\n\nBundled resources (exact paths for read_skill_resource): ${skill.resources.join(', ')}`
        : '\n\nBundled resources: (none)';
      return { ok: true, text: `${skill.body || '(empty SKILL.md)'}${resList}` };
    }

    if (toolDef.name === 'read_skill_resource') {
      const slug = normalizeSkillSlug(args?.skill);
      const skill = enabledSkills.find(s => s.name === slug);
      if (!skill) return { ok: false, text: `Skill "${args?.skill}" is not enabled or not found.` };
      try {
        const content = readSkillResourceFile(skill, args?.path);
        return { ok: true, text: content };
      } catch (e) {
        const avail = Array.isArray(skill.resources) && skill.resources.length
          ? ` Available files: SKILL.md, ${skill.resources.join(', ')}`
          : ' Available files: SKILL.md';
        return { ok: false, text: `Resource "${args?.path}" not found in ${skill.name} (${e.message}).${avail}` };
      }
    }

    if (toolDef.name === 'render_word_template') {
      if (!renderStore) return { ok: false, text: 'Render store is not available.' };
      const renderers = findWordRendererSkills(enabledSkills);
      const requested = args?.renderer_skill ? normalizeSkillSlug(args.renderer_skill) : '';
      const rendererSkill = (requested && renderers.find(s => s.name === requested)) || renderers[0];
      return runSkillDocxRender({
        rendererSkill,
        enabledSkills,
        type: args?.type,
        specJson: args?.spec_json,
        skipLint: Boolean(args?.skip_lint),
        renderStore,
      });
    }

    if (toolDef.serverName && toolDef.serverName !== 'built-in' && mcpManager) {
      const res = await mcpManager.callTool(toolDef.serverName, toolDef.originalName, args, { signal });
      return { ok: !res.isError, text: res.text || '(empty result)' };
    }

    return { ok: false, text: `No executor for tool "${toolDef.name}".` };
  } catch (err) {
    return { ok: false, text: `Tool error: ${String(err.message || err).slice(0, 400)}` };
  }
}
