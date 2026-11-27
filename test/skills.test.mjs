import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  parseSkillMarkdown,
  discoverAgentsFiles,
  discoverGlobalSkills,
  resolveActiveSkills,
  buildSkillsPromptSection,
  createSkillsLoader,
} from '../lib/skills.mjs';

function makeTempHome() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-skills-'));
}

test('parseSkillMarkdown parses YAML frontmatter, multiline descriptions, and host filters', () => {
  const raw = [
    '---',
    'name: board-memo',
    'description: >',
    '  Rewrite drafts into crisp executive',
    '  board memo format.',
    'hosts: [word, powerpoint]',
    '---',
    '# Board Memo Rules',
    '1. Lead with the decision.',
    '2. Keep bullets under 15 words.',
  ].join('\n');

  const parsed = parseSkillMarkdown(raw, 'fallback-dir');
  assert.equal(parsed.name, 'board-memo');
  assert.equal(parsed.description, 'Rewrite drafts into crisp executive board memo format.');
  assert.deepEqual(parsed.hosts, ['word', 'powerpoint']);
  assert.match(parsed.body, /Lead with the decision/);
});

test('discoverAgentsFiles loads global AGENTS.md/CLAUDE.md and deduplicates symlinks', () => {
  const home = makeTempHome();
  try {
    fs.mkdirSync(path.join(home, '.agents'), { recursive: true });
    fs.mkdirSync(path.join(home, '.claude'), { recursive: true });
    fs.mkdirSync(path.join(home, '.config', 'word-poc'), { recursive: true });

    const agentsFile = path.join(home, '.agents', 'AGENTS.md');
    fs.writeFileSync(agentsFile, 'Always use British spelling and concise bullet points.\n');
    // Symlink CLAUDE.md -> AGENTS.md; should NOT be double-loaded
    fs.symlinkSync(agentsFile, path.join(home, '.claude', 'CLAUDE.md'));
    fs.writeFileSync(path.join(home, '.config', 'word-poc', 'AGENTS.md'), 'Prefer active voice.\n');

    const files = discoverAgentsFiles({ homeDir: home, env: {} });
    assert.equal(files.length, 2);
    assert.equal(files[0].displayPath, '~/.config/word-poc/AGENTS.md');
    assert.equal(files[1].displayPath, '~/.agents/AGENTS.md');
  } finally {
    fs.rmSync(home, { recursive: true, force: true });
  }
});

test('discoverGlobalSkills scans global directories, respects precedence, and skips _lib/hidden folders', () => {
  const home = makeTempHome();
  try {
    // Higher priority: ~/.agents/skills/exec-tone
    fs.mkdirSync(path.join(home, '.agents', 'skills', 'exec-tone'), { recursive: true });
    fs.writeFileSync(
      path.join(home, '.agents', 'skills', 'exec-tone', 'SKILL.md'),
      '---\nname: exec-tone\ndescription: Primary exec tone\n---\nUse crisp C-suite voice.',
    );

    // Lower priority duplicate in ~/skills/exec-tone + unique ~/skills/dcf-audit + internal ~/skills/_lib
    fs.mkdirSync(path.join(home, 'skills', 'exec-tone'), { recursive: true });
    fs.writeFileSync(
      path.join(home, 'skills', 'exec-tone', 'SKILL.md'),
      '---\nname: exec-tone\ndescription: Shadowed duplicate\n---\nOld instructions.',
    );
    fs.mkdirSync(path.join(home, 'skills', 'dcf-audit'), { recursive: true });
    fs.writeFileSync(
      path.join(home, 'skills', 'dcf-audit', 'SKILL.md'),
      '---\nname: dcf-audit\ndescription: Audit DCF financial models\nhosts: [excel]\n---\nCheck WACC, terminal growth, and cash flow signs.',
    );
    fs.mkdirSync(path.join(home, 'skills', '_lib'), { recursive: true });
    fs.writeFileSync(path.join(home, 'skills', '_lib', 'SKILL.md'), 'should be skipped');

    const skills = discoverGlobalSkills({ homeDir: home, env: {} });
    assert.equal(skills.length, 2);
    const exec = skills.find(s => s.name === 'exec-tone');
    const dcf = skills.find(s => s.name === 'dcf-audit');
    assert.equal(exec.description, 'Primary exec tone');
    assert.equal(exec.source, '~/.agents/skills');
    assert.equal(dcf.source, '~/skills');
    assert.deepEqual(dcf.hosts, ['excel']);
  } finally {
    fs.rmSync(home, { recursive: true, force: true });
  }
});

test('resolveActiveSkills & buildSkillsPromptSection handle /skill-name invocation, pinning, host filtering, and disabling', () => {
  const home = makeTempHome();
  try {
    fs.mkdirSync(path.join(home, '.agents', 'skills', 'dcf-audit'), { recursive: true });
    fs.writeFileSync(
      path.join(home, '.agents', 'skills', 'dcf-audit', 'SKILL.md'),
      '---\nname: dcf-audit\ndescription: Audit DCF formulas\nhosts: [excel]\n---\nVerify WACC and terminal value formulas.',
    );
    fs.mkdirSync(path.join(home, '.agents', 'skills', 'pyramid'), { recursive: true });
    fs.writeFileSync(
      path.join(home, '.agents', 'skills', 'pyramid', 'SKILL.md'),
      '---\nname: pyramid\ndescription: Apply Minto Pyramid Principle\n---\nState the answer first, then supporting pillars.',
    );
    fs.writeFileSync(path.join(home, '.agents', 'AGENTS.md'), 'Be direct and avoid jargon.');

    const loader = createSkillsLoader({ homeDir: home, env: {} });
    const { agentsFiles, skills } = loader.load();

    // In Excel, both skills are compatible; invoking /dcf-audit in user message activates it
    const excelResolved = resolveActiveSkills({
      availableSkills: skills,
      disabledSkills: [],
      explicitSkills: ['pyramid'],
      messages: [{ role: 'user', content: 'Please run /dcf-audit on this tab' }],
      host: 'excel',
    });
    assert.equal(excelResolved.enabledSkills.length, 2);
    assert.deepEqual(excelResolved.activeSkills.map(s => s.name).sort(), ['dcf-audit', 'pyramid']);

    const prompt = buildSkillsPromptSection({
      agentsFiles,
      agentsEnabled: true,
      enabledSkills: excelResolved.enabledSkills,
      activeSkills: excelResolved.activeSkills,
    });
    assert.match(prompt, /<global_instructions>[\s\S]*Be direct and avoid jargon/);
    assert.match(prompt, /<available_skills>[\s\S]*\/dcf-audit/);
    assert.match(prompt, /<active_skill name="dcf-audit"[\s\S]*Verify WACC/);
    assert.match(prompt, /<active_skill name="pyramid"[\s\S]*Minto|State the answer first/);

    // In Word, Excel-only dcf-audit is filtered out; disabledSkills also hides pyramid
    const wordResolved = resolveActiveSkills({
      availableSkills: skills,
      disabledSkills: ['pyramid'],
      explicitSkills: ['dcf-audit', 'pyramid'],
      messages: [{ role: 'user', content: '/dcf-audit' }],
      host: 'word',
    });
    assert.equal(wordResolved.enabledSkills.length, 0);
    assert.equal(wordResolved.activeSkills.length, 0);

    // When agentsEnabled is false, <global_instructions> is omitted
    const agentsOffPrompt = buildSkillsPromptSection({
      agentsFiles,
      agentsEnabled: false,
      enabledSkills: [],
      activeSkills: [],
    });
    assert.equal(agentsOffPrompt, '');
  } finally {
    fs.rmSync(home, { recursive: true, force: true });
  }
});

test('discovers Pi, OpenCode, Gemini, and Cursor instructions/skills, respects AGENTS.override.md, and caches with TTL', () => {
  const home = makeTempHome();
  try {
    // Pi with AGENTS.override.md superseding AGENTS.md in the same directory
    fs.mkdirSync(path.join(home, '.pi', 'agent', 'skills', 'pi-writer', 'references'), { recursive: true });
    fs.writeFileSync(path.join(home, '.pi', 'agent', 'AGENTS.md'), 'Old pi rule.');
    fs.writeFileSync(path.join(home, '.pi', 'agent', 'AGENTS.override.md'), 'Override pi rule wins.');
    fs.writeFileSync(
      path.join(home, '.pi', 'agent', 'skills', 'pi-writer', 'SKILL.md'),
      '---\nname: pi-writer\ndescription: Pi writing skill\n---\nWrite tightly.',
    );
    fs.writeFileSync(path.join(home, '.pi', 'agent', 'skills', 'pi-writer', 'references', 'tone.md'), 'Crisp tone guide.');

    // OpenCode, Gemini, and Cursor
    fs.mkdirSync(path.join(home, '.config', 'opencode', 'skills', 'oc-slides'), { recursive: true });
    fs.writeFileSync(path.join(home, '.config', 'opencode', 'AGENTS.md'), 'OpenCode global rule.');
    fs.writeFileSync(
      path.join(home, '.config', 'opencode', 'skills', 'oc-slides', 'SKILL.md'),
      '---\nname: oc-slides\ndescription: OpenCode slide skill\n---\nKeep slides clean.',
    );
    fs.mkdirSync(path.join(home, '.gemini'), { recursive: true });
    fs.writeFileSync(path.join(home, '.gemini', 'GEMINI.md'), 'Gemini global rule.');
    fs.mkdirSync(path.join(home, '.cursor', 'skills', 'cursor-polish'), { recursive: true });
    fs.writeFileSync(
      path.join(home, '.cursor', 'skills', 'cursor-polish', 'SKILL.md'),
      '---\nname: cursor-polish\ndescription: Cursor polish skill\n---\nPolish grammar.',
    );

    const loader = createSkillsLoader({ homeDir: home, env: {}, cacheTtlMs: 60000 });
    const first = loader.load();

    // AGENTS.override.md replaced AGENTS.md in ~/.pi/agent
    const piPaths = first.agentsFiles.map(f => f.displayPath);
    assert.ok(piPaths.includes('~/.pi/agent/AGENTS.override.md'));
    assert.ok(!piPaths.includes('~/.pi/agent/AGENTS.md'));
    assert.ok(piPaths.includes('~/.config/opencode/AGENTS.md'));
    assert.ok(piPaths.includes('~/.gemini/GEMINI.md'));

    assert.equal(first.skills.length, 3);
    const piSkill = first.skills.find(s => s.name === 'pi-writer');
    assert.deepEqual(piSkill.resources, ['references/tone.md']);

    // Cache hit returns identical object until force: true
    assert.strictEqual(loader.load(), first);
    const refreshed = loader.load({ force: true });
    assert.notStrictEqual(refreshed, first);
  } finally {
    fs.rmSync(home, { recursive: true, force: true });
  }
});
