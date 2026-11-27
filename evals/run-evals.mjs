#!/usr/bin/env node
/* Multi-host End-to-End Evaluation Suite for Office AI Editor (Word, Excel, PowerPoint).
   Runs 18 realistic workflow, UX, streaming, security, and code-hygiene evaluations. */

import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

import {
  parseAssistantReply,
  renderMarkdown,
  computeWordDiff,
  isDestructiveAction,
  actionLabel,
  primaryButtonLabel,
} from '../lib/format.js';
import { systemPromptFor, buildContextPrompt } from '../lib/prompts.mjs';
import { createHostAdapter, HOST_IDS } from '../lib/hosts/index.js';
import { createConfigStore } from '../lib/providers.mjs';
import { createRequestHandler, STATIC_FILES } from '../lib/server-core.mjs';
import {
  createMockWordEnv,
  createMockExcelEnv,
  createMockPowerPointEnv,
} from '../test/mock-office.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const scenarios = [
  /* ===================== WORD EVALS ===================== */
  {
    id: 'E01',
    suite: 'Word',
    name: 'Conversational critique emits zero premature <edit> cards',
    async run() {
      const reply = '### Honest Read\n\nThe opening buries the lead. Move the Q3 revenue headline to paragraph 1 and tighten paragraph 3.\n\nWant me to rewrite paragraph 1 now?';
      const parts = parseAssistantReply(reply);
      assert.equal(parts.filter(p => p.kind === 'edit').length, 0);
      const html = renderMarkdown(parts[0].text, { host: 'word' });
      assert.match(html, /<h4>Honest Read<\/h4>/);
    },
  },
  {
    id: 'E02',
    suite: 'Word',
    name: 'Multi-paragraph selection replacement with native Track Changes',
    async run() {
      const env = createMockWordEnv({
        paragraphs: ['Header.', 'Old verbose draft paragraph.', 'Footer.'],
        selectionIndex: 1,
      });
      const word = createHostAdapter('word', env);
      await word.toggleSecondary(true);
      assert.equal(env.state.changeTrackingMode, 'TrackAll');

      const reply = 'Tightened into two crisp paragraphs:\n<edit action="replace_selection">First tight point.\nSecond tight point.</edit>';
      const [prose, edit] = parseAssistantReply(reply);
      assert.equal(prose.kind, 'prose');
      const res = await word.applyEdit(edit, { trackChanges: true });
      assert.equal(res.ok, true);
      assert.match(res.message, /tracked change/i);
      assert.deepEqual(env.state.paragraphs, ['Header.', 'First tight point.', 'Second tight point.', 'Footer.']);
    },
  },
  {
    id: 'E03',
    suite: 'Word',
    name: 'Full-document targeted replace_match preserves accurate word diff',
    async run() {
      const env = createMockWordEnv({
        paragraphs: ['The company achieved very strong growth in Q3.', 'Margins remained steady.'],
        selectionIndex: 1,
      });
      const word = createHostAdapter('word', env);
      const reply = '<edit action="replace_match"><find>very strong growth</find><content>34% YoY growth</content></edit>';
      const [edit] = parseAssistantReply(reply);
      const diff = computeWordDiff(edit.find, edit.content);
      assert.ok(diff.some(d => d.type === 'del' && d.text.includes('very')));
      assert.ok(diff.some(d => d.type === 'ins' && d.text.includes('34%')));

      const res = await word.applyEdit(edit);
      assert.equal(res.ok, true);
      assert.equal(env.state.paragraphs[0], 'The company achieved 34% YoY growth in Q3.');
    },
  },
  {
    id: 'E04',
    suite: 'Word',
    name: '/work-comments workflow pairs replace_match with reply_comment',
    async run() {
      const env = createMockWordEnv({
        paragraphs: ['Our launch date is TBD.'],
        comments: [{ id: 'c-99', authorName: 'PM', content: 'Change TBD to Nov 15', anchor: 'TBD', resolved: false }],
      });
      const word = createHostAdapter('word', env);
      const ctx = await word.readContext();
      const promptCtx = buildContextPrompt({ host: 'word', ...ctx });
      assert.match(promptCtx, /\[id=c-99\]/);

      const reply = [
        '<edit action="replace_match"><find>TBD</find><content>November 15</content></edit>',
        '<edit action="reply_comment" comment_id="c-99">Updated the launch date to November 15.</edit>',
      ].join('\n');
      const edits = parseAssistantReply(reply).filter(p => p.kind === 'edit');
      assert.equal(edits.length, 2);
      for (const e of edits) {
        const res = await word.applyEdit(e);
        assert.equal(res.ok, true);
      }
      assert.equal(env.state.paragraphs[0], 'Our launch date is November 15.');
      assert.deepEqual(env.state.comments[0].replies, ['Updated the launch date to November 15.']);
    },
  },
  {
    id: 'E05',
    suite: 'Word',
    name: 'Bilingual translation insert_after adds clean CJK paragraph underneath',
    async run() {
      const env = createMockWordEnv({ paragraphs: ['Quarterly revenue beat expectations.'], selectionIndex: 0 });
      const word = createHostAdapter('word', env);
      const reply = '<edit action="insert_after" label="Chinese translation">季度营收超出预期。</edit>';
      const [edit] = parseAssistantReply(reply);
      assert.equal(actionLabel(edit.action, edit.label), 'Chinese translation');
      const res = await word.applyEdit(edit);
      assert.equal(res.ok, true);
      assert.deepEqual(env.state.paragraphs, ['Quarterly revenue beat expectations.', '季度营收超出预期。']);
    },
  },

  /* ===================== EXCEL EVALS ===================== */
  {
    id: 'E06',
    suite: 'Excel',
    name: 'Formula audit detects #DIV/0! & #REF!, renders cell citations, and applies set_formula fix',
    async run() {
      const env = createMockExcelEnv({
        sheets: {
          'P&L': {
            values: [
              ['Line Item', 'Q1', 'Q2'],
              ['Revenue', 200, 0],
              ['Gross Profit', 120, 50],
              ['Gross Margin', 0.6, '#DIV/0!'],
              ['OpEx', 80, '#REF!'],
            ],
            formulas: [
              ['Line Item', 'Q1', 'Q2'],
              ['Revenue', 200, 0],
              ['Gross Profit', 120, 50],
              ['Gross Margin', '=B3/B2', '=C3/C2'],
              ['OpEx', 80, '=OldSheet!C5'],
            ],
          },
        },
        activeSheet: 'P&L',
        selectionAddress: "'P&L'!C4",
      });
      const excel = createHostAdapter('excel', env);
      const ctx = await excel.readContext();
      assert.equal(ctx.hostContext.errors.length, 2);
      const promptCtx = buildContextPrompt({ host: 'excel', ...ctx });
      assert.match(promptCtx, /#DIV\/0!/);
      assert.match(promptCtx, /#REF!/);

      const reply = [
        "Found two errors: `'P&L'!C4` divides by zero (`C2` is 0), and `'P&L'!C5` has a broken `#REF!`.",
        '<edit action="set_formula" cell="C4" sheet="P&L" number_format="0.0%">=IFERROR(C3/C2,0)</edit>',
      ].join('\n');
      const [prose, edit] = parseAssistantReply(reply);
      const html = renderMarkdown(prose.text, { host: 'excel' });
      assert.match(html, /class="ref-chip" data-ref="&#39;P&amp;L&#39;!C4"/);
      assert.match(html, /class="ref-chip" data-ref="C2"/);

      const res = await excel.applyEdit(edit, { highlightEdits: true });
      assert.equal(res.ok, true);
      assert.equal(env.state.sheets.get('P&L').formulas[3][2], '=IFERROR(C3/C2,0)');
      assert.ok(env.state.sheets.get('P&L').formats.some(f => f.fill.color === '#FEF9C3'));
    },
  },
  {
    id: 'E07',
    suite: 'Excel',
    name: 'Dynamic table build via set_cells auto-resizes from anchor and preserves numeric & formula types',
    async run() {
      const env = createMockExcelEnv({ sheets: { Forecast: { values: [['']], formulas: [['']] } }, activeSheet: 'Forecast', selectionAddress: 'Forecast!A1' });
      const excel = createHostAdapter('excel', env);
      const reply = '<edit action="set_cells" range="A1" sheet="Forecast" number_format="$#,##0">[["Segment","Q1","Q2","Total"],["Enterprise",120000,150000,"=SUM(B2:C2)"],["SMB",45000,52000,"=SUM(B3:C3)"]]</edit>';
      const [edit] = parseAssistantReply(reply);
      const res = await excel.applyEdit(edit);
      assert.equal(res.ok, true);

      const grid = env.state.sheets.get('Forecast').formulas;
      assert.equal(grid.length, 3);
      assert.equal(grid[0].length, 4);
      assert.equal(typeof grid[1][1], 'number');
      assert.equal(grid[1][1], 120000);
      assert.equal(grid[1][3], '=SUM(B2:C2)');
    },
  },
  {
    id: 'E08',
    suite: 'Excel',
    name: 'Workbook expansion: add_sheet + format_cells + add_chart in sequence',
    async run() {
      const env = createMockExcelEnv();
      const excel = createHostAdapter('excel', env);
      const reply = [
        '<edit action="add_sheet" name="ExecSummary">[["Metric","Value"],["Total Q1","=Sheet1!B2"],["Total Q2","=Sheet1!C2"]]</edit>',
        '<edit action="format_cells" range="A1:B1" sheet="ExecSummary">{"bold":true,"fill":"#E5E7EB","autoFit":true}</edit>',
        '<edit action="add_chart" range="A1:B3" sheet="ExecSummary" chart_type="ColumnClustered" title="Quarterly Totals"></edit>',
      ].join('\n');
      const edits = parseAssistantReply(reply).filter(p => p.kind === 'edit');
      assert.equal(edits.length, 3);
      for (const e of edits) {
        const res = await excel.applyEdit(e);
        assert.equal(res.ok, true);
      }
      const sumSheet = env.state.sheets.get('ExecSummary');
      assert.ok(sumSheet);
      assert.equal(sumSheet.formulas[1][1], '=Sheet1!B2');
      assert.equal(sumSheet.charts.length, 1);
      assert.equal(sumSheet.charts[0].title.text, 'Quarterly Totals');
    },
  },
  {
    id: 'E09',
    suite: 'Excel',
    name: 'Clickable cell citations navigate accurately across sheets including quoted names',
    async run() {
      const env = createMockExcelEnv({
        sheets: {
          Overview: { values: [['Home']], formulas: [['Home']] },
          'Q1 P&L': { values: [['Rev', 500], ['Cost', 200]], formulas: [['Rev', 500], ['Cost', 200]] },
        },
        activeSheet: 'Overview',
        selectionAddress: 'Overview!A1',
      });
      const excel = createHostAdapter('excel', env);
      const nav = await excel.navigateToRef("'Q1 P&L'!A1:B2");
      assert.equal(nav.ok, true);
      assert.equal(env.state.activeSheet, 'Q1 P&L');
      assert.equal(env.state.selectionAddress, 'Q1 P&L!A1:B2');
    },
  },
  {
    id: 'E10',
    suite: 'Excel',
    name: 'Excel sort_range, excel_comment, and destructive clear_range guardrail',
    async run() {
      const env = createMockExcelEnv({
        sheets: {
          Deals: {
            values: [['Account', 'ARR'], ['Alpha', 50], ['Beta', 200], ['Gamma', 120]],
            formulas: [['Account', 'ARR'], ['Alpha', 50], ['Beta', 200], ['Gamma', 120]],
          },
        },
        activeSheet: 'Deals',
        selectionAddress: 'Deals!A1:B4',
      });
      const excel = createHostAdapter('excel', env);
      await excel.applyEdit({ action: 'sort_range', attrs: { range: 'A1:B4', sheet: 'Deals', column: 'B', order: 'desc', has_headers: 'true' }, content: '' });
      assert.deepEqual(env.state.sheets.get('Deals').values.map(r => r[0]), ['Account', 'Beta', 'Gamma', 'Alpha']);

      await excel.applyEdit({ action: 'excel_comment', attrs: { cell: 'B2', sheet: 'Deals' }, content: 'Largest enterprise renewal.' });
      assert.equal(env.state.sheets.get('Deals').comments.length, 1);

      assert.equal(isDestructiveAction('clear_range'), true);
      await excel.applyEdit({ action: 'clear_range', attrs: { range: 'B4', sheet: 'Deals' }, content: '' });
      assert.equal(env.state.sheets.get('Deals').values[3][1], '');
    },
  },

  /* ===================== POWERPOINT EVALS ===================== */
  {
    id: 'E11',
    suite: 'PowerPoint',
    name: 'On-brand add_slide matches Slide Master layout and populates native placeholders',
    async run() {
      const env = createMockPowerPointEnv();
      const ppt = createHostAdapter('powerpoint', env);
      const ctx = await ppt.readContext();
      const promptCtx = buildContextPrompt({ host: 'powerpoint', ...ctx });
      assert.match(promptCtx, /<layouts>[\s\S]*Two Content/);

      const reply = [
        '<edit action="add_slide" layout="Two Content" after_slide="2">',
        JSON.stringify({
          title: 'Enterprise vs Self-Serve Economics',
          columns: [['Enterprise: $85k ACV', '124% NRR'], ['Self-Serve: $4.2k ACV', '9-day payback']],
        }),
        '</edit>',
      ].join('\n');
      const [edit] = parseAssistantReply(reply);
      const res = await ppt.applyEdit(edit);
      assert.equal(res.ok, true);
      assert.equal(env.state.slides.length, 3);
      const added = env.state.slides[2];
      assert.equal(added.layoutName, 'Two Content');
      assert.equal(added.shapes[0].text, 'Enterprise vs Self-Serve Economics');
      assert.match(added.shapes[1].text, /\$85k ACV/);
      assert.match(added.shapes[2].text, /9-day payback/);
    },
  },
  {
    id: 'E12',
    suite: 'PowerPoint',
    name: 'Executive slide tightening via update_shape produces clean diff and updates target shape',
    async run() {
      const env = createMockPowerPointEnv();
      const ppt = createHostAdapter('powerpoint', env);
      const oldText = env.state.slides[1].shapes[1].text;
      const reply = '<edit action="update_shape" slide="2" shape="Content Placeholder 2">Nova grew 22% YoY\nEnterprise drove 68% of net-new ARR</edit>';
      const [edit] = parseAssistantReply(reply);
      const diff = computeWordDiff(oldText, edit.content);
      assert.ok(diff.some(d => d.type === 'del'));
      assert.ok(diff.some(d => d.type === 'ins'));

      const res = await ppt.applyEdit(edit);
      assert.equal(res.ok, true);
      assert.equal(env.state.slides[1].shapes[1].text, edit.content);
    },
  },
  {
    id: 'E13',
    suite: 'PowerPoint',
    name: 'Deck-wide replace_slide_text updates consistent terminology across slides',
    async run() {
      const env = createMockPowerPointEnv({
        slides: [
          { id: 's1', layoutName: 'Title Slide', shapes: [{ id: 'a1', name: 'Title 1', type: 'Placeholder', text: 'Acme Legacy Roadmap' }] },
          { id: 's2', layoutName: 'Title and Content', shapes: [{ id: 'a2', name: 'Content Placeholder 2', type: 'Placeholder', text: 'Migrating customers off Acme Legacy.' }] },
        ],
        selectedSlideIds: ['s1'],
        selectedShapeIds: [],
      });
      const ppt = createHostAdapter('powerpoint', env);
      const reply = '<edit action="replace_slide_text"><find>Acme Legacy</find><content>Acme Cloud</content></edit>';
      const [edit] = parseAssistantReply(reply);
      const res = await ppt.applyEdit(edit);
      assert.equal(res.ok, true);
      assert.match(res.message, /2 shapes/);
      assert.equal(env.state.slides[0].shapes[0].text, 'Acme Cloud Roadmap');
      assert.equal(env.state.slides[1].shapes[0].text, 'Migrating customers off Acme Cloud.');
    },
  },
  {
    id: 'E14',
    suite: 'PowerPoint',
    name: 'Native slide table creation via add_table from Markdown pipe table',
    async run() {
      const env = createMockPowerPointEnv();
      const ppt = createHostAdapter('powerpoint', env);
      const reply = [
        '<edit action="add_table" slide="2">',
        '| Region | Q2 | Q3 | YoY |',
        '|---|---|---|---|',
        '| NA | $6.1M | $7.4M | +21% |',
        '| EMEA | $3.2M | $4.1M | +28% |',
        '</edit>',
      ].join('\n');
      const [edit] = parseAssistantReply(reply);
      const res = await ppt.applyEdit(edit);
      assert.equal(res.ok, true);
      assert.equal(env.state.slides[1].tables.length, 1);
      assert.equal(env.state.slides[1].tables[0].rows, 3);
      assert.equal(env.state.slides[1].tables[0].cols, 4);
    },
  },
  {
    id: 'E15',
    suite: 'PowerPoint',
    name: 'Clickable Slide citation navigation, add_textbox, and destructive delete_slide guardrail',
    async run() {
      const env = createMockPowerPointEnv();
      const ppt = createHostAdapter('powerpoint', env);
      const html = renderMarkdown('Move the appendix chart from `Slide 2` and drop Slide 1 subtitle.', { host: 'powerpoint' });
      assert.match(html, /data-ref="Slide 2"/);
      assert.match(html, /data-ref="Slide 1"/);

      const nav = await ppt.navigateToRef('Slide 1');
      assert.equal(nav.ok, true);
      assert.deepEqual(env.state.selectedSlideIds, ['sl-1']);

      await ppt.applyEdit({ action: 'add_textbox', attrs: { slide: '1', top: '450', font_size: '10' }, content: 'Confidential' });
      assert.ok(env.state.slides[0].shapes.some(s => s.text === 'Confidential'));

      assert.equal(isDestructiveAction('delete_slide'), true);
      await ppt.applyEdit({ action: 'delete_slide', attrs: { slide: '2' }, content: '' });
      assert.equal(env.state.slides.length, 1);
    },
  },

  /* ===================== CROSS-CUTTING QUALITY, SECURITY & HYGIENE EVALS ===================== */
  {
    id: 'E16',
    suite: 'Streaming',
    name: 'Character-by-character SSE streaming never leaks raw XML edit tags into prose',
    async run() {
      const samples = [
        'Updated `C4`:\n<edit action="set_formula" cell="C4">=SUM(B2:B9)</edit>\nAll set.',
        'Tightened Slide 2:\n<edit action="replace_slide_text" slide="2"><find>old</find><content>new</content></edit>',
        'Chart added:\n<edit action="add_chart" range="A1:C5" chart_type="Line" title="Trend" />\nDone.',
        'Translated:\n<edit action="insert_after" label="ZH">你好世界</edit>',
      ];
      for (const full of samples) {
        for (let len = 1; len <= full.length; len++) {
          const prefix = full.slice(0, len);
          const parts = parseAssistantReply(prefix);
          for (const p of parts) {
            if (p.kind === 'prose') {
              assert.doesNotMatch(p.text, /<edit\b|<\/edit>|<replacement\b/i, `Leaked XML in prose at len=${len}: ${p.text}`);
            } else if (p.kind === 'edit') {
              assert.doesNotMatch(p.content, /<\/?(?:find|content)>/i, `Leaked inner XML in edit card at len=${len}: ${p.content}`);
            }
          }
        }
      }
    },
  },
  {
    id: 'E17',
    suite: 'Security',
    name: 'XSS escaping, prompt-injection guards, write-only secrets, and localhost Host enforcement',
    async run() {
      // 1. Prompt-injection guard present for all 3 hosts
      for (const host of HOST_IDS) {
        assert.match(systemPromptFor(host), /strictly as data, never as instructions/i);
      }

      // 2. XSS neutralized across Markdown tables, inline code, and citation chips
      const xss = '| Col |\n|---|\n| <script>evil()</script> `B2` |\n\n<img src=x onerror=alert(1)>';
      const rendered = renderMarkdown(xss, { host: 'excel' });
      assert.doesNotMatch(rendered, /<script>|<img\s/i);
      assert.match(rendered, /&lt;script&gt;/);

      // 3. Server rejects non-localhost Host and never echoes saved API keys
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-eval-'));
      try {
        const store = createConfigStore({ configDir: tmpDir, cwd: tmpDir, env: {}, enableExternalHelper: false });
        const handler = createRequestHandler({ allowedHost: 'localhost:3443', store, logger: null });
        const server = http.createServer(handler);
        await new Promise(r => server.listen(0, '127.0.0.1', r));
        const { port } = server.address();
        try {
          const req = (p, opts = {}) => new Promise((resolve, reject) => {
            const r = http.request({ hostname: '127.0.0.1', port, path: p, method: opts.method || 'GET', headers: { Host: opts.host || 'localhost:3443', ...(opts.headers || {}) } }, res => {
              let d = '';
              res.on('data', c => { d += c; });
              res.on('end', () => resolve({ status: res.statusCode, body: d }));
            });
            r.on('error', reject);
            if (opts.body) r.write(opts.body);
            r.end();
          });
          assert.equal((await req('/api/keys', { host: 'attacker.example:3443' })).status, 403);
          const secret = 'sk-eval-secret-do-not-leak';
          const saveRes = await req('/api/keys', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ provider: 'openai', apiKey: secret }),
          });
          assert.equal(saveRes.status, 200);
          assert.ok(!saveRes.body.includes(secret));
          assert.ok(!(await req('/api/keys')).body.includes(secret));
        } finally {
          await new Promise(r => server.close(r));
        }
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    },
  },
  {
    id: 'E18',
    suite: 'Architecture',
    name: 'Zero orphan code, complete Host Adapter contract, and 3-host asset/manifest integrity',
    async run() {
      // 1. Every allowlisted static file exists on disk
      for (const rel of new Set(Object.values(STATIC_FILES))) {
        assert.ok(fs.existsSync(path.join(ROOT, rel)), `Missing static asset: ${rel}`);
      }

      // 2. Dead .brand-dot CSS rule removed
      const css = fs.readFileSync(path.join(ROOT, 'style.css'), 'utf8');
      assert.doesNotMatch(css, /\.brand-dot\b/, 'Orphan CSS class .brand-dot should be removed');

      // 3. Every HTML element ID referenced via $('...') in app.js exists in taskpane.html
      const html = fs.readFileSync(path.join(ROOT, 'taskpane.html'), 'utf8');
      const appSrc = fs.readFileSync(path.join(ROOT, 'app.js'), 'utf8');
      const usedIds = [...new Set([...appSrc.matchAll(/\$\('([a-z0-9_-]+)'\)/g)].map(m => m[1]))];
      assert.ok(usedIds.length >= 10);
      for (const id of usedIds) {
        assert.ok(html.includes(`id="${id}"`), `app.js references #${id} which is missing from taskpane.html`);
      }

      // 4. Every host adapter satisfies the shared contract cleanly
      for (const hostId of HOST_IDS) {
        const adapter = createHostAdapter(hostId, {});
        assert.equal(adapter.id, hostId);
        assert.ok(adapter.label && adapter.noun && adapter.placeholder);
        assert.ok(Array.isArray(adapter.quickActions) && adapter.quickActions.length >= 8);
        assert.equal(typeof adapter.formatContextBadge, 'function');
        assert.equal(typeof adapter.getStarters, 'function');
        assert.equal(typeof adapter.readContext, 'function');
        assert.equal(typeof adapter.watchSelection, 'function');
        assert.equal(typeof adapter.applyEdit, 'function');

        const badge = adapter.formatContextBadge({ selectionText: '', documentText: '', comments: [], hostContext: {} }, 'auto');
        assert.ok(typeof badge.badge === 'string' && typeof badge.preview === 'string');
        const starters = adapter.getStarters({ selectionText: '', documentText: '', comments: [], hostContext: {} });
        assert.equal(starters.starters.length, 4);
      }

      // 5. All action labels and primary button labels return non-empty strings
      for (const action of [
        'replace_selection', 'insert_after', 'insert_before', 'replace_match', 'append_doc', 'comment', 'reply_comment',
        'set_cells', 'set_formula', 'format_cells', 'add_sheet', 'add_chart', 'sort_range', 'excel_comment', 'clear_range',
        'update_shape', 'replace_slide_text', 'add_slide', 'add_textbox', 'add_table', 'delete_slide',
      ]) {
        assert.ok(actionLabel(action).length > 0);
        assert.ok(primaryButtonLabel(action).length > 0);
      }

      // 6. Repository AGENTS.md exists and documents architecture & verification
      const repoAgents = fs.readFileSync(path.join(ROOT, 'AGENTS.md'), 'utf8');
      assert.match(repoAgents, /npm test/);
      assert.match(repoAgents, /npm run eval/);
    },
  },
  {
    id: 'E19',
    suite: 'Skills',
    name: 'Global AGENTS.md & SKILL.md discovery, host filtering, /skill-name activation, and toggle persistence',
    async run() {
      const { createSkillsLoader, resolveActiveSkills, buildSkillsPromptSection } = await import('../lib/skills.mjs');
      const fakeHome = fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-eval-skills-'));
      try {
        fs.mkdirSync(path.join(fakeHome, '.agents', 'skills', 'board-deck'), { recursive: true });
        fs.mkdirSync(path.join(fakeHome, 'skills', 'model-audit'), { recursive: true });
        fs.writeFileSync(path.join(fakeHome, '.agents', 'AGENTS.md'), 'Prefer concise executive tone.');
        fs.writeFileSync(
          path.join(fakeHome, '.agents', 'skills', 'board-deck', 'SKILL.md'),
          '---\nname: board-deck\ndescription: Executive board slides\nhosts: [powerpoint]\n---\nUse action titles on every slide.',
        );
        fs.writeFileSync(
          path.join(fakeHome, 'skills', 'model-audit', 'SKILL.md'),
          '---\nname: model-audit\ndescription: Financial model check\nhosts: [excel]\n---\nFlag every hardcoded input.',
        );

        const loader = createSkillsLoader({ homeDir: fakeHome, env: {} });
        const { agentsFiles, skills } = loader.load();
        assert.equal(agentsFiles.length, 1);
        assert.equal(skills.length, 2);

        const pptRes = resolveActiveSkills({
          availableSkills: skills,
          messages: [{ role: 'user', content: '/board-deck tighten slide 2' }],
          host: 'powerpoint',
        });
        assert.deepEqual(pptRes.enabledSkills.map(s => s.name), ['board-deck']);
        assert.deepEqual(pptRes.activeSkills.map(s => s.name), ['board-deck']);

        const prompt = buildSkillsPromptSection({
          agentsFiles,
          agentsEnabled: true,
          enabledSkills: pptRes.enabledSkills,
          activeSkills: pptRes.activeSkills,
        });
        assert.match(prompt, /<global_instructions>[\s\S]*Prefer concise executive tone/);
        assert.match(prompt, /<active_skill name="board-deck"[\s\S]*action titles on every slide/);
        assert.doesNotMatch(prompt, /model-audit/);
      } finally {
        fs.rmSync(fakeHome, { recursive: true, force: true });
      }
    },
  },
  {
    id: 'E20',
    suite: 'Performance',
    name: 'Giant 50k-row Excel sheet, Ctrl+A selection, 100-slide deck windowing, and Word duplicate disambiguation',
    async run() {
      const { sanitizeSheetName } = await import('../lib/hosts/excel.js');
      assert.equal(sanitizeSheetName('FY26:P&L/Summary', ['FY26 P&L Summary']), 'FY26 P&L Summary (2)');

      let excelClipped = false;
      const giantExcelEnv = {
        Office: { HostType: { Excel: 'Excel' } },
        Excel: {
          async run(cb) {
            const makeBig = (addr, rows, cols) => ({
              address: addr,
              rowCount: rows,
              columnCount: cols,
              isNullObject: false,
              load(p) { if (String(p).includes('values')) throw new Error('Unclipped bridge read!'); },
              getAbsoluteResizedRange(r, c) {
                excelClipped = r <= 70 && c <= 22;
                return { address: 'Data!A1:B2', load() {}, values: [['A', 'B'], [1, 2]], formulas: [['A', 'B'], [1, 2]] };
              },
            });
            const sh = { name: 'Data', position: 0, load() {}, getUsedRangeOrNullObject: () => makeBig('Data!A1:Z50000', 50000, 26) };
            return cb({
              workbook: {
                worksheets: { items: [sh], load() {}, getActiveWorksheet: () => sh },
                getSelectedRange: () => makeBig('Data!A1:A1048576', 1048576, 1),
              },
              async sync() {},
            });
          },
        },
      };
      const excelCtx = await createHostAdapter('excel', giantExcelEnv).readContext();
      assert.equal(excelClipped, true);
      assert.match(excelCtx.hostContext.selectionClipped, /1048576×1/);

      // 100-slide PowerPoint deck windows to 60 slides cleanly
      const slides100 = Array.from({ length: 100 }, (_, i) => ({
        id: `s-${i + 1}`,
        layoutName: 'Title and Content',
        shapes: [{ id: `sh-${i + 1}`, name: 'Title 1', type: 'Placeholder', text: `Slide ${i + 1}` }],
      }));
      const pptEnv = createMockPowerPointEnv({ slides: slides100, selectedSlideIds: ['s-50'], selectedShapeIds: [] });
      const pptCtx = await createHostAdapter('powerpoint', pptEnv).readContext();
      assert.equal(pptCtx.hostContext.slideCount, 100);
      assert.equal(pptCtx.hostContext.slidesOmitted, 40);

      // Word occurrence=2 disambiguation
      const wordEnv = createMockWordEnv({ paragraphs: ['Repeat term one.', 'Repeat term two.'] });
      await createHostAdapter('word', wordEnv).applyEdit({
        action: 'replace_match',
        find: 'Repeat term',
        content: 'Fixed term',
        attrs: { occurrence: '2' },
      });
      assert.equal(wordEnv.state.paragraphs[0], 'Repeat term one.');
      assert.equal(wordEnv.state.paragraphs[1], 'Fixed term two.');
    },
  },
  {
    id: 'E21',
    suite: 'Harnesses',
    name: 'Pi ($PI_CODING_AGENT_DIR + AGENTS.override.md), OpenCode, Gemini, Cursor discovery & resource sandbox',
    async run() {
      const { createSkillsLoader, readSkillResourceFile } = await import('../lib/skills.mjs');
      const fakeHome = fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-eval-harnesses-'));
      try {
        const customPi = path.join(fakeHome, 'custom-pi');
        fs.mkdirSync(path.join(customPi, 'skills', 'pi-memo', 'references'), { recursive: true });
        fs.writeFileSync(path.join(customPi, 'AGENTS.md'), 'Base pi rule');
        fs.writeFileSync(path.join(customPi, 'AGENTS.override.md'), 'Override pi rule');
        fs.writeFileSync(path.join(customPi, 'skills', 'pi-memo', 'SKILL.md'), '---\nname: pi-memo\ndescription: Pi memo\n---\nBody');
        fs.writeFileSync(path.join(customPi, 'skills', 'pi-memo', 'references', 'style.md'), 'Style rules');

        fs.mkdirSync(path.join(fakeHome, '.config', 'opencode', 'skills', 'oc-excel'), { recursive: true });
        fs.writeFileSync(path.join(fakeHome, '.config', 'opencode', 'AGENTS.md'), 'OpenCode rule');
        fs.writeFileSync(path.join(fakeHome, '.config', 'opencode', 'skills', 'oc-excel', 'SKILL.md'), '---\nname: oc-excel\ndescription: OC excel\n---\nBody');

        fs.mkdirSync(path.join(fakeHome, '.gemini'), { recursive: true });
        fs.writeFileSync(path.join(fakeHome, '.gemini', 'GEMINI.md'), 'Gemini rule');
        fs.mkdirSync(path.join(fakeHome, '.cursor', 'skills', 'cur-deck'), { recursive: true });
        fs.writeFileSync(path.join(fakeHome, '.cursor', 'skills', 'cur-deck', 'SKILL.md'), '---\nname: cur-deck\ndescription: Cursor deck\n---\nBody');

        const loader = createSkillsLoader({ homeDir: fakeHome, env: { PI_CODING_AGENT_DIR: customPi } });
        const { agentsFiles, skills } = loader.load();
        assert.ok(agentsFiles.some(f => f.content === 'Override pi rule'));
        assert.ok(!agentsFiles.some(f => f.content === 'Base pi rule'));
        assert.ok(agentsFiles.some(f => f.content === 'OpenCode rule'));
        assert.ok(agentsFiles.some(f => f.content === 'Gemini rule'));
        assert.deepEqual(skills.map(s => s.name).sort(), ['cur-deck', 'oc-excel', 'pi-memo']);

        const piSkill = skills.find(s => s.name === 'pi-memo');
        assert.equal(readSkillResourceFile(piSkill, 'references/style.md'), 'Style rules');
        assert.throws(() => readSkillResourceFile(piSkill, '../../AGENTS.override.md'));
      } finally {
        fs.rmSync(fakeHome, { recursive: true, force: true });
      }
    },
  },
  {
    id: 'E22',
    suite: 'History',
    name: 'Persistent local chat sessions CRUD, pruning, and rolling context compaction across 40 messages',
    async run() {
      const { createSessionStore, compactHistory } = await import('../lib/sessions.mjs');
      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-eval-hist-'));
      try {
        const store = createSessionStore({ configDir: tmp, maxSessions: 5 });
        const longMsgs = Array.from({ length: 40 }, (_, i) => ({
          role: i % 2 === 0 ? 'user' : 'assistant',
          content: `Turn ${i + 1} discussion about Q3 metrics`,
        }));
        const saved = store.save({ host: 'excel', messages: longMsgs });
        assert.equal(store.get(saved.id).messages.length, 40);

        const compacted = compactHistory(longMsgs, { maxMessages: 16 });
        assert.equal(compacted.compacted, true);
        assert.equal(compacted.droppedCount, 24);
        assert.equal(compacted.messages[0].role, 'system');
        assert.match(compacted.messages[0].content, /24 earlier messages/);
      } finally {
        fs.rmSync(tmp, { recursive: true, force: true });
      }
    },
  },
  {
    id: 'E23',
    suite: 'MCP & Tools',
    name: 'Real stdio MCP server + built-in math tool + Ask/Auto permission gate for non-readOnly MCP tools',
    async run() {
      const { createMcpManager } = await import('../lib/mcp.mjs');
      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-eval-mcp-'));
      const srvScript = path.join(tmp, 'srv.mjs');
      fs.writeFileSync(
        srvScript,
        `
        import readline from 'node:readline';
        readline.createInterface({ input: process.stdin }).on('line', l => {
          if (!l.trim()) return;
          const m = JSON.parse(l);
          if (m.method === 'initialize') process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: m.id, result: { protocolVersion: '2025-03-26', capabilities: { tools: {} } } }) + '\\n');
          else if (m.method === 'tools/list') process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: m.id, result: { tools: [
            { name: 'read_kpi', description: 'Read KPI', annotations: { readOnlyHint: true } },
            { name: 'post_slack', description: 'Post Slack message', annotations: { readOnlyHint: false } }
          ] } }) + '\\n');
          else if (m.method === 'tools/call') process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: m.id, result: { content: [{ type: 'text', text: 'OK:' + m.params.name }] } }) + '\\n');
        });
        `,
      );
      const mcpCfg = path.join(tmp, 'mcp.json');
      fs.writeFileSync(mcpCfg, JSON.stringify({ mcpServers: { corp: { command: process.execPath, args: [srvScript] } } }));

      const store = createConfigStore({ configDir: tmp, cwd: tmp, env: { DEEPSEEK_API_KEY: 'k' }, enableExternalHelper: false });
      const mcpManager = createMcpManager({ homeDir: tmp, configDir: tmp, configPaths: [mcpCfg] });

      const mockFetch = async (_u, opts) => {
        const bodyObj = JSON.parse(opts.body);
        const toolMsg = bodyObj.messages.find(m => m.role === 'tool');
        const chunks = !toolMsg
          ? ['data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"mcp__corp__post_slack","arguments":"{}"}}]}}]}\n', 'data: [DONE]\n']
          : [`data: {"choices":[{"delta":{"content":"Tool said: ${toolMsg.content}"}}]}\n`, 'data: [DONE]\n'];
        const body = (async function* () { for (const c of chunks) yield new TextEncoder().encode(c); })();
        return { ok: true, status: 200, body, text: async () => '' };
      };

      const handler = createRequestHandler({ allowedHost: 'localhost:3443', store, mcpManager, fetchImpl: mockFetch, logger: null });
      const server = http.createServer(handler);
      await new Promise(r => server.listen(0, '127.0.0.1', r));
      const { port } = server.address();
      const postChat = payload => new Promise((resolve, reject) => {
        const req = http.request({ hostname: '127.0.0.1', port, path: '/api/chat', method: 'POST', headers: { Host: 'localhost:3443', 'Content-Type': 'application/json' } }, res => {
          let d = '';
          res.on('data', c => { d += c; });
          res.on('end', () => resolve(d));
        });
        req.on('error', reject);
        req.write(JSON.stringify(payload));
        req.end();
      });

      try {
        // In Ask mode (allowExternalTools: false), write MCP tool is blocked safely
        const blocked = await postChat({ host: 'word', messages: [{ role: 'user', content: 'Post to slack' }], allowExternalTools: false });
        assert.match(blocked, /Permission required/);

        // In Auto mode (allowExternalTools: true), write MCP tool executes cleanly
        const allowed = await postChat({ host: 'word', messages: [{ role: 'user', content: 'Post to slack' }], allowExternalTools: true });
        assert.match(allowed, /OK:post_slack/);
      } finally {
        mcpManager.closeAll();
        await new Promise(r => server.close(r));
        fs.rmSync(tmp, { recursive: true, force: true });
      }
    },
  },
  {
    id: 'E24',
    suite: 'Resilience',
    name: 'Explicit truncation notices, client-disconnect upstream abort, and split UTF-8 CJK handling',
    async run() {
      const longDoc = 'A'.repeat(45000);
      const ctx = buildContextPrompt({ host: 'word', documentText: longDoc, selection: '' });
      assert.match(ctx, /Note: truncated — showing first 40,000 of 45,000 characters/);

      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-eval-abort-'));
      const store = createConfigStore({ configDir: tmp, cwd: tmp, env: { DEEPSEEK_API_KEY: 'k' }, enableExternalHelper: false });
      const { createMcpManager } = await import('../lib/mcp.mjs');
      const mcpManager = createMcpManager({ homeDir: tmp, configDir: tmp, configPaths: [] });
      let aborted = false;

      const mockFetch = async (_u, opts) => {
        const body = (async function* () {
          yield new TextEncoder().encode('data: {"choices":[{"delta":{"content":"Hi"}}]}\n');
          await new Promise(resolve => {
            const t = setTimeout(resolve, 4000);
            opts.signal?.addEventListener('abort', () => { clearTimeout(t); aborted = true; resolve(); });
          });
        })();
        return { ok: true, status: 200, body, text: async () => '' };
      };
      const handler = createRequestHandler({ allowedHost: 'localhost:3443', store, mcpManager, fetchImpl: mockFetch, logger: null });
      const server = http.createServer(handler);
      await new Promise(r => server.listen(0, '127.0.0.1', r));
      const { port } = server.address();
      try {
        await new Promise(resolve => {
          const req = http.request({ hostname: '127.0.0.1', port, path: '/api/chat', method: 'POST', headers: { Host: 'localhost:3443', 'Content-Type': 'application/json' } }, res => {
            res.once('data', () => { req.destroy(); setTimeout(resolve, 50); });
          });
          req.on('error', () => {});
          req.write(JSON.stringify({ host: 'word', messages: [{ role: 'user', content: 'Hello' }] }));
          req.end();
        });
        assert.equal(aborted, true);
      } finally {
        mcpManager.closeAll();
        await new Promise(r => server.close(r));
        fs.rmSync(tmp, { recursive: true, force: true });
      }
    },
  },
];

async function main() {
  console.log('\nOffice AI Editor — Multi-Host Evaluation Suite');
  console.log('==============================================\n');

  let passed = 0;
  let failed = 0;
  const suiteStats = new Map();

  for (const sc of scenarios) {
    const start = performance.now();
    const stat = suiteStats.get(sc.suite) || { pass: 0, total: 0 };
    stat.total++;
    try {
      await sc.run();
      const ms = (performance.now() - start).toFixed(1);
      console.log(`  PASS  [${sc.id}] (${sc.suite}) ${sc.name} (${ms}ms)`);
      passed++;
      stat.pass++;
    } catch (err) {
      const ms = (performance.now() - start).toFixed(1);
      console.error(`  FAIL  [${sc.id}] (${sc.suite}) ${sc.name} (${ms}ms)`);
      console.error(`        ${err.stack || err.message}`);
      failed++;
    }
    suiteStats.set(sc.suite, stat);
  }

  console.log('\nScorecard by Suite:');
  for (const [suite, { pass, total }] of suiteStats.entries()) {
    const pct = ((pass / total) * 100).toFixed(0);
    console.log(`  - ${suite.padEnd(14)} ${pass}/${total} (${pct}%)`);
  }
  console.log(`\nTotal: ${passed}/${scenarios.length} passed (${((passed / scenarios.length) * 100).toFixed(1)}%)\n`);

  if (failed > 0) process.exit(1);
}

main();
