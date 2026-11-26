import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseAssistantReply,
  actionLabel,
  primaryButtonLabel,
  isDestructiveAction,
  renderMarkdown,
  computeWordDiff,
  countWords,
  splitParagraphs,
} from '../lib/format.js';

test('parseAssistantReply parses structured edit tags and prose', () => {
  const reply = [
    'Here is the Chinese translation to place underneath.',
    '<edit action="insert_after" label="Chinese translation">',
    '这是一个文本测试。我只是把它放在这里。',
    '</edit>',
    'Want a more formal tone?',
  ].join('\n');

  const parts = parseAssistantReply(reply);
  assert.equal(parts.length, 3);
  assert.equal(parts[0].kind, 'prose');
  assert.equal(parts[1].kind, 'edit');
  assert.equal(parts[1].action, 'insert_after');
  assert.equal(parts[1].label, 'Chinese translation');
  assert.equal(parts[1].content, '这是一个文本测试。我只是把它放在这里。');
  assert.equal(parts[1].streaming, false);
  assert.equal(parts[2].kind, 'prose');
});

test('parseAssistantReply parses replace_match with find and content', () => {
  const reply = `<edit action="replace_match">
<find>old sentence here</find>
<content>sharper sentence here</content>
</edit>`;

  const parts = parseAssistantReply(reply);
  assert.equal(parts.length, 1);
  assert.equal(parts[0].action, 'replace_match');
  assert.equal(parts[0].find, 'old sentence here');
  assert.equal(parts[0].content, 'sharper sentence here');
});

test('parseAssistantReply parses reply_comment with comment_id', () => {
  const reply = `<edit action="reply_comment" comment_id="c-42" label="Reply to Tristan">Tightened the wording as requested.</edit>`;
  const parts = parseAssistantReply(reply);
  assert.equal(parts.length, 1);
  assert.equal(parts[0].action, 'reply_comment');
  assert.equal(parts[0].commentId, 'c-42');
  assert.equal(parts[0].content, 'Tightened the wording as requested.');
});

test('parseAssistantReply parses Excel and PowerPoint actions with rich attributes', () => {
  const excelReply = [
    'Fixed the margin formula in `C10`:',
    '<edit action="set_formula" cell="C10" sheet="P&L" number_format="0.0%">=IFERROR(C9/C5,0)</edit>',
    '<edit action="add_chart" range="A1:C6" chart_type="ColumnClustered" title="Revenue" sheet="P&L"></edit>',
  ].join('\n');

  const eParts = parseAssistantReply(excelReply);
  assert.equal(eParts.length, 3);
  assert.equal(eParts[1].action, 'set_formula');
  assert.equal(eParts[1].attrs.cell, 'C10');
  assert.equal(eParts[1].attrs.sheet, 'P&L');
  assert.equal(eParts[1].attrs.number_format, '0.0%');
  assert.equal(eParts[1].content, '=IFERROR(C9/C5,0)');
  assert.equal(eParts[2].action, 'add_chart');
  assert.equal(eParts[2].attrs.chart_type, 'ColumnClustered');
  assert.equal(actionLabel('set_formula', '', eParts[1].attrs), 'Formula P&L!C10');
  assert.equal(primaryButtonLabel('add_chart'), 'Create chart');

  const pptReply = [
    '<edit action="add_slide" layout="Title and Content" after_slide="2">',
    'Title: Q3 Results',
    '- Revenue grew 32%',
    '</edit>',
    '<edit action="delete_slide" slide="5"></edit>',
  ].join('\n');
  const pParts = parseAssistantReply(pptReply);
  assert.equal(pParts.length, 2);
  assert.equal(pParts[0].action, 'add_slide');
  assert.equal(pParts[0].attrs.layout, 'Title and Content');
  assert.equal(pParts[0].attrs.after_slide, '2');
  assert.equal(actionLabel('add_slide', '', pParts[0].attrs), 'Add slide · Title and Content');
  assert.equal(pParts[1].action, 'delete_slide');
  assert.equal(isDestructiveAction('delete_slide'), true);
  assert.equal(isDestructiveAction('clear_range'), true);
  assert.equal(isDestructiveAction('set_cells'), false);
});

test('parseAssistantReply supports legacy replacement tag and streaming incomplete tags', () => {
  const legacy = parseAssistantReply('<replacement>\nHello world\n</replacement>');
  assert.equal(legacy.length, 1);
  assert.equal(legacy[0].action, 'replace_selection');
  assert.equal(legacy[0].content, 'Hello world');
  assert.equal(legacy[0].streaming, false);

  const streaming = parseAssistantReply('Sure:\n<edit action="insert_after">\nFirst line');
  assert.equal(streaming.length, 2);
  assert.equal(streaming[1].kind, 'edit');
  assert.equal(streaming[1].streaming, true);
  assert.equal(streaming[1].content, 'First line');

  const partialOpen = parseAssistantReply('Working on it <edit act');
  assert.equal(partialOpen.length, 1);
  assert.equal(partialOpen[0].kind, 'prose');
  assert.equal(partialOpen[0].text.trim(), 'Working on it');

  // Self-closing bodyless edit tags (common in Excel/PowerPoint)
  const selfClosing = parseAssistantReply('Removing slide:\n<edit action="delete_slide" slide="4" />\nDone.');
  assert.equal(selfClosing.length, 3);
  assert.equal(selfClosing[1].action, 'delete_slide');
  assert.equal(selfClosing[1].attrs.slide, '4');
  assert.equal(selfClosing[1].streaming, false);
  assert.equal(selfClosing[2].kind, 'prose');

  // Streaming unclosed <find> / <content> never flashes raw XML tags
  const midFind = parseAssistantReply('<edit action="replace_match">\n<find>old text');
  assert.equal(midFind[0].streaming, true);
  assert.doesNotMatch(midFind[0].content, /<find>/);
  const midContent = parseAssistantReply('<edit action="replace_match"><find>old</find><content>new text');
  assert.equal(midContent[0].find, 'old');
  assert.equal(midContent[0].content, 'new text');
});

test('renderMarkdown formats headings, lists, tables, bold, code and escapes HTML', () => {
  const md = [
    '## Summary',
    '',
    'This has **bold**, *italic*, `code`, and <script>alert(1)</script>.',
    '',
    '| Metric | Q1 | Q2 |',
    '|---|---:|---:|',
    '| Revenue | $100 | $125 |',
    '| <img src=x onerror=alert(1)> | 10% | 15% |',
    '',
    '- Point one',
    '- Point two',
  ].join('\n');

  const html = renderMarkdown(md);
  assert.match(html, /<h3>Summary<\/h3>/);
  assert.match(html, /<strong>bold<\/strong>/);
  assert.match(html, /<em>italic<\/em>/);
  assert.match(html, /<code>code<\/code>/);
  assert.doesNotMatch(html, /<script>/);
  assert.doesNotMatch(html, /<img /);
  assert.match(html, /&lt;script&gt;/);
  assert.match(html, /<table class="md-table">/);
  assert.match(html, /<th>Metric<\/th>/);
  assert.match(html, /<td>\$125<\/td>/);
  assert.match(html, /<ul><li>Point one<\/li><li>Point two<\/li><\/ul>/);
});

test('renderMarkdown turns Excel cell citations and PowerPoint slide references into clickable chips', () => {
  const excelHtml = renderMarkdown("Check `C10`, `'P&L'!B2:D5`, [Sheet2!A4], `Q1`, `FY26`, `Sheet1!Q1`, and `SUM(A1:A5)`.", { host: 'excel' });
  assert.match(excelHtml, /class="ref-chip" data-ref="C10"/);
  assert.match(excelHtml, /class="ref-chip" data-ref="&#39;P&amp;L&#39;!B2:D5"/);
  assert.match(excelHtml, /class="ref-chip" data-ref="Sheet2!A4"/);
  assert.match(excelHtml, /class="ref-chip" data-ref="Sheet1!Q1"/);
  // Finance shorthand periods must stay <code>, not false-positive cell links
  assert.match(excelHtml, /<code>Q1<\/code>/);
  assert.match(excelHtml, /<code>FY26<\/code>/);
  // Formula code span should stay regular <code>, not a cell ref chip
  assert.match(excelHtml, /<code>SUM\(A1:A5\)<\/code>/);

  const pptHtml = renderMarkdown('Tighten `Slide 2` and merge Slide 4 with [Slide 5].', { host: 'powerpoint' });
  assert.match(pptHtml, /data-ref="Slide 2"/);
  assert.match(pptHtml, /data-ref="Slide 4"/);
  assert.match(pptHtml, /data-ref="Slide 5"/);

  // Word host should NOT turn plain text into Excel/PPT chips
  const wordHtml = renderMarkdown('See Slide 2 and `B4`.', { host: 'word' });
  assert.doesNotMatch(wordHtml, /ref-chip/);
});

test('computeWordDiff handles English and Chinese edits', () => {
  const diff = computeWordDiff('The quick brown fox', 'The fast brown fox');
  assert.deepEqual(
    diff.map(d => [d.type, d.text]),
    [
      ['eq', 'The '],
      ['del', 'quick'],
      ['ins', 'fast'],
      ['eq', ' brown fox'],
    ],
  );

  const zh = computeWordDiff('这是一个测试', '这是一个好测试');
  assert.ok(zh.some(d => d.type === 'ins' && d.text === '好'));
});

test('countWords and splitParagraphs handle English, CJK, and line breaks', () => {
  assert.equal(countWords('Hello world'), 2);
  assert.equal(countWords('这是一个测试'), 6);
  assert.equal(countWords('Hello 世界'), 3);
  assert.equal(countWords('   '), 0);
  assert.deepEqual(splitParagraphs('\nFirst\n\nSecond\n'), ['First', '', 'Second']);
});
