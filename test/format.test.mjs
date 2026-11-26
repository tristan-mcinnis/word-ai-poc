import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseAssistantReply,
  renderMarkdown,
  computeWordDiff,
  countWords,
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
});

test('renderMarkdown formats headings, lists, bold, code and escapes HTML', () => {
  const md = [
    '## Summary',
    '',
    'This has **bold**, *italic*, `code`, and <script>alert(1)</script>.',
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
  assert.match(html, /&lt;script&gt;/);
  assert.match(html, /<ul><li>Point one<\/li><li>Point two<\/li><\/ul>/);
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

test('countWords counts both English words and CJK characters', () => {
  assert.equal(countWords('Hello world'), 2);
  assert.equal(countWords('这是一个测试'), 6);
  assert.equal(countWords('Hello 世界'), 3);
  assert.equal(countWords('   '), 0);
});
