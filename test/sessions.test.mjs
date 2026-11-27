import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  createSessionStore,
  compactHistory,
  isValidSessionId,
} from '../lib/sessions.mjs';

function makeTempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'office-ai-sess-'));
}

test('createSessionStore saves, lists by host, restores, deletes, and prunes old sessions', () => {
  const dir = makeTempDir();
  try {
    const store = createSessionStore({ configDir: dir, maxSessions: 3 });

    const s1 = store.save({
      host: 'excel',
      messages: [
        { role: 'user', content: 'Audit the Q3 financial model' },
        { role: 'assistant', content: 'Found 2 formula errors.' },
      ],
    });
    assert.ok(isValidSessionId(s1.id));
    assert.equal(s1.title, 'Audit the Q3 financial model');

    const s2 = store.save({
      host: 'word',
      messages: [{ role: 'user', content: 'Tighten intro paragraph' }],
    });

    // Update existing session s1 in place
    const s1Updated = store.save({
      id: s1.id,
      host: 'excel',
      messages: [
        { role: 'user', content: 'Audit the Q3 financial model' },
        { role: 'assistant', content: 'Found 2 formula errors.' },
        { role: 'user', content: 'Fix C4' },
      ],
    });
    assert.equal(s1Updated.id, s1.id);
    assert.equal(s1Updated.messageCount, 3);

    assert.equal(store.list({ host: 'excel' }).length, 1);
    assert.equal(store.list({ host: 'word' }).length, 1);
    assert.equal(store.get(s1.id).messages.length, 3);

    // Path traversal IDs are rejected
    assert.equal(store.get('../config'), null);
    assert.equal(store.remove('../config'), false);

    // Adding 2 more sessions triggers pruning to maxSessions=3
    store.save({ host: 'powerpoint', messages: [{ role: 'user', content: 'Slide deck 1' }] });
    store.save({ host: 'powerpoint', messages: [{ role: 'user', content: 'Slide deck 2' }] });
    assert.equal(store.list().length, 3);

    assert.equal(store.remove(s2.id) || store.remove(s1.id), true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('compactHistory preserves recent turns verbatim and summarizes older turns instead of dropping blindly', () => {
  const msgs = [];
  for (let i = 1; i <= 30; i++) {
    msgs.push({
      role: i % 2 === 1 ? 'user' : 'assistant',
      content: i % 2 === 1
        ? `User request number ${i} about revenue forecast`
        : `Assistant reply ${i} <edit action="set_formula" cell="B${i}">=SUM(A1:A${i})</edit>`,
    });
  }

  const res = compactHistory(msgs, { maxMessages: 10, maxTotalChars: 20000 });
  assert.equal(res.compacted, true);
  assert.equal(res.droppedCount, 20);
  assert.equal(res.messages.length, 11); // 1 rolling summary + 10 recent messages
  assert.equal(res.messages[0].role, 'system');
  assert.match(res.messages[0].content, /Earlier conversation summary \(20 earlier messages/);
  assert.match(res.messages[0].content, /\[proposed set_formula\]/);
  assert.equal(res.messages[10].content, msgs[29].content);

  // Short conversations are NOT compacted
  const shortRes = compactHistory(msgs.slice(0, 4), { maxMessages: 10 });
  assert.equal(shortRes.compacted, false);
  assert.equal(shortRes.messages.length, 4);
});
