import test from 'node:test';
import assert from 'node:assert/strict';
import { parseEnvContent, upsertEnvKey } from '../lib/env.mjs';

test('parseEnvContent parses comments, quotes, and values', () => {
  const raw = [
    '# Local API keys',
    'OPENAI_API_KEY=sk-test-123',
    'ANTHROPIC_API_KEY="sk-ant-456"',
    "GEMINI_API_KEY='AIza-test'",
    '',
  ].join('\n');

  assert.deepEqual(parseEnvContent(raw), {
    OPENAI_API_KEY: 'sk-test-123',
    ANTHROPIC_API_KEY: 'sk-ant-456',
    GEMINI_API_KEY: 'AIza-test',
  });
});

test('upsertEnvKey inserts, updates, and removes keys while preserving comments', () => {
  const initial = '# My keys\nOPENAI_API_KEY=old-val\n';
  const updated = upsertEnvKey(initial, 'OPENAI_API_KEY', 'new-val');
  assert.equal(updated, '# My keys\nOPENAI_API_KEY=new-val\n');

  const added = upsertEnvKey(updated, 'DEEPSEEK_API_KEY', 'ds-val');
  assert.deepEqual(parseEnvContent(added), {
    OPENAI_API_KEY: 'new-val',
    DEEPSEEK_API_KEY: 'ds-val',
  });

  const removed = upsertEnvKey(added, 'OPENAI_API_KEY', '');
  assert.deepEqual(parseEnvContent(removed), {
    DEEPSEEK_API_KEY: 'ds-val',
  });
  assert.match(removed, /^# My keys/m);
});
