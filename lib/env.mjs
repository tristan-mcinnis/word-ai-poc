/* Pure helpers for reading and updating ~/.config/word-poc/.env safely. */

export function parseEnvContent(raw) {
  const out = {};
  if (!raw) return out;
  for (const line of String(raw).split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq <= 0) continue;
    const k = trimmed.slice(0, eq).trim();
    let v = trimmed.slice(eq + 1).trim();
    if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
      v = v.slice(1, -1);
    }
    if (k && v) out[k] = v;
  }
  return out;
}

/**
 * Insert, update, or remove (when value is empty) a single KEY=VALUE entry
 * while preserving unrelated lines and comments.
 */
export function upsertEnvKey(raw, key, value) {
  const cleanKey = String(key || '').trim();
  if (!/^[A-Z0-9_]+$/.test(cleanKey)) {
    throw new Error('Invalid environment key name');
  }
  const cleanVal = String(value || '').replace(/[\r\n]/g, '').trim();
  const lines = raw ? String(raw).split(/\r?\n/) : [];
  const filtered = lines.filter(line => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) return true;
    const eq = trimmed.indexOf('=');
    if (eq <= 0) return true;
    return trimmed.slice(0, eq).trim() !== cleanKey;
  });

  // Trim trailing blank lines
  while (filtered.length && !filtered[filtered.length - 1].trim()) filtered.pop();

  if (cleanVal) {
    const needsQuotes = /[\s#"'`]/.test(cleanVal);
    const serialized = needsQuotes ? JSON.stringify(cleanVal) : cleanVal;
    filtered.push(`${cleanKey}=${serialized}`);
  }

  return filtered.length ? filtered.join('\n') + '\n' : '';
}
