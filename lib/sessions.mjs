/* Persistent local chat session storage (~/.config/word-poc/sessions/)
   and deterministic rolling context compaction for long conversations. */

import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';

const SAFE_ID_RE = /^[a-zA-Z0-9_-]{4,64}$/;
const MAX_SESSIONS = 50;
const MAX_STORED_MESSAGES = 100;
const MAX_MESSAGE_CHARS = 24000;

export function isValidSessionId(id) {
  return typeof id === 'string' && SAFE_ID_RE.test(id);
}

export function generateSessionId() {
  return `s_${Date.now().toString(36)}_${crypto.randomBytes(4).toString('hex')}`;
}

export function sanitizeMessages(raw) {
  if (!Array.isArray(raw)) return [];
  return raw
    .filter(m => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string')
    .map(m => ({ role: m.role, content: m.content.slice(0, MAX_MESSAGE_CHARS) }))
    .slice(-MAX_STORED_MESSAGES);
}

export function deriveSessionTitle(messages, fallback = 'Untitled chat') {
  const firstUser = (messages || []).find(m => m.role === 'user' && m.content.trim());
  if (!firstUser) return fallback;
  const clean = firstUser.content.replace(/\s+/g, ' ').trim();
  return clean.length > 64 ? `${clean.slice(0, 63)}…` : clean;
}

/**
 * Compact a conversation history so recent turns stay verbatim while older turns
 * beyond maxMessages or maxTotalChars are folded into a compact rolling summary.
 */
export function compactHistory(rawMessages, { maxMessages = 20, maxTotalChars = 48000 } = {}) {
  const cleaned = sanitizeMessages(rawMessages);
  if (!cleaned.length) return { messages: [], compacted: false, droppedCount: 0 };

  // Keep recent messages within both message count and char budget
  const recent = [];
  let charBudget = maxTotalChars;
  for (let i = cleaned.length - 1; i >= 0; i--) {
    const msg = cleaned[i];
    if (recent.length >= maxMessages) break;
    if (recent.length >= 2 && msg.content.length > charBudget) break;
    recent.unshift(msg);
    charBudget -= msg.content.length;
  }

  const droppedCount = cleaned.length - recent.length;
  if (droppedCount <= 0) {
    return { messages: recent, compacted: false, droppedCount: 0 };
  }

  const older = cleaned.slice(0, droppedCount);
  const bullets = [];
  let summaryChars = 0;
  for (const m of older) {
    const roleLabel = m.role === 'user' ? 'User' : 'Assistant';
    // Strip verbose <edit> XML bodies from summary while noting edit actions taken
    const condensed = m.content
      .replace(/<edit\b([^>]*)>[\s\S]*?<\/edit>/gi, (_, attrs) => {
        const act = /action=["']([^"']+)["']/i.exec(attrs || '');
        return `[proposed ${act ? act[1] : 'edit'}]`;
      })
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 180);
    if (!condensed) continue;
    const line = `- ${roleLabel}: ${condensed}`;
    if (summaryChars + line.length > 3200) break;
    bullets.push(line);
    summaryChars += line.length;
  }

  const summaryMsg = {
    role: 'system',
    content: `Earlier conversation summary (${droppedCount} earlier messages compacted to save context):\n${bullets.join('\n')}`,
  };

  return {
    messages: [summaryMsg, ...recent],
    compacted: true,
    droppedCount,
  };
}

export function createSessionStore({
  configDir = path.join(os.homedir(), '.config', 'word-poc'),
  maxSessions = MAX_SESSIONS,
} = {}) {
  const sessionsDir = path.join(configDir, 'sessions');

  function ensureDir() {
    fs.mkdirSync(sessionsDir, { recursive: true, mode: 0o700 });
  }

  function fileFor(id) {
    if (!isValidSessionId(id)) throw new Error('Invalid session ID');
    return path.join(sessionsDir, `${id}.json`);
  }

  function list({ host = '' } = {}) {
    let files;
    try {
      files = fs.readdirSync(sessionsDir).filter(f => f.endsWith('.json'));
    } catch {
      return [];
    }

    const items = [];
    for (const f of files) {
      try {
        const raw = JSON.parse(fs.readFileSync(path.join(sessionsDir, f), 'utf8'));
        if (!raw || !isValidSessionId(raw.id)) continue;
        if (host && raw.host && raw.host !== host) continue;
        items.push({
          id: raw.id,
          host: raw.host || 'word',
          title: raw.title || 'Untitled chat',
          messageCount: Array.isArray(raw.messages) ? raw.messages.length : 0,
          createdAt: raw.createdAt || raw.updatedAt || 0,
          updatedAt: raw.updatedAt || raw.createdAt || 0,
        });
      } catch {
        // Ignore corrupt session files
      }
    }
    items.sort((a, b) => b.updatedAt - a.updatedAt);
    return items;
  }

  function get(id) {
    if (!isValidSessionId(id)) return null;
    try {
      const raw = JSON.parse(fs.readFileSync(fileFor(id), 'utf8'));
      if (!raw || !isValidSessionId(raw.id)) return null;
      return {
        id: raw.id,
        host: raw.host || 'word',
        title: raw.title || deriveSessionTitle(raw.messages),
        createdAt: raw.createdAt || raw.updatedAt || Date.now(),
        updatedAt: raw.updatedAt || Date.now(),
        messages: sanitizeMessages(raw.messages),
      };
    } catch {
      return null;
    }
  }

  function pruneOld() {
    const all = list();
    if (all.length <= maxSessions) return;
    for (const stale of all.slice(maxSessions)) {
      try { fs.unlinkSync(fileFor(stale.id)); } catch {}
    }
  }

  function save({ id, host = 'word', title, messages }) {
    ensureDir();
    const sessionId = isValidSessionId(id) ? id : generateSessionId();
    const existing = get(sessionId);
    const cleaned = sanitizeMessages(messages);
    const now = Date.now();
    const record = {
      id: sessionId,
      host: String(host || existing?.host || 'word').toLowerCase(),
      title: String(title || deriveSessionTitle(cleaned, existing?.title || 'Untitled chat')).slice(0, 120),
      createdAt: existing?.createdAt || now,
      updatedAt: now,
      messages: cleaned,
    };
    fs.writeFileSync(fileFor(sessionId), JSON.stringify(record, null, 2) + '\n', { mode: 0o600 });
    pruneOld();
    return {
      id: record.id,
      host: record.host,
      title: record.title,
      messageCount: record.messages.length,
      createdAt: record.createdAt,
      updatedAt: record.updatedAt,
    };
  }

  function remove(id) {
    if (!isValidSessionId(id)) return false;
    try {
      fs.unlinkSync(fileFor(id));
      return true;
    } catch {
      return false;
    }
  }

  return { list, get, save, remove };
}
