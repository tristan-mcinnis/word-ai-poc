/* Shared pure helpers for parsing assistant replies, rendering safe Markdown,
   and computing word-level diffs. Runs identically in browser ES modules and Node tests. */

const VALID_ACTIONS = new Set([
  'replace_selection',
  'insert_after',
  'insert_before',
  'replace_match',
  'append_doc',
  'comment',
  'reply_comment',
]);

export function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function parseAttributes(raw) {
  const attrs = {};
  const re = /([a-zA-Z0-9_-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(raw || '')) !== null) {
    attrs[m[1].toLowerCase()] = m[2] ?? m[3] ?? '';
  }
  return attrs;
}

function parseEditInner(rawInner, action) {
  const trimmed = rawInner.replace(/^\r?\n/, '').replace(/\r?\n$/, '');
  const findMatch = /<find>([\s\S]*?)<\/find>/i.exec(trimmed);
  const contentMatch = /<content>([\s\S]*?)<\/content>/i.exec(trimmed);

  let find = '';
  let content = trimmed;

  if (findMatch) {
    find = findMatch[1].replace(/^\r?\n/, '').replace(/\r?\n$/, '');
    if (contentMatch) {
      content = contentMatch[1].replace(/^\r?\n/, '').replace(/\r?\n$/, '');
    } else {
      content = trimmed.replace(/<find>[\s\S]*?<\/find>/i, '').replace(/^\r?\n+/, '').replace(/\r?\n$/, '');
    }
  } else if (contentMatch) {
    content = contentMatch[1].replace(/^\r?\n/, '').replace(/\r?\n$/, '');
  }

  return { action, find, content };
}

/**
 * Parse an assistant reply into ordered parts:
 * - { kind: 'prose', text }
 * - { kind: 'edit', action, label, find, content, streaming }
 *
 * Supports both <edit action="...">...</edit> and legacy <replacement>...</replacement>,
 * and handles still-streaming unclosed tags gracefully so raw XML never flashes in the UI.
 */
export function parseAssistantReply(text) {
  if (!text) return [];
  const parts = [];
  const tagRe = /<(edit(?:\s+[^>]*)?|replacement)>([\s\S]*?)(?:<\/(edit|replacement)>|$)/gi;
  let last = 0;
  let match;

  while ((match = tagRe.exec(text)) !== null) {
    if (match.index > last) {
      const prose = text.slice(last, match.index);
      if (prose.trim()) parts.push({ kind: 'prose', text: prose });
    }

    const openTag = match[1];
    const body = match[2];
    const closed = Boolean(match[3]);
    const isReplacement = openTag.toLowerCase() === 'replacement';

    let action = 'replace_selection';
    let label = '';
    let commentId = '';
    if (!isReplacement) {
      const attrs = parseAttributes(openTag.slice(4));
      if (attrs.action && VALID_ACTIONS.has(attrs.action.toLowerCase())) {
        action = attrs.action.toLowerCase();
      }
      if (attrs.label) label = attrs.label.trim();
      if (attrs.comment_id) commentId = attrs.comment_id.trim();
    }

    const parsed = parseEditInner(body, action);
    parts.push({
      kind: 'edit',
      action: parsed.action,
      label,
      commentId,
      find: parsed.find,
      content: parsed.content,
      streaming: !closed,
    });

    last = tagRe.lastIndex;
  }

  if (last < text.length) {
    let tail = text.slice(last);
    // Hide an incomplete opening tag like "<ed" or "<edit action=" while streaming
    tail = tail.replace(/<(?:edit(?:\s+[^>]*)?|replacement|\/(?:edit|replacement)?)?$/i, '');
    if (tail.trim()) parts.push({ kind: 'prose', text: tail });
  }

  return parts;
}

export function actionLabel(action, customLabel = '') {
  if (customLabel) return customLabel;
  switch (action) {
    case 'replace_selection': return 'Replace selection';
    case 'insert_after': return 'Insert below';
    case 'insert_before': return 'Insert above';
    case 'replace_match': return 'Edit in document';
    case 'append_doc': return 'Add to end of document';
    case 'comment': return 'Add Word comment';
    case 'reply_comment': return 'Reply to comment';
    default: return 'Suggested edit';
  }
}

export function primaryButtonLabel(action) {
  switch (action) {
    case 'replace_selection': return 'Allow replace';
    case 'insert_after': return 'Insert below';
    case 'insert_before': return 'Insert above';
    case 'replace_match': return 'Apply edit';
    case 'append_doc': return 'Append';
    case 'comment': return 'Add comment';
    case 'reply_comment': return 'Post reply';
    default: return 'Apply';
  }
}

/* ---------- Lightweight XSS-safe Markdown renderer ---------- */

function renderInline(text) {
  let out = escapeHtml(text);
  // Inline code first so markdown inside backticks is preserved
  const codeSpans = [];
  out = out.replace(/`([^`\n]+)`/g, (_, code) => {
    const idx = codeSpans.push(`<code>${code}</code>`) - 1;
    return `\u0000CODE${idx}\u0000`;
  });
  out = out.replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>');
  out = out.replace(/__([^_\n]+)__/g, '<strong>$1</strong>');
  out = out.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1<em>$2</em>');
  out = out.replace(/~~([^~\n]+)~~/g, '<del>$1</del>');
  out = out.replace(/\u0000CODE(\d+)\u0000/g, (_, idx) => codeSpans[Number(idx)]);
  return out;
}

export function renderMarkdown(md) {
  if (!md || !md.trim()) return '';
  const lines = String(md).replace(/\r\n?/g, '\n').split('\n');
  const html = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Blank line
    if (!line.trim()) {
      i++;
      continue;
    }

    // Fenced code block
    if (/^```/.test(line.trim())) {
      const codeLines = [];
      i++;
      while (i < lines.length && !/^```/.test(lines[i].trim())) {
        codeLines.push(lines[i]);
        i++;
      }
      if (i < lines.length) i++;
      html.push(`<pre><code>${escapeHtml(codeLines.join('\n'))}</code></pre>`);
      continue;
    }

    // Horizontal rule
    if (/^(?:-{3,}|\*{3,}|_{3,})$/.test(line.trim())) {
      html.push('<hr>');
      i++;
      continue;
    }

    // Headings
    const heading = /^(#{1,4})\s+(.+)$/.exec(line.trim());
    if (heading) {
      const level = Math.min(heading[1].length + 1, 5); // h2..h5 inside pane
      html.push(`<h${level}>${renderInline(heading[2].trim())}</h${level}>`);
      i++;
      continue;
    }

    // Blockquote
    if (/^>\s?/.test(line.trim())) {
      const quoteLines = [];
      while (i < lines.length && /^>\s?/.test(lines[i].trim())) {
        quoteLines.push(lines[i].trim().replace(/^>\s?/, ''));
        i++;
      }
      html.push(`<blockquote>${quoteLines.map(renderInline).join('<br>')}</blockquote>`);
      continue;
    }

    // Unordered list
    if (/^[-*•]\s+/.test(line.trim())) {
      const items = [];
      while (i < lines.length && /^[-*•]\s+/.test(lines[i].trim())) {
        items.push(`<li>${renderInline(lines[i].trim().replace(/^[-*•]\s+/, ''))}</li>`);
        i++;
      }
      html.push(`<ul>${items.join('')}</ul>`);
      continue;
    }

    // Ordered list
    if (/^\d+[.)]\s+/.test(line.trim())) {
      const items = [];
      while (i < lines.length && /^\d+[.)]\s+/.test(lines[i].trim())) {
        items.push(`<li>${renderInline(lines[i].trim().replace(/^\d+[.)]\s+/, ''))}</li>`);
        i++;
      }
      html.push(`<ol>${items.join('')}</ol>`);
      continue;
    }

    // Regular paragraph (group consecutive plain lines)
    const para = [];
    while (
      i < lines.length &&
      lines[i].trim() &&
      !/^```/.test(lines[i].trim()) &&
      !/^#{1,4}\s+/.test(lines[i].trim()) &&
      !/^>\s?/.test(lines[i].trim()) &&
      !/^[-*•]\s+/.test(lines[i].trim()) &&
      !/^\d+[.)]\s+/.test(lines[i].trim()) &&
      !/^(?:-{3,}|\*{3,}|_{3,})$/.test(lines[i].trim())
    ) {
      para.push(renderInline(lines[i]));
      i++;
    }
    html.push(`<p>${para.join('<br>')}</p>`);
  }

  return html.join('');
}

/* ---------- Word / CJK-aware inline diff ---------- */

function tokenize(text) {
  // Split into whitespace, CJK characters, or Latin/numeric words + punctuation
  return String(text).match(/\s+|[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]|[A-Za-z0-9_']+|\S/g) || [];
}

/**
 * Returns array of { type: 'eq' | 'del' | 'ins', text }.
 * Falls back cleanly for very large passages to stay fast in the taskpane.
 */
export function computeWordDiff(oldText, newText) {
  const a = tokenize(oldText || '');
  const b = tokenize(newText || '');
  if (!a.length && !b.length) return [];
  if (!a.length) return [{ type: 'ins', text: newText }];
  if (!b.length) return [{ type: 'del', text: oldText }];

  // Guard quadratic memory on huge blocks
  if (a.length * b.length > 90000) {
    return [
      { type: 'del', text: oldText },
      { type: 'ins', text: newText },
    ];
  }

  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Uint16Array(n + 1));

  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      if (a[i] === b[j]) dp[i][j] = dp[i + 1][j + 1] + 1;
      else dp[i][j] = dp[i + 1][j] >= dp[i][j + 1] ? dp[i + 1][j] : dp[i][j + 1];
    }
  }

  const raw = [];
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (a[i] === b[j]) {
      raw.push({ type: 'eq', text: a[i++] });
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      raw.push({ type: 'del', text: a[i++] });
    } else {
      raw.push({ type: 'ins', text: b[j++] });
    }
  }
  while (i < m) raw.push({ type: 'del', text: a[i++] });
  while (j < n) raw.push({ type: 'ins', text: b[j++] });

  // Merge adjacent runs of the same type
  const merged = [];
  for (const item of raw) {
    const prev = merged[merged.length - 1];
    if (prev && prev.type === item.type) prev.text += item.text;
    else merged.push({ ...item });
  }
  return merged;
}

export function renderDiffHtml(oldText, newText) {
  return computeWordDiff(oldText, newText)
    .map(part => {
      const safe = escapeHtml(part.text);
      if (part.type === 'del') return `<del class="diff-del">${safe}</del>`;
      if (part.type === 'ins') return `<ins class="diff-ins">${safe}</ins>`;
      return safe;
    })
    .join('');
}

export function countWords(text) {
  const clean = String(text || '').trim();
  if (!clean) return 0;
  const cjk = (clean.match(/[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g) || []).length;
  const nonCjk = clean
    .replace(/[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean).length;
  return cjk + nonCjk;
}
