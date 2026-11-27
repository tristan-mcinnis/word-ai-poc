/* Shared pure helpers for parsing assistant replies, rendering safe Markdown
   (including tables & clickable Excel/PowerPoint citations), and computing
   word-level diffs. Runs identically in browser ES modules and Node tests. */

export const VALID_ACTIONS = new Set([
  // Word actions
  'replace_selection',
  'insert_after',
  'insert_before',
  'replace_match',
  'append_doc',
  'replace_doc',
  'insert_docx',
  'comment',
  'reply_comment',
  // Excel actions
  'set_cells',
  'set_formula',
  'format_cells',
  'add_sheet',
  'add_chart',
  'sort_range',
  'excel_comment',
  'clear_range',
  // PowerPoint actions
  'update_shape',
  'replace_slide_text',
  'add_slide',
  'add_textbox',
  'add_table',
  'delete_slide',
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

function parseEditInner(rawInner, action, streaming = false) {
  let trimmed = rawInner.replace(/^\r?\n/, '').replace(/\r?\n$/, '');
  const findMatch = /<find>([\s\S]*?)<\/find>/i.exec(trimmed);
  const contentMatch = /<content>([\s\S]*?)<\/content>/i.exec(trimmed);

  let find = '';
  let content = trimmed;

  if (findMatch) {
    find = findMatch[1].replace(/^\r?\n/, '').replace(/\r?\n$/, '');
    if (contentMatch) {
      content = contentMatch[1].replace(/^\r?\n/, '').replace(/\r?\n$/, '');
    } else {
      content = trimmed
        .replace(/<find>[\s\S]*?<\/find>/i, '')
        .replace(/^\s*<content>\r?\n?/i, '')
        .replace(/<\/?(?:content|find)?[^>]*$/i, '')
        .replace(/^\r?\n+/, '')
        .replace(/\r?\n$/, '');
    }
  } else if (contentMatch) {
    content = contentMatch[1].replace(/^\r?\n/, '').replace(/\r?\n$/, '');
  } else if (streaming) {
    // Prevent raw <find> or <content> tags from flashing while mid-stream
    if (/^\s*<find>/i.test(trimmed)) {
      find = trimmed.replace(/^\s*<find>\r?\n?/i, '').replace(/<\/?f(?:i(?:n(?:d)?)?)?$/i, '');
      content = '';
    } else {
      content = trimmed
        .replace(/^\s*<content>\r?\n?/i, '')
        .replace(/<\/?(?:edit|replacement|find|content)?[^>]*$/i, '');
    }
  }

  return { action, find, content };
}

/**
 * Parse an assistant reply into ordered parts:
 * - { kind: 'prose', text }
 * - { kind: 'edit', action, label, commentId, attrs, find, content, streaming }
 *
 * Supports <edit action="..." ...>...</edit>, self-closing <edit ... />, and legacy
 * <replacement>...</replacement>, handling still-streaming unclosed tags gracefully
 * so raw XML never flashes in the UI.
 */
export function parseAssistantReply(text) {
  if (!text) return [];
  // Strip any leaked provider-internal DSML tool-call blocks (e.g. DeepSeek <|DSML|...>)
  const cleaned = String(text).replace(/<\s*\|\s*\|?\s*DSML[\s\S]*?<\/\s*\|\s*\|?\s*DSML[^>]*>/gi, '');
  const parts = [];
  const tagRe = /<(edit(?:\s+[^>]*?)?)\s*\/>|<(edit(?:\s+[^>]*)?|replacement)>([\s\S]*?)(?:<\/(edit|replacement)>|$)/gi;
  let last = 0;
  let match;

  while ((match = tagRe.exec(cleaned)) !== null) {
    if (match.index > last) {
      const prose = cleaned.slice(last, match.index);
      if (prose.trim()) parts.push({ kind: 'prose', text: prose });
    }

    const isSelfClosing = Boolean(match[1]);
    const openTag = isSelfClosing ? match[1] : match[2];
    const body = isSelfClosing ? '' : match[3];
    const closed = isSelfClosing || Boolean(match[4]);
    const isReplacement = openTag.toLowerCase() === 'replacement';

    let action = 'replace_selection';
    let label = '';
    let commentId = '';
    let attrs = {};

    if (!isReplacement) {
      attrs = parseAttributes(openTag.slice(4));
      if (attrs.action && VALID_ACTIONS.has(attrs.action.toLowerCase())) {
        action = attrs.action.toLowerCase();
      }
      if (attrs.label) label = attrs.label.trim();
      if (attrs.comment_id) commentId = attrs.comment_id.trim();
    }

    const parsed = parseEditInner(body, action, !closed);
    parts.push({
      kind: 'edit',
      action: parsed.action,
      label,
      commentId,
      attrs,
      find: parsed.find,
      content: parsed.content,
      streaming: !closed,
    });

    last = tagRe.lastIndex;
  }

  if (last < cleaned.length) {
    let tail = cleaned.slice(last);
    // Hide an incomplete opening or closing tag like "<ed", "<edit action=", or "</ed" while streaming
    tail = tail.replace(/<(?:edit(?:\s+[^>]*)?|replacement|\/(?:edit|replacement)?)?$/i, '');
    if (tail.trim()) parts.push({ kind: 'prose', text: tail });
  }

  return parts;
}

export function actionLabel(action, customLabel = '', attrs = {}) {
  if (customLabel) return customLabel;
  const rangeOrCell = attrs.range || attrs.cell || '';
  const sheetPrefix = attrs.sheet ? `${attrs.sheet}!` : '';
  const targetRef = rangeOrCell ? (rangeOrCell.includes('!') ? rangeOrCell : `${sheetPrefix}${rangeOrCell}`) : '';

  switch (action) {
    // Word
    case 'replace_selection': return 'Replace selection';
    case 'insert_after': return 'Insert below';
    case 'insert_before': return 'Insert above';
    case 'replace_match': return 'Edit in document';
    case 'append_doc': return 'Add to end of document';
    case 'replace_doc': return 'Replace entire document';
    case 'insert_docx': return 'Insert rendered document (.docx)';
    case 'comment': return 'Add Word comment';
    case 'reply_comment': return 'Reply to comment';
    // Excel
    case 'set_cells': return targetRef ? `Write ${targetRef}` : 'Write cells';
    case 'set_formula': return targetRef ? `Formula ${targetRef}` : 'Set formula';
    case 'format_cells': return targetRef ? `Format ${targetRef}` : 'Format cells';
    case 'add_sheet': return attrs.name ? `New sheet · ${attrs.name}` : 'Add worksheet';
    case 'add_chart': return attrs.title ? `Chart · ${attrs.title}` : `Add ${attrs.chart_type || 'chart'}`;
    case 'sort_range': return targetRef ? `Sort ${targetRef}` : 'Sort range';
    case 'excel_comment': return targetRef ? `Comment on ${targetRef}` : 'Add cell comment';
    case 'clear_range': return targetRef ? `Clear ${targetRef}` : 'Clear range';
    // PowerPoint
    case 'update_shape': {
      const parts = ['Update'];
      if (attrs.slide) parts.push(`Slide ${attrs.slide}`);
      if (attrs.shape) parts.push(attrs.shape);
      return parts.length > 1 ? parts.join(' · ') : 'Update slide text';
    }
    case 'replace_slide_text': return attrs.slide ? `Edit Slide ${attrs.slide}` : 'Edit slide text';
    case 'add_slide': return attrs.layout ? `Add slide · ${attrs.layout}` : 'Add new slide';
    case 'add_textbox': return attrs.slide ? `Add text box · Slide ${attrs.slide}` : 'Add text box';
    case 'add_table': return attrs.slide ? `Add table · Slide ${attrs.slide}` : 'Add slide table';
    case 'delete_slide': return attrs.slide ? `Delete Slide ${attrs.slide}` : 'Delete slide';
    default: return 'Suggested edit';
  }
}

export function primaryButtonLabel(action) {
  switch (action) {
    // Word
    case 'replace_selection': return 'Allow replace';
    case 'insert_after': return 'Insert below';
    case 'insert_before': return 'Insert above';
    case 'replace_match': return 'Apply edit';
    case 'append_doc': return 'Append';
    case 'replace_doc': return 'Replace document';
    case 'insert_docx': return 'Open rendered .docx';
    case 'comment': return 'Add comment';
    case 'reply_comment': return 'Post reply';
    // Excel
    case 'set_cells': return 'Apply to sheet';
    case 'set_formula': return 'Set formula';
    case 'format_cells': return 'Apply format';
    case 'add_sheet': return 'Create sheet';
    case 'add_chart': return 'Create chart';
    case 'sort_range': return 'Sort range';
    case 'excel_comment': return 'Add comment';
    case 'clear_range': return 'Clear cells';
    // PowerPoint
    case 'update_shape': return 'Apply to slide';
    case 'replace_slide_text': return 'Apply edit';
    case 'add_slide': return 'Insert slide';
    case 'add_textbox': return 'Add text box';
    case 'add_table': return 'Insert table';
    case 'delete_slide': return 'Delete slide';
    default: return 'Apply';
  }
}

export function isDestructiveAction(action) {
  return action === 'delete_slide' || action === 'clear_range' || action === 'replace_doc';
}

/* ---------- Lightweight XSS-safe Markdown renderer (with tables & host citations) ---------- */

const EXCEL_CELL_REF_RE = /^(?:(?:'[^']{1,60}'|[A-Za-z_][A-Za-z0-9_ .&-]{0,30})!)?\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?$/;
const EXCEL_NON_CELL_TOKENS = /^(?:Q[1-4]|H[12]|FY\d{2,4}|CY\d{2,4}|P[1-4]|V[1-9]\d?)$/;
const PPT_SLIDE_REF_RE = /^Slide\s+(\d{1,4})$/i;

function isLikelyExcelCitation(raw) {
  const clean = String(raw || '').trim();
  if (!EXCEL_CELL_REF_RE.test(clean)) return false;
  // Explicit sheet prefix, range colon, or $ anchor is unambiguously a cell/range reference
  if (clean.includes('!') || clean.includes(':') || clean.includes('$')) return true;
  return !EXCEL_NON_CELL_TOKENS.test(clean);
}

function makeRefChip(refText) {
  const safe = escapeHtml(refText);
  return `<button type="button" class="ref-chip" data-ref="${safe}" title="Jump to ${safe}">${safe}</button>`;
}

function renderInline(text, host = '') {
  let out = escapeHtml(text);
  const tokens = [];
  const stash = html => {
    const idx = tokens.push(html) - 1;
    return `\u0000TOK${idx}\u0000`;
  };

  // Inline code first so markdown inside backticks is preserved
  out = out.replace(/`([^`\n]+)`/g, (_, code) => {
    const raw = code.replace(/&amp;/g, '&').replace(/&#39;/g, "'").replace(/&quot;/g, '"');
    if (host === 'excel' && isLikelyExcelCitation(raw)) {
      return stash(makeRefChip(raw.trim()));
    }
    if (host === 'powerpoint' && PPT_SLIDE_REF_RE.test(raw.trim())) {
      return stash(makeRefChip(raw.trim()));
    }
    return stash(`<code>${code}</code>`);
  });

  // Bracketed citations like [Sheet1!B4], [B2:D10], or [Slide 3]
  if (host === 'excel') {
    out = out.replace(/\[([^\[\]\n]{2,50})\]/g, (full, inner) => {
      const raw = inner.replace(/&amp;/g, '&').replace(/&#39;/g, "'").trim();
      if (isLikelyExcelCitation(raw)) return stash(makeRefChip(raw));
      return full;
    });
  } else if (host === 'powerpoint') {
    out = out.replace(/\[?(Slide\s+\d{1,4})\]?/g, (full, inner) => {
      if (PPT_SLIDE_REF_RE.test(inner.trim())) return stash(makeRefChip(inner.trim()));
      return full;
    });
  }

  out = out.replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>');
  out = out.replace(/__([^_\n]+)__/g, '<strong>$1</strong>');
  out = out.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1<em>$2</em>');
  out = out.replace(/~~([^~\n]+)~~/g, '<del>$1</del>');
  out = out.replace(/\u0000TOK(\d+)\u0000/g, (_, idx) => tokens[Number(idx)]);
  return out;
}

function isTableSeparator(line) {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  if (!trimmed.includes('-')) return false;
  const cols = trimmed.split('|').map(c => c.trim());
  return cols.length >= 1 && cols.every(c => /^:?-{2,}:?$/.test(c));
}

function splitTableRow(line) {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  return trimmed.split('|').map(c => c.trim());
}

export function renderMarkdown(md, options = {}) {
  if (!md || !md.trim()) return '';
  const host = options.host || '';
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

    // Markdown pipe table (header + separator + 0..N rows)
    if (line.includes('|') && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
      const headers = splitTableRow(line);
      i += 2;
      const bodyRows = [];
      while (i < lines.length && lines[i].trim() && lines[i].includes('|')) {
        bodyRows.push(splitTableRow(lines[i]));
        i++;
      }
      const thead = `<thead><tr>${headers.map(h => `<th>${renderInline(h, host)}</th>`).join('')}</tr></thead>`;
      const tbody = bodyRows.length
        ? `<tbody>${bodyRows.map(r => `<tr>${headers.map((_, ci) => `<td>${renderInline(r[ci] ?? '', host)}</td>`).join('')}</tr>`).join('')}</tbody>`
        : '';
      html.push(`<div class="md-table-wrap"><table class="md-table">${thead}${tbody}</table></div>`);
      continue;
    }

    // Headings
    const heading = /^(#{1,4})\s+(.+)$/.exec(line.trim());
    if (heading) {
      const level = Math.min(heading[1].length + 1, 5); // h2..h5 inside pane
      html.push(`<h${level}>${renderInline(heading[2].trim(), host)}</h${level}>`);
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
      html.push(`<blockquote>${quoteLines.map(l => renderInline(l, host)).join('<br>')}</blockquote>`);
      continue;
    }

    // Unordered list
    if (/^[-*•]\s+/.test(line.trim())) {
      const items = [];
      while (i < lines.length && /^[-*•]\s+/.test(lines[i].trim())) {
        items.push(`<li>${renderInline(lines[i].trim().replace(/^[-*•]\s+/, ''), host)}</li>`);
        i++;
      }
      html.push(`<ul>${items.join('')}</ul>`);
      continue;
    }

    // Ordered list
    if (/^\d+[.)]\s+/.test(line.trim())) {
      const items = [];
      while (i < lines.length && /^\d+[.)]\s+/.test(lines[i].trim())) {
        items.push(`<li>${renderInline(lines[i].trim().replace(/^\d+[.)]\s+/, ''), host)}</li>`);
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
      !/^(?:-{3,}|\*{3,}|_{3,})$/.test(lines[i].trim()) &&
      !(lines[i].includes('|') && i + 1 < lines.length && isTableSeparator(lines[i + 1]))
    ) {
      para.push(renderInline(lines[i], host));
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

export function splitParagraphs(text) {
  return String(text).replace(/\r\n?/g, '\n').split('\n').filter((line, idx, arr) => {
    if (idx === 0 || idx === arr.length - 1) return line.trim().length > 0;
    return true;
  });
}
