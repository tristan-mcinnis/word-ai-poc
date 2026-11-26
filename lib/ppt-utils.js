/* Pure PowerPoint helpers for deck serialization, Slide Master layout matching,
   slide specification parsing, and placeholder classification. */

/**
 * Parse a slide specification from an <edit action="add_slide"> body.
 * Supports JSON ({ title, subtitle, bullets, body, columns }) or clean structured lines:
 *   Title: ...
 *   Subtitle: ...
 *   - Bullet 1
 *   - Bullet 2
 */
export function parseSlideSpec(raw) {
  const text = String(raw || '').trim();
  if (!text) return { title: '', subtitle: '', bullets: [], body: '', columns: [] };

  const unwrapped = text.replace(/^```(?:json)?\s*\n?/i, '').replace(/\n?```\s*$/, '').trim();

  if (unwrapped.startsWith('{')) {
    try {
      const j = JSON.parse(unwrapped);
      if (j && typeof j === 'object') {
        const bullets = Array.isArray(j.bullets) ? j.bullets.map(b => String(b).trim()).filter(Boolean) : [];
        const columns = Array.isArray(j.columns)
          ? j.columns.map(col => Array.isArray(col) ? col.map(String).join('\n') : String(col || '').trim()).filter(Boolean)
          : [];
        return {
          title: typeof j.title === 'string' ? j.title.trim() : '',
          subtitle: typeof j.subtitle === 'string' ? j.subtitle.trim() : '',
          bullets,
          body: typeof j.body === 'string' ? j.body.trim() : bullets.join('\n'),
          columns,
        };
      }
    } catch {
      // Fall through to structured text parsing
    }
  }

  const lines = unwrapped.split(/\r?\n/);
  let title = '';
  let subtitle = '';
  const bullets = [];
  const bodyLines = [];
  let inBody = false;

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      if (inBody && bodyLines.length) bodyLines.push('');
      continue;
    }

    const titleMatch = /^(?:#{1,2}\s+|title:\s*)(.+)$/i.exec(line);
    if (!title && titleMatch) {
      title = titleMatch[1].trim();
      continue;
    }

    const subMatch = /^subtitle:\s*(.+)$/i.exec(line);
    if (!subtitle && subMatch) {
      subtitle = subMatch[1].trim();
      continue;
    }

    const bodyLabel = /^(?:body|content|bullets?):\s*(.*)$/i.exec(line);
    if (bodyLabel) {
      inBody = true;
      if (bodyLabel[1].trim()) {
        bodyLines.push(bodyLabel[1].trim());
      }
      continue;
    }

    const bulletMatch = /^[-*•]\s+(.+)$/.exec(line);
    if (bulletMatch) {
      inBody = true;
      const item = bulletMatch[1].trim();
      bullets.push(item);
      bodyLines.push(item);
      continue;
    }

    inBody = true;
    bodyLines.push(line);
  }

  // If no explicit Title: prefix was used, treat the first non-bullet line as title when there are multiple lines
  if (!title && bodyLines.length >= 2 && bullets.length === 0) {
    title = bodyLines.shift();
  }

  while (bodyLines.length && !bodyLines[bodyLines.length - 1]) bodyLines.pop();

  return {
    title,
    subtitle,
    bullets,
    body: bodyLines.join('\n').trim(),
    columns: [],
  };
}

/**
 * Match a requested layout name (e.g. "Title and Content", "Two Content", "Blank")
 * against the presentation's actual Slide Master layouts:
 * layouts: Array<{ slideMasterId, slideMasterName, layoutId, layoutName }>
 */
export function matchSlideLayout(layouts, requestedName = '') {
  if (!Array.isArray(layouts) || layouts.length === 0) return null;
  const req = String(requestedName || '').trim().toLowerCase();

  if (req) {
    // 1. Exact case-insensitive match
    const exact = layouts.find(l => (l.layoutName || '').trim().toLowerCase() === req);
    if (exact) return exact;

    // 2. Normalized alphanumeric match
    const normReq = req.replace(/[^a-z0-9]+/g, '');
    const normMatch = layouts.find(l => (l.layoutName || '').toLowerCase().replace(/[^a-z0-9]+/g, '') === normReq);
    if (normMatch) return normMatch;

    // 3. Substring match
    const subMatch = layouts.find(l => (l.layoutName || '').toLowerCase().includes(req));
    if (subMatch) return subMatch;

    // 4. Token overlap match (e.g. "two column" -> "Two Content")
    const reqTokens = req.split(/[^a-z0-9]+/).filter(Boolean);
    let best = null;
    let bestScore = 0;
    for (const l of layouts) {
      const lTokens = (l.layoutName || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
      const score = reqTokens.filter(t => lTokens.includes(t)).length;
      if (score > bestScore) {
        bestScore = score;
        best = l;
      }
    }
    if (best && bestScore > 0) return best;
  }

  // Default preference: "Title and Content", then any layout containing "content", then index 1 or 0
  return (
    layouts.find(l => /title\s+and\s+content/i.test(l.layoutName || '')) ||
    layouts.find(l => /content/i.test(l.layoutName || '')) ||
    layouts[1] ||
    layouts[0]
  );
}

/**
 * Given shapes on a newly created slide (with { name, type, top, left }),
 * classify which shape should receive the title, subtitle, and body content(s).
 */
export function classifyPlaceholders(shapes) {
  if (!Array.isArray(shapes) || shapes.length === 0) {
    return { titleShape: null, subtitleShape: null, bodyShapes: [] };
  }

  let titleShape = null;
  let subtitleShape = null;
  const bodyShapes = [];

  // Sort top-to-bottom, left-to-right for deterministic fallback
  const sorted = shapes.slice().sort((a, b) => {
    const dt = (a.top ?? 0) - (b.top ?? 0);
    if (Math.abs(dt) > 12) return dt;
    return (a.left ?? 0) - (b.left ?? 0);
  });

  for (const s of sorted) {
    const name = String(s.name || '').toLowerCase();
    if (/footer|slide\s*number|date/i.test(name)) continue;

    if (!subtitleShape && /subtitle/i.test(name)) {
      subtitleShape = s;
    } else if (!titleShape && /title/i.test(name)) {
      titleShape = s;
    } else if (/content|body|text|object|placeholder/i.test(name)) {
      bodyShapes.push(s);
    }
  }

  // Positional fallback when names are localized or generic
  if (!titleShape && !subtitleShape && bodyShapes.length === 0) {
    const textCandidates = sorted.filter(s => !/footer|slide\s*number|date/i.test(String(s.name || '')));
    if (textCandidates.length >= 1) titleShape = textCandidates[0];
    if (textCandidates.length >= 2) bodyShapes.push(...textCandidates.slice(1));
  } else if (!titleShape && bodyShapes.length >= 2) {
    titleShape = bodyShapes.shift();
  }

  return { titleShape, subtitleShape, bodyShapes };
}

/**
 * Format the presentation slides and layouts into clean, compact context text for the AI.
 */
export function formatPresentationSnapshot({ slides = [], layouts = [], selectedSlideIndices = [], maxChars = 36000 }) {
  const parts = [];
  const selectedSet = new Set(selectedSlideIndices);

  if (layouts.length > 0) {
    const uniqueNames = [...new Set(layouts.map(l => l.layoutName).filter(Boolean))];
    if (uniqueNames.length) {
      parts.push(`Available Slide Master layouts: ${uniqueNames.join(', ')}`);
    }
  }

  if (!slides.length) {
    parts.push('Presentation currently has no slides.');
    return parts.join('\n');
  }

  for (const slide of slides) {
    const marker = selectedSet.has(slide.index) ? ' [SELECTED]' : '';
    const layoutTag = slide.layoutName ? ` (${slide.layoutName})` : '';
    const header = `--- Slide ${slide.index}${layoutTag}${marker} ---`;
    const shapeLines = [];

    for (const shape of slide.shapes || []) {
      const label = shape.name || shape.type || shape.id;
      const text = String(shape.text || '').trim();
      if (text) {
        const compact = text.replace(/\r\n?/g, '\n');
        shapeLines.push(`[${label}] ${compact}`);
      } else if (shape.type && /table|chart|image|picture/i.test(shape.type)) {
        shapeLines.push(`[${label} (${shape.type})]`);
      }
    }

    const block = shapeLines.length ? `${header}\n${shapeLines.join('\n')}` : `${header}\n(empty slide)`;
    if (parts.join('\n\n').length + block.length > maxChars) {
      parts.push(`... (${slides.length - slide.index + 1} more slides truncated)`);
      break;
    }
    parts.push(block);
  }

  return parts.join('\n\n');
}

/**
 * Parse a PowerPoint slide reference like "Slide 3" or "3" -> 1-based slide number (or null).
 */
export function parseSlideRef(raw) {
  const m = /^(?:slide\s*#?\s*)?(\d+)$/i.exec(String(raw || '').trim());
  if (!m) return null;
  const num = parseInt(m[1], 10);
  return num >= 1 ? num : null;
}
