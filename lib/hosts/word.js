/* Microsoft Word host adapter (Office.js WordApi). */

import { countWords, splitParagraphs } from '../format.js';

const STANDARD_STYLES = [
  'Title',
  'Subtitle',
  'Heading 1',
  'Heading 2',
  'Heading 3',
  'Heading 4',
  'Normal',
  'Quote',
  'Intense Quote',
  'List Paragraph',
];

/**
 * Parse a Word edit body into structured paragraph blocks with optional Word
 * style names and native list types (bullet / number).
 */
export function parseStyledWordBlocks(content, defaultStyle = '') {
  const lines = splitParagraphs(content);
  const fallback = String(defaultStyle || '').trim();
  return lines.map((rawLine, idx) => {
    const trimmed = rawLine.trim();

    // Explicit [Style Name] prefix, e.g. "[Title] The Role of LLMs"
    const bracket = /^\[([A-Za-z0-9 _-]{2,40})\]\s+(.+)$/.exec(trimmed);
    if (bracket) {
      return { text: bracket[2].trim(), style: bracket[1].trim(), listType: null };
    }

    // Markdown headings -> Word Heading 1-4 (or explicit Title if defaultStyle="Title" on first line)
    const heading = /^(#{1,4})\s+(.+)$/.exec(trimmed);
    if (heading) {
      const level = heading[1].length;
      const style = idx === 0 && fallback && level === 1 ? fallback : `Heading ${level}`;
      return { text: heading[2].trim(), style, listType: null };
    }

    // Native Word bullet list item
    const bullet = /^[-*•]\s+(.+)$/.exec(trimmed);
    if (bullet) {
      return { text: bullet[1].trim(), style: null, listType: 'bullet' };
    }

    // Native Word numbered list item
    const numbered = /^\d+[.)]\s+(.+)$/.exec(trimmed);
    if (numbered) {
      return { text: numbered[1].trim(), style: null, listType: 'number' };
    }

    // Blockquote -> Quote style
    const quote = /^>\s+(.+)$/.exec(trimmed);
    if (quote) {
      return { text: quote[1].trim(), style: 'Quote', listType: null };
    }

    return {
      text: rawLine,
      style: idx === 0 && fallback ? fallback : null,
      listType: null,
    };
  });
}

function applyParagraphStyleSafe(para, styleName) {
  if (!para || !styleName) return;
  try {
    para.style = styleName;
  } catch {
    // Ignore if style name is not supported in this Word build/document.
  }
}

async function applyListFormatting(ctx, Word, entries) {
  let activeList = null;
  let activeType = null;

  for (const { para, block } of entries) {
    if (!block.listType) {
      activeList = null;
      activeType = null;
      continue;
    }
    if (!para || typeof para.startNewList !== 'function') continue;

    try {
      if (!activeList || activeType !== block.listType) {
        activeList = para.startNewList();
        activeType = block.listType;
        if (block.listType === 'number' && typeof activeList.setLevelNumbering === 'function' && Word?.ListNumbering?.arabic) {
          activeList.setLevelNumbering(0, Word.ListNumbering.arabic, [0, '.']);
        }
        activeList.load('id');
        await ctx.sync();
      } else if (typeof para.attachToList === 'function' && activeList.id !== undefined) {
        para.attachToList(activeList.id, 0);
      }
    } catch {
      activeList = null;
      activeType = null;
    }
  }
  if (entries.some(e => e.block.listType)) {
    await ctx.sync();
  }
}

export const WORD_QUICK_ACTIONS = [
  { cmd: '/copy-edit', title: 'Copy-edit document', prompt: 'Copy-edit this document for typos, grammar, wordiness, and awkward phrasing. Propose targeted edits using replace_match.' },
  { cmd: '/check-doc', title: 'Check document', prompt: 'Audit the whole document for structure, clarity, inconsistencies, undefined terms, and gaps. Give me a prioritised read before making edits.' },
  { cmd: '/work-comments', title: 'Work through comments', prompt: 'Work through all open Word comments in this document. For each comment, propose the exact text edit using replace_match and a reply_comment explaining what you changed.' },
  { cmd: '/tighten', title: 'Tighten', prompt: 'Tighten this. Remove filler and hedging while keeping the meaning and voice.' },
  { cmd: '/shorten', title: 'Shorten', prompt: 'Shorten this by about a third, keeping every key fact.' },
  { cmd: '/rewrite', title: 'Rewrite clearly', prompt: 'Rewrite this so it reads clearly and directly, same meaning and approximate length.' },
  { cmd: '/plain', title: 'Plain language', prompt: 'Rewrite in plain language: short sentences, common words, one idea each.' },
  { cmd: '/bullets', title: 'Turn into bullets', prompt: 'Turn this into a crisp bullet list, one idea per bullet.' },
  { cmd: '/zh', title: 'Translate to Chinese (below)', prompt: 'Translate this into natural Chinese and insert it directly underneath.' },
  { cmd: '/en', title: 'Translate to English (below)', prompt: 'Translate this into natural English and insert it directly underneath.' },
  { cmd: '/summarise', title: 'Summarise', prompt: 'Summarise the main points concisely.' },
  { cmd: '/comment', title: 'Add Word comment', prompt: 'Write a short, concrete editorial review comment on this passage as a Word comment.' },
];

async function findMatchingRange(ctx, snippet, { occurrence = 1 } = {}) {
  const clean = String(snippet || '').trim();
  if (!clean) return null;
  const targetOcc = Math.max(1, parseInt(occurrence, 10) || 1);

  // Word's body.search() has a 255-character limit and works best on single-line matches.
  if (clean.length <= 240 && !/[\r\n]/.test(clean)) {
    const results = ctx.document.body.search(clean, { matchCase: true });
    results.load('items');
    await ctx.sync();
    if (results.items.length > 0) {
      return results.items[Math.min(targetOcc - 1, results.items.length - 1)];
    }
  }

  // Fallback: match against paragraph text directly (supports single AND multi-paragraph ranges)
  const paras = ctx.document.body.paragraphs;
  paras.load('items/text');
  await ctx.sync();

  const normalizedTarget = clean.replace(/\s+/g, ' ');
  let matched = 0;
  for (const p of paras.items) {
    const pNorm = (p.text || '').replace(/\s+/g, ' ').trim();
    if (pNorm && (pNorm === normalizedTarget || pNorm.includes(normalizedTarget))) {
      matched++;
      if (matched === targetOcc) return p.getRange();
    }
  }

  // Multi-paragraph span match: check consecutive non-empty paragraphs
  const snippetLines = clean.split(/[\r\n]+/).map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
  if (snippetLines.length >= 2) {
    const firstLine = snippetLines[0];
    const lastLine = snippetLines[snippetLines.length - 1];
    for (let i = 0; i < paras.items.length; i++) {
      const startNorm = (paras.items[i].text || '').replace(/\s+/g, ' ').trim();
      if (!startNorm || (!startNorm.includes(firstLine) && !firstLine.includes(startNorm))) continue;
      for (let j = i; j < Math.min(paras.items.length, i + snippetLines.length + 4); j++) {
        const endNorm = (paras.items[j].text || '').replace(/\s+/g, ' ').trim();
        if (endNorm && (endNorm.includes(lastLine) || lastLine.includes(endNorm))) {
          return paras.items[i].getRange().expandTo(paras.items[j].getRange());
        }
      }
    }
  }

  return null;
}

export function createWordAdapter(env = globalThis) {
  const getWord = () => env.Word;
  const getOffice = () => env.Office;

  return {
    id: 'word',
    label: 'Word',
    noun: 'document',
    placeholder: 'Ask anything about your document, or type / for actions…',
    quickActions: WORD_QUICK_ACTIONS,
    secondaryToggle: {
      key: 'trackChanges',
      label: 'Track',
      titleOn: 'Track Changes is ON (edits appear as native Word redlines)',
      titleOff: 'Track Changes is OFF (click to enable native Word redlines)',
    },
    modeLabels: {
      auto: 'Context: Auto (whole doc + selection)',
      selection: 'Context: Selection only',
      document: 'Context: Whole document',
    },

    formatContextBadge(ctx, contextMode) {
      const selClean = (ctx.selectionText || '').replace(/\s+/g, ' ').trim();
      const docWords = countWords(ctx.documentText || '');
      const openComments = (ctx.comments || []).length;
      const commentSuffix = openComments ? ` · ${openComments} comment${openComments === 1 ? '' : 's'}` : '';

      if (contextMode === 'selection') {
        return {
          badge: selClean ? `Selection · ${selClean.length}c` : 'Selection only',
          preview: selClean ? selClean.slice(0, 80) : 'Select text in Word',
          hasSelection: Boolean(selClean),
        };
      }
      if (contextMode === 'document') {
        return {
          badge: `Whole doc · ${docWords}w`,
          preview: docWords ? `Full document attached${commentSuffix}` : 'Document is empty',
          hasSelection: false,
        };
      }
      if (selClean) {
        return {
          badge: `Selection · ${selClean.length}c`,
          preview: selClean.slice(0, 80),
          hasSelection: true,
        };
      }
      return {
        badge: docWords ? `Document · ${docWords}w` : 'Document',
        preview: docWords ? `Full doc attached${commentSuffix}` : 'Start typing in Word',
        hasSelection: false,
      };
    },

    getStarters(ctx) {
      const hasSelection = Boolean((ctx.selectionText || '').trim());
      const docWords = countWords(ctx.documentText || '');
      const starters = hasSelection
        ? [
            { title: 'Talk through how to improve this', sub: 'Critique and ideas before rewriting', prompt: 'What do you think of the selected passage? Point out what works, what feels weak, and how you would sharpen it before we rewrite.' },
            { title: 'Tighten this passage', sub: 'Cut filler while keeping your voice', prompt: 'Tighten the selected passage. Remove filler and hedging, keep the meaning.' },
            { title: 'Translate to Chinese below', sub: 'Insert natural Chinese underneath', prompt: 'Translate the selected passage into natural Chinese and insert it directly underneath.' },
            { title: 'Explore alternative angles', sub: '2–3 different ways to frame this', prompt: 'Give me 2 or 3 different ways to frame the selected passage, and tell me which one you think lands best.' },
          ]
        : [
            { title: 'Think through this document', sub: 'Honest read on structure, gaps, and flow', prompt: 'Read the document and give me your honest take on the structure, clarity, and what is missing or could be stronger. Just chat with me first before making edits.' },
            { title: 'Where should I edit first?', sub: 'Prioritised list of highest-impact fixes', prompt: 'If we only had 5 minutes to improve this document, which 2 or 3 spots would you fix first and why?' },
            { title: 'Summarise the main argument', sub: 'Quick read on what it actually says', prompt: 'Summarise what this document currently says in 3 crisp bullets, and flag anything that feels unclear.' },
            { title: 'Draft the next section', sub: 'Continue naturally in the same voice', prompt: 'Based on the document so far, draft a concise next paragraph to append at the end.' },
          ];

      const subtitle = hasSelection
        ? `Selection attached (${ctx.selectionText.trim().length} chars) · Ask a question or describe an edit.`
        : docWords
          ? `Reading your document (${docWords} words) · Ask anything or select text to focus.`
          : 'Chat about your document, or select any text to edit it in place.';

      return { heading: 'What are we working on?', subtitle, starters };
    },

    async readContext() {
      const Word = getWord();
      const Office = getOffice();
      if (!Word || !Word.run) return null;

      let result = null;
      await Word.run(async ctx => {
        const sel = ctx.document.getSelection();
        const body = ctx.document.body;
        sel.load('text');
        body.load('text');
        const canTrack = Boolean(Office?.context?.requirements?.isSetSupported?.('WordApi', '1.5'));
        if (canTrack) ctx.document.load('changeTrackingMode');
        await ctx.sync();

        let trackChanges;
        if (canTrack && ctx.document.changeTrackingMode !== undefined) {
          trackChanges = ctx.document.changeTrackingMode !== Word.ChangeTrackingMode.off;
        }

        let comments = [];
        try {
          if (typeof body.getComments === 'function') {
            const commentsCol = body.getComments();
            commentsCol.load('items/id,items/content,items/authorName,items/resolved');
            await ctx.sync();
            const openItems = (commentsCol.items || []).filter(c => !c.resolved).slice(0, 30);
            const ranges = openItems.map(c => {
              const r = c.getRange();
              r.load('text');
              return r;
            });
            if (ranges.length) await ctx.sync();
            comments = openItems.map((c, idx) => ({
              id: c.id,
              author: c.authorName || 'Reviewer',
              content: c.content || '',
              anchor: (ranges[idx]?.text || '').trim(),
            }));
          }
        } catch {
          // Comments API unavailable on older Word builds.
        }

        let styles = [];
        try {
          const styleSet = new Set();
          if (typeof ctx.document.getStyles === 'function') {
            const docStyles = ctx.document.getStyles();
            docStyles.load('items/nameLocal,items/type,items/builtIn,items/inUse');
            await ctx.sync();
            const items = docStyles.items || [];
            // Prioritize in-use styles, custom template styles, and core structural styles
            for (const s of items) {
              const isPara = !s.type || String(s.type).toLowerCase() === 'paragraph';
              if (!isPara || !s.nameLocal) continue;
              if (s.inUse || s.builtIn === false || STANDARD_STYLES.includes(s.nameLocal)) {
                styleSet.add(s.nameLocal);
              }
            }
          }
          // Also include styles actually present on paragraphs in the body
          if (body.paragraphs) {
            body.paragraphs.load('items/style');
            await ctx.sync();
            for (const p of body.paragraphs.items || []) {
              if (p.style) styleSet.add(p.style);
            }
          }
          for (const std of STANDARD_STYLES) styleSet.add(std);
          styles = [...styleSet].slice(0, 30);
        } catch {
          styles = [...STANDARD_STYLES];
        }

        result = {
          selectionText: sel.text || '',
          documentText: body.text || '',
          comments,
          hostContext: { styles },
          trackChanges,
        };
      });
      return result;
    },

    watchSelection(onChange, onFallback) {
      const Office = getOffice();
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange, result => {
          if (result.status !== Office.AsyncResultStatus.Succeeded && onFallback) {
            onFallback('Auto-selection unavailable here. Use the refresh icon after selecting.');
          }
        });
      } catch {
        if (onFallback) onFallback('Auto-selection unavailable here. Use the refresh icon after selecting.');
      }
    },

    async toggleSecondary(enabled) {
      const Word = getWord();
      const Office = getOffice();
      if (!Word || !Word.run) {
        return { ok: true, value: enabled, message: enabled ? 'Track Changes ON' : 'Track Changes OFF' };
      }
      await Word.run(async ctx => {
        if (!Office?.context?.requirements?.isSetSupported?.('WordApi', '1.5')) {
          throw new Error('This Word version does not support toggling Track Changes from add-ins. Toggle it in the Review tab.');
        }
        ctx.document.changeTrackingMode = enabled ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
        await ctx.sync();
      });
      return {
        ok: true,
        value: enabled,
        message: enabled ? 'Track Changes ON · edits will show as redlines' : 'Track Changes OFF',
      };
    },

    async applyEdit(edit, { expectedSelection = '', currentSelection = '', overrideAction = null, trackChanges = false } = {}) {
      const Word = getWord();
      const action = overrideAction || edit.action;
      const content = (edit.content || '').trim();
      if (!content) return { ok: false, message: 'Nothing to apply.' };
      if (!Word || !Word.run) {
        return { ok: false, message: 'Open inside Microsoft Word to apply edits directly.' };
      }

      const defaultStyle = edit.attrs?.style || '';

      await Word.run(async ctx => {
        const sel = ctx.document.getSelection();
        sel.load('text');
        await ctx.sync();

        if (action === 'append_doc') {
          const blocks = parseStyledWordBlocks(content, defaultStyle);
          const created = [];
          for (const block of blocks) {
            const para = ctx.document.body.insertParagraph(block.text, Word.InsertLocation.end);
            applyParagraphStyleSafe(para, block.style);
            created.push({ para, block });
          }
          await ctx.sync();
          await applyListFormatting(ctx, Word, created);
          return;
        }

        if (action === 'comment') {
          if (typeof sel.insertComment !== 'function') {
            throw new Error('This version of Word does not support adding comments from add-ins.');
          }
          sel.insertComment(content);
          await ctx.sync();
          return;
        }

        if (action === 'reply_comment') {
          if (typeof ctx.document.body.getComments !== 'function') {
            throw new Error('This version of Word does not support replying to comments from add-ins.');
          }
          const commentsCol = ctx.document.body.getComments();
          commentsCol.load('items/id');
          await ctx.sync();
          const targetComment = commentsCol.items.find(c => String(c.id) === String(edit.commentId));
          if (!targetComment) {
            throw new Error('Could not find that comment thread in the document.');
          }
          targetComment.reply(content);
          await ctx.sync();
          return;
        }

        if (action === 'replace_match') {
          const target = await findMatchingRange(
            ctx,
            edit.find || expectedSelection || currentSelection,
            { occurrence: edit.attrs?.occurrence },
          );
          if (!target) {
            throw new Error('Could not find the original passage in the document. Select it manually and use Replace.');
          }
          const blocks = parseStyledWordBlocks(content, defaultStyle);
          const created = [];
          target.insertText(blocks[0].text, Word.InsertLocation.replace);
          const firstPara = target.paragraphs?.getFirstOrNullObject?.() || target.paragraphs?.getFirst?.() || null;
          if (firstPara && blocks[0].style) applyParagraphStyleSafe(firstPara, blocks[0].style);
          created.push({ para: firstPara, block: blocks[0] });
          const trailing = [];
          for (let i = blocks.length - 1; i >= 1; i--) {
            const para = target.insertParagraph(blocks[i].text, Word.InsertLocation.after);
            applyParagraphStyleSafe(para, blocks[i].style);
            trailing.unshift({ para, block: blocks[i] });
          }
          await ctx.sync();
          await applyListFormatting(ctx, Word, [...created, ...trailing]);
          return;
        }

        if (action === 'insert_after' || action === 'insert_before') {
          const blocks = parseStyledWordBlocks(content, defaultStyle);
          const created = [];
          if (action === 'insert_after') {
            for (let i = blocks.length - 1; i >= 0; i--) {
              const para = sel.insertParagraph(blocks[i].text, Word.InsertLocation.after);
              applyParagraphStyleSafe(para, blocks[i].style);
              created.unshift({ para, block: blocks[i] });
            }
          } else {
            for (const block of blocks) {
              const para = sel.insertParagraph(block.text, Word.InsertLocation.before);
              applyParagraphStyleSafe(para, block.style);
              created.push({ para, block });
            }
          }
          await ctx.sync();
          await applyListFormatting(ctx, Word, created);
          return;
        }

        // Default: replace_selection
        let targetRange = sel;
        const selText = (sel.text || '').trim();
        const expectedClean = (expectedSelection || '').trim();
        const fallbackClean = expectedClean || (currentSelection || '').trim();

        if (!selText || (expectedClean && sel.text !== expectedSelection)) {
          if (fallbackClean) {
            const located = await findMatchingRange(ctx, fallbackClean);
            if (located) {
              targetRange = located;
            } else if (!selText) {
              throw new Error('Could not locate the selected passage. Re-select it in Word or use Insert below.');
            }
          } else {
            throw new Error('Nothing is selected in Word. Select text first, or use Insert below.');
          }
        }

        const blocks = parseStyledWordBlocks(content, defaultStyle);
        const created = [];
        targetRange.insertText(blocks[0].text, Word.InsertLocation.replace);
        const firstPara = targetRange.paragraphs?.getFirstOrNullObject?.() || targetRange.paragraphs?.getFirst?.() || null;
        if (firstPara && blocks[0].style) applyParagraphStyleSafe(firstPara, blocks[0].style);
        created.push({ para: firstPara, block: blocks[0] });
        const trailing = [];
        for (let i = blocks.length - 1; i >= 1; i--) {
          const para = targetRange.insertParagraph(blocks[i].text, Word.InsertLocation.after);
          applyParagraphStyleSafe(para, blocks[i].style);
          trailing.unshift({ para, block: blocks[i] });
        }
        await ctx.sync();
        await applyListFormatting(ctx, Word, [...created, ...trailing]);
      });

      const suffix = trackChanges ? ' (tracked change)' : ' · Cmd+Z to undo';
      const msgMap = {
        replace_selection: `Replaced in document${suffix}`,
        replace_match: `Edit applied${suffix}`,
        insert_after: `Inserted below${suffix}`,
        insert_before: `Inserted above${suffix}`,
        append_doc: `Appended to document${suffix}`,
        comment: 'Comment added in Word',
        reply_comment: 'Posted reply to Word comment',
      };
      return { ok: true, message: msgMap[action] || `Applied${suffix}` };
    },
  };
}
