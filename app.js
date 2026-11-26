/* global Office, Word */
import {
  parseAssistantReply,
  actionLabel,
  primaryButtonLabel,
  renderMarkdown,
  renderDiffHtml,
  countWords,
} from './lib/format.js';

const $ = id => document.getElementById(id);

const QUICK_ACTIONS = [
  { cmd: '/tighten', title: 'Tighten', prompt: 'Tighten this. Remove filler and hedging while keeping the meaning and voice.' },
  { cmd: '/shorten', title: 'Shorten', prompt: 'Shorten this by about a third, keeping every key fact.' },
  { cmd: '/rewrite', title: 'Rewrite clearly', prompt: 'Rewrite this so it reads clearly and directly, same meaning and approximate length.' },
  { cmd: '/plain', title: 'Plain language', prompt: 'Rewrite in plain language: short sentences, common words, one idea each.' },
  { cmd: '/bullets', title: 'Turn into bullets', prompt: 'Turn this into a crisp bullet list, one idea per bullet.' },
  { cmd: '/zh', title: 'Translate to Chinese (below)', prompt: 'Translate this into natural Chinese and insert it directly underneath.' },
  { cmd: '/en', title: 'Translate to English (below)', prompt: 'Translate this into natural English and insert it directly underneath.' },
  { cmd: '/summarise', title: 'Summarise', prompt: 'Summarise the main points concisely.' },
  { cmd: '/critique', title: 'Critique & improve', prompt: 'Critique this for clarity, structure, and weak spots. Point out what to fix first.' },
  { cmd: '/comment', title: 'Add Word comment', prompt: 'Write a short, concrete editorial review comment on this passage as a Word comment.' },
];

const CONTEXT_MODES = ['auto', 'selection', 'document'];

const state = {
  inWord: false,
  messages: [],          // { role, content } sent to the model
  selectionText: '',     // current Word selection
  documentText: '',      // full Word document body text
  contextMode: 'auto',   // 'auto' | 'selection' | 'document'
  controller: null,
  readTimer: null,
  reading: false,
  toastTimer: null,
  menuOpen: false,
  menuIndex: 0,
  filteredActions: QUICK_ACTIONS,
  userScrolledUp: false,
  keyInfo: null,
};

/* ---------- Toast notifications ---------- */

function showToast(text, isError = false, durationMs = 3200) {
  clearTimeout(state.toastTimer);
  const toast = $('toast');
  $('toast-text').textContent = text;
  toast.classList.toggle('err', isError);
  toast.hidden = false;
  if (!isError && durationMs > 0) {
    state.toastTimer = setTimeout(() => { toast.hidden = true; }, durationMs);
  }
}

$('toast-close').onclick = () => {
  clearTimeout(state.toastTimer);
  $('toast').hidden = true;
};

/* ---------- Document & selection reading ---------- */

function updateContextUI() {
  const badge = $('context-badge');
  const preview = $('context-preview');
  const selClean = state.selectionText.replace(/\s+/g, ' ').trim();
  const docWords = countWords(state.documentText);

  if (state.contextMode === 'selection') {
    badge.textContent = selClean ? `Selection · ${selClean.length}c` : 'Selection only';
    badge.classList.toggle('has-selection', Boolean(selClean));
    preview.textContent = selClean ? selClean.slice(0, 80) : 'Select text in Word';
  } else if (state.contextMode === 'document') {
    badge.textContent = `Whole doc · ${docWords}w`;
    badge.classList.remove('has-selection');
    preview.textContent = docWords ? 'Full document attached' : 'Document is empty';
  } else {
    // Auto: full doc in background + active selection highlighted
    if (selClean) {
      badge.textContent = `Selection · ${selClean.length}c`;
      badge.classList.add('has-selection');
      preview.textContent = selClean.slice(0, 80);
    } else {
      badge.textContent = docWords ? `Document · ${docWords}w` : 'Document';
      badge.classList.remove('has-selection');
      preview.textContent = docWords ? 'Full doc attached · select to focus' : 'Start typing in Word';
    }
  }

  if (state.messages.length === 0) renderEmptyState();
}

async function readWordContext() {
  if (!state.inWord || state.reading) return;
  state.reading = true;
  try {
    await Word.run(async ctx => {
      const sel = ctx.document.getSelection();
      const body = ctx.document.body;
      sel.load('text');
      body.load('text');
      await ctx.sync();
      const changed = sel.text !== state.selectionText || body.text !== state.documentText;
      state.selectionText = sel.text || '';
      state.documentText = body.text || '';
      if (changed) updateContextUI();
    });
  } catch {
    // Word selection changes can fire mid-sync; the next event catches up.
  } finally {
    state.reading = false;
  }
}

function scheduleRead() {
  clearTimeout(state.readTimer);
  state.readTimer = setTimeout(readWordContext, 250);
}

function watchSelection() {
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRead, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showToast('Auto-selection unavailable here. Use the refresh icon after selecting.');
      }
    });
  } catch {
    showToast('Auto-selection unavailable here. Use the refresh icon after selecting.');
  }
  readWordContext();
}

$('refresh').onclick = async () => {
  await readWordContext();
  showToast('Document context refreshed.', false, 1800);
};

$('context-pill').onclick = () => {
  const idx = CONTEXT_MODES.indexOf(state.contextMode);
  state.contextMode = CONTEXT_MODES[(idx + 1) % CONTEXT_MODES.length];
  updateContextUI();
  const labels = {
    auto: 'Context: Auto (whole doc + selection)',
    selection: 'Context: Selection only',
    document: 'Context: Whole document',
  };
  showToast(labels[state.contextMode], false, 2000);
};

/* ---------- Smart auto-scroll ---------- */

const transcriptEl = $('transcript');
transcriptEl.addEventListener('scroll', () => {
  const distFromBottom = transcriptEl.scrollHeight - transcriptEl.scrollTop - transcriptEl.clientHeight;
  state.userScrolledUp = distFromBottom > 60;
});

function scrollToBottom(force = false) {
  if (force || !state.userScrolledUp) {
    transcriptEl.scrollTop = transcriptEl.scrollHeight;
  }
}

/* ---------- Empty state ---------- */

function renderEmptyState() {
  if (state.messages.length > 0) return;
  const hasSelection = Boolean(state.selectionText.trim());
  const docWords = countWords(state.documentText);

  const starters = hasSelection
    ? [
        { title: 'Tighten this passage', sub: 'Cut filler while keeping your voice', prompt: 'Tighten the selected passage. Remove filler and hedging, keep the meaning.' },
        { title: 'Translate to Chinese below', sub: 'Insert natural Chinese underneath', prompt: 'Translate the selected passage into natural Chinese and insert it directly underneath.' },
        { title: 'Critique what could be stronger', sub: 'Direct feedback without rewriting yet', prompt: 'Critique this passage. What is unclear, wordy, or weak?' },
        { title: 'Turn into crisp bullets', sub: 'One clear idea per bullet', prompt: 'Turn the selected passage into a short bullet list, one idea per bullet.' },
      ]
    : [
        { title: 'Summarise this document', sub: 'Key takeaways in 3–4 bullets', prompt: 'Summarise the key points of this document in 3 to 4 crisp bullets.' },
        { title: 'Review for clarity & flow', sub: 'Find weak spots and suggest edits', prompt: 'Review this document for clarity, structure, and flow. Point out the main issues and suggest concrete edits.' },
        { title: 'Draft the next section', sub: 'Continue naturally in the same voice', prompt: 'Based on the document so far, draft a concise next paragraph to append at the end.' },
        { title: 'Tighten the whole draft', sub: 'Propose targeted edits across the doc', prompt: 'Identify the wordiest sentences in this document and propose sharper replacements.' },
      ];

  const subtitle = hasSelection
    ? `Selection attached (${state.selectionText.trim().length} chars) · Ask a question or describe an edit.`
    : docWords
      ? `Reading your document (${docWords} words) · Ask anything or select text to focus.`
      : 'Chat about your document, or select any text to edit it in place.';

  transcriptEl.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'empty-state';
  wrap.innerHTML = `<div class="empty-head"><h2>What are we working on?</h2><p></p></div><div class="starter-grid"></div>`;
  wrap.querySelector('p').textContent = subtitle;

  const grid = wrap.querySelector('.starter-grid');
  for (const s of starters) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'starter-btn';
    btn.innerHTML = `<span class="starter-title"></span><span class="starter-sub"></span>`;
    btn.querySelector('.starter-title').textContent = s.title;
    btn.querySelector('.starter-sub').textContent = s.sub;
    btn.onclick = () => sendTurn(s.prompt);
    grid.appendChild(btn);
  }
  transcriptEl.appendChild(wrap);
}

/* ---------- Writing edits back to Word ---------- */

function splitParagraphs(text) {
  return String(text).replace(/\r\n?/g, '\n').split('\n').filter((line, idx, arr) => {
    // Keep intentional blank lines inside multi-paragraph text, trim only leading/trailing empty
    if (idx === 0 || idx === arr.length - 1) return line.trim().length > 0;
    return true;
  });
}

async function findMatchingRange(ctx, snippet) {
  const clean = snippet.trim();
  if (!clean) return null;

  // Word's body.search() has a 255-character limit.
  if (clean.length <= 240) {
    const results = ctx.document.body.search(clean, { matchCase: true });
    results.load('items');
    await ctx.sync();
    if (results.items.length > 0) return results.items[0];
  }

  // Fallback: match against paragraph text directly (works for long passages)
  const paras = ctx.document.body.paragraphs;
  paras.load('items/text');
  await ctx.sync();
  const normalizedTarget = clean.replace(/\s+/g, ' ');
  for (const p of paras.items) {
    const pNorm = (p.text || '').replace(/\s+/g, ' ').trim();
    if (pNorm === normalizedTarget || pNorm.includes(normalizedTarget)) {
      return p.getRange();
    }
  }
  return null;
}

async function applyEdit(edit, expectedSelection = '', overrideAction = null) {
  const action = overrideAction || edit.action;
  const content = (edit.content || '').trim();
  if (!content) return false;

  if (!state.inWord) {
    showToast('Open inside Microsoft Word to apply edits directly.', true);
    return false;
  }

  try {
    await Word.run(async ctx => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();

      if (action === 'append_doc') {
        const lines = splitParagraphs(content);
        for (const line of lines) {
          ctx.document.body.insertParagraph(line, Word.InsertLocation.end);
        }
        await ctx.sync();
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

      if (action === 'replace_match') {
        const target = await findMatchingRange(ctx, edit.find || expectedSelection);
        if (!target) {
          throw new Error('Could not find the original passage in the document. Select it manually and use Replace.');
        }
        const lines = splitParagraphs(content);
        target.insertText(lines[0], Word.InsertLocation.replace);
        for (let i = lines.length - 1; i >= 1; i--) {
          target.insertParagraph(lines[i], Word.InsertLocation.after);
        }
        await ctx.sync();
        return;
      }

      if (action === 'insert_after' || action === 'insert_before') {
        const lines = splitParagraphs(content);
        if (action === 'insert_after') {
          // Insert in reverse order after the selection so paragraphs stay in exact order
          for (let i = lines.length - 1; i >= 0; i--) {
            sel.insertParagraph(lines[i], Word.InsertLocation.after);
          }
        } else {
          for (const line of lines) {
            sel.insertParagraph(line, Word.InsertLocation.before);
          }
        }
        await ctx.sync();
        return;
      }

      // Default: replace_selection
      let targetRange = sel;
      const selText = (sel.text || '').trim();
      const expectedClean = (expectedSelection || '').trim();

      if (!selText || (expectedClean && sel.text !== expectedSelection)) {
        // Smart recovery: if the user clicked elsewhere, try locating the original passage!
        if (expectedClean) {
          const located = await findMatchingRange(ctx, expectedClean);
          if (located) {
            targetRange = located;
          } else {
            throw new Error('Selection changed since this suggestion. Re-select the passage or use Insert below.');
          }
        } else if (!selText) {
          throw new Error('Nothing is selected in Word. Select text first, or use Insert below.');
        }
      }

      const lines = splitParagraphs(content);
      targetRange.insertText(lines[0], Word.InsertLocation.replace);
      for (let i = lines.length - 1; i >= 1; i--) {
        targetRange.insertParagraph(lines[i], Word.InsertLocation.after);
      }
      await ctx.sync();
    });

    const msgMap = {
      replace_selection: 'Replaced in document · Cmd+Z to undo',
      replace_match: 'Edit applied · Cmd+Z to undo',
      insert_after: 'Inserted below · Cmd+Z to undo',
      insert_before: 'Inserted above · Cmd+Z to undo',
      append_doc: 'Appended to document · Cmd+Z to undo',
      comment: 'Comment added in Word',
    };
    showToast(msgMap[action] || 'Applied · Cmd+Z to undo');
    scheduleRead();
    return true;
  } catch (e) {
    showToast(e.message || String(e), true);
    return false;
  }
}

/* ---------- Rendering conversation turns ---------- */

function addUserTurn(text, attachedSelection) {
  if (state.messages.length === 0) transcriptEl.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'turn user';

  const cleanSel = (attachedSelection || '').replace(/\s+/g, ' ').trim();
  if (cleanSel && state.contextMode !== 'document') {
    const tag = document.createElement('div');
    tag.className = 'user-context-tag';
    tag.textContent = `Selection: "${cleanSel.slice(0, 48)}${cleanSel.length > 48 ? '…' : ''}"`;
    tag.title = attachedSelection;
    wrap.appendChild(tag);
  }

  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  bubble.textContent = text;
  wrap.appendChild(bubble);
  transcriptEl.appendChild(wrap);
  scrollToBottom(true);
}

function createAssistantTurn() {
  const wrap = document.createElement('div');
  wrap.className = 'turn assistant';
  const body = document.createElement('div');
  body.className = 'assistant-body';
  body.innerHTML = `<div class="thinking"><span class="thinking-dots"><span></span><span></span><span></span></span><span>Thinking…</span></div>`;
  const meta = document.createElement('div');
  meta.className = 'turn-meta';
  meta.hidden = true;
  wrap.append(body, meta);
  transcriptEl.appendChild(wrap);
  scrollToBottom(true);
  return { wrap, body, meta, diffStates: new Map(), appliedSet: new Set() };
}

function buildEditCard(edit, index, turn, expectedSelection) {
  const card = document.createElement('div');
  card.className = 'edit-card';
  if (turn.appliedSet.has(index)) card.classList.add('applied');

  const head = document.createElement('div');
  head.className = 'edit-card-head';

  const badge = document.createElement('span');
  badge.className = 'edit-badge';
  badge.textContent = actionLabel(edit.action, edit.label);
  head.appendChild(badge);

  const baselineText = edit.action === 'replace_match'
    ? edit.find
    : edit.action === 'replace_selection'
      ? expectedSelection
      : '';
  const canDiff = Boolean(baselineText && baselineText.trim() && edit.content.trim() && !edit.streaming);

  let diffBtn = null;
  if (canDiff) {
    const showingDiff = Boolean(turn.diffStates.get(index));
    diffBtn = document.createElement('button');
    diffBtn.type = 'button';
    diffBtn.className = 'diff-toggle' + (showingDiff ? ' active' : '');
    diffBtn.textContent = showingDiff ? 'Preview' : 'Diff';
    head.appendChild(diffBtn);
  }

  card.appendChild(head);

  if (edit.action === 'replace_match' && edit.find && !turn.diffStates.get(index)) {
    const findLine = document.createElement('div');
    findLine.className = 'edit-find-preview';
    findLine.textContent = `Replacing: "${edit.find.replace(/\s+/g, ' ').trim()}"`;
    findLine.title = edit.find;
    card.appendChild(findLine);
  }

  const body = document.createElement('div');
  body.className = 'edit-card-body';

  const renderCardBody = () => {
    if (canDiff && turn.diffStates.get(index)) {
      body.innerHTML = renderDiffHtml(baselineText, edit.content);
    } else {
      body.textContent = edit.content || (edit.streaming ? 'Writing…' : '');
    }
  };
  renderCardBody();
  card.appendChild(body);

  if (diffBtn) {
    diffBtn.onclick = () => {
      const next = !turn.diffStates.get(index);
      turn.diffStates.set(index, next);
      diffBtn.classList.toggle('active', next);
      diffBtn.textContent = next ? 'Preview' : 'Diff';
      renderCardBody();
    };
  }

  if (!edit.streaming) {
    const foot = document.createElement('div');
    foot.className = 'edit-card-foot';

    const primary = document.createElement('button');
    primary.type = 'button';
    primary.className = 'btn-primary' + (turn.appliedSet.has(index) ? ' done' : '');
    primary.textContent = turn.appliedSet.has(index) ? 'Applied ✓' : primaryButtonLabel(edit.action);
    primary.onclick = async () => {
      const ok = await applyEdit(edit, expectedSelection);
      if (ok) {
        turn.appliedSet.add(index);
        card.classList.add('applied');
        primary.classList.add('done');
        primary.textContent = 'Applied ✓';
      }
    };
    foot.appendChild(primary);

    if (edit.action === 'replace_selection') {
      const insertBtn = document.createElement('button');
      insertBtn.type = 'button';
      insertBtn.className = 'btn-secondary';
      insertBtn.textContent = 'Insert below';
      insertBtn.onclick = async () => {
        const ok = await applyEdit(edit, expectedSelection, 'insert_after');
        if (ok) {
          turn.appliedSet.add(index);
          card.classList.add('applied');
        }
      };
      foot.appendChild(insertBtn);
    }

    const copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.className = 'btn-secondary';
    copyBtn.textContent = 'Copy';
    copyBtn.onclick = async () => {
      try {
        await navigator.clipboard.writeText(edit.content);
        showToast('Copied to clipboard.', false, 1800);
      } catch {
        showToast('Copy failed.', true);
      }
    };
    foot.appendChild(copyBtn);

    card.appendChild(foot);
  }

  return card;
}

function renderAssistantTurn(turn, text, expectedSelection, isFinal = false, metaLabel = '') {
  const parts = parseAssistantReply(text);
  turn.body.innerHTML = '';

  if (!parts.length) {
    turn.body.innerHTML = `<div class="thinking"><span class="thinking-dots"><span></span><span></span><span></span></span><span>Thinking…</span></div>`;
    return;
  }

  const edits = [];
  let editIndex = 0;

  for (const part of parts) {
    if (part.kind === 'prose') {
      const prose = document.createElement('div');
      prose.className = 'prose';
      prose.innerHTML = renderMarkdown(part.text);
      turn.body.appendChild(prose);
    } else if (part.kind === 'edit') {
      const idx = editIndex++;
      edits.push({ edit: part, idx });
      turn.body.appendChild(buildEditCard(part, idx, turn, expectedSelection));
    }
  }

  if (isFinal && edits.length >= 2) {
    const bar = document.createElement('div');
    bar.className = 'apply-all-bar';
    const allBtn = document.createElement('button');
    allBtn.type = 'button';
    allBtn.className = 'btn-primary';
    allBtn.textContent = `Apply all (${edits.length})`;
    allBtn.onclick = async () => {
      let appliedCount = 0;
      for (const { edit, idx } of edits) {
        if (turn.appliedSet.has(idx)) continue;
        const ok = await applyEdit(edit, expectedSelection);
        if (ok) {
          turn.appliedSet.add(idx);
          appliedCount++;
        }
      }
      if (appliedCount > 0) {
        renderAssistantTurn(turn, text, expectedSelection, true, metaLabel);
        showToast(`Applied ${appliedCount} edit${appliedCount === 1 ? '' : 's'} · Cmd+Z to undo`);
      }
    };
    bar.appendChild(allBtn);
    turn.body.appendChild(bar);
  }

  if (isFinal) {
    turn.meta.hidden = false;
    turn.meta.innerHTML = '';
    const left = document.createElement('span');
    left.textContent = metaLabel;
    const copyAll = document.createElement('button');
    copyAll.type = 'button';
    copyAll.className = 'meta-btn';
    copyAll.textContent = 'Copy reply';
    copyAll.onclick = async () => {
      try {
        await navigator.clipboard.writeText(text.replace(/<\/?(?:edit[^>]*|replacement|find|content)>/gi, '').trim());
        showToast('Reply copied.', false, 1600);
      } catch {
        showToast('Copy failed.', true);
      }
    };
    turn.meta.append(left, copyAll);
  }

  scrollToBottom(false);
}

/* ---------- Composer, slash commands & quick actions ---------- */

const inputEl = $('input');

function autoResizeInput() {
  inputEl.style.height = 'auto';
  inputEl.style.height = Math.min(inputEl.scrollHeight, 140) + 'px';
}

function updateSendState() {
  const busy = Boolean(state.controller);
  $('send').hidden = busy;
  $('stop').hidden = !busy;
  $('send').disabled = busy || !inputEl.value.trim();
}

function renderActionsMenu(filter = '') {
  const q = filter.trim().toLowerCase();
  state.filteredActions = q
    ? QUICK_ACTIONS.filter(a => a.cmd.includes(q) || a.title.toLowerCase().includes(q))
    : QUICK_ACTIONS;

  const menu = $('actions-menu');
  menu.innerHTML = '';
  if (!state.filteredActions.length) {
    closeActionsMenu();
    return;
  }

  state.menuIndex = Math.min(state.menuIndex, state.filteredActions.length - 1);
  state.filteredActions.forEach((action, idx) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'action-item' + (idx === state.menuIndex ? ' active' : '');
    btn.setAttribute('role', 'option');
    btn.innerHTML = `<span class="action-item-title"></span><span class="action-item-cmd"></span>`;
    btn.querySelector('.action-item-title').textContent = action.title;
    btn.querySelector('.action-item-cmd').textContent = action.cmd;
    btn.onclick = () => {
      closeActionsMenu();
      inputEl.value = '';
      autoResizeInput();
      updateSendState();
      sendTurn(action.prompt);
    };
    menu.appendChild(btn);
  });

  menu.hidden = false;
  state.menuOpen = true;
  $('actions-btn').classList.add('open');
}

function closeActionsMenu() {
  $('actions-menu').hidden = true;
  state.menuOpen = false;
  state.menuIndex = 0;
  $('actions-btn').classList.remove('open');
}

$('actions-btn').onclick = () => {
  if (state.menuOpen) closeActionsMenu();
  else renderActionsMenu('');
};

inputEl.addEventListener('input', () => {
  autoResizeInput();
  updateSendState();
  const val = inputEl.value;
  if (val.startsWith('/') && !val.includes(' ') && !val.includes('\n')) {
    renderActionsMenu(val);
  } else if (state.menuOpen) {
    closeActionsMenu();
  }
});

inputEl.addEventListener('keydown', e => {
  if (state.menuOpen) {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      state.menuIndex = (state.menuIndex + 1) % state.filteredActions.length;
      renderActionsMenu(inputEl.value.startsWith('/') ? inputEl.value : '');
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      state.menuIndex = (state.menuIndex - 1 + state.filteredActions.length) % state.filteredActions.length;
      renderActionsMenu(inputEl.value.startsWith('/') ? inputEl.value : '');
      return;
    }
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      const chosen = state.filteredActions[state.menuIndex];
      if (chosen) {
        closeActionsMenu();
        inputEl.value = '';
        autoResizeInput();
        updateSendState();
        sendTurn(chosen.prompt);
      }
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      closeActionsMenu();
      return;
    }
  }

  if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    sendTurn(inputEl.value);
  } else if (e.key === 'Escape' && state.controller) {
    state.controller.abort();
  }
});

document.addEventListener('click', e => {
  if (state.menuOpen && !$('actions-menu').contains(e.target) && !$('actions-btn').contains(e.target)) {
    closeActionsMenu();
  }
});

/* ---------- Providers & chat streaming ---------- */

async function loadProviders() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    const sel = $('provider');
    sel.innerHTML = '';
    for (const p of data.providers) {
      const opt = document.createElement('option');
      opt.value = p.id;
      opt.textContent = p.label;
      opt.title = `${p.label} · ${p.model}`;
      if (p.id === data.active.provider) opt.selected = true;
      sel.appendChild(opt);
    }
    sel.onchange = async () => {
      const res2 = await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ provider: sel.value }),
      });
      const cfg = await res2.json();
      if (cfg.error) showToast(cfg.error, true);
      else showToast(`Switched to ${sel.options[sel.selectedIndex].textContent} (${cfg.active.model})`, false, 2000);
    };
  } catch {
    showToast('Cannot reach the local pane server. Is the login agent running?', true);
  }
}

async function sendTurn(userText) {
  const text = String(userText || '').trim();
  if (!text || state.controller) return;

  closeActionsMenu();
  inputEl.value = '';
  autoResizeInput();

  // Pull fresh document/selection right before sending so the model sees recent typing
  await readWordContext();

  const usedSelection = state.selectionText;
  const usedDoc = state.documentText;

  addUserTurn(text, usedSelection);
  state.messages.push({ role: 'user', content: text });

  const turn = createAssistantTurn();
  let reply = '';
  let metaLabel = '';
  const controller = new AbortController();
  state.controller = controller;
  state.userScrolledUp = false;
  updateSendState();

  try {
    const res = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        messages: state.messages,
        selection: usedSelection,
        documentText: usedDoc,
        contextMode: state.contextMode,
      }),
      signal: controller.signal,
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: `Server returned ${res.status}` }));
      if (res.status === 502 && /No API key/i.test(err.error || '')) {
        openKeysPanel($('provider').value);
      }
      throw new Error(err.error || `Server returned ${res.status}`);
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';

    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      let i;
      while ((i = buf.indexOf('\n\n')) >= 0) {
        const frame = buf.slice(0, i);
        buf = buf.slice(i + 2);
        if (!frame.startsWith('data:')) continue;
        const payload = JSON.parse(frame.slice(5).trim());
        if (payload.delta) {
          reply += payload.delta;
          renderAssistantTurn(turn, reply, usedSelection, false);
        }
        if (payload.error) throw new Error(payload.error);
        if (payload.done) {
          metaLabel = payload.providerLabel ? `${payload.providerLabel} · ${payload.model}` : payload.model;
        }
      }
    }

    if (!reply.trim()) throw new Error('The model returned nothing.');
    state.messages.push({ role: 'assistant', content: reply });
    renderAssistantTurn(turn, reply, usedSelection, true, metaLabel);
  } catch (e) {
    if (e.name === 'AbortError') {
      showToast('Stopped.', false, 1800);
      if (reply.trim()) {
        state.messages.push({ role: 'assistant', content: reply });
        renderAssistantTurn(turn, reply, usedSelection, true, 'Stopped');
      } else {
        turn.wrap.remove();
        state.messages.pop();
        if (state.messages.length === 0) renderEmptyState();
      }
    } else {
      showToast(e.message || String(e), true);
      if (reply.trim()) {
        renderAssistantTurn(turn, reply, usedSelection, true, 'Error');
      } else {
        turn.wrap.remove();
        state.messages.pop();
        if (state.messages.length === 0) renderEmptyState();
      }
    }
  } finally {
    state.controller = null;
    updateSendState();
    inputEl.focus();
  }
}

$('send').onclick = () => sendTurn(inputEl.value);
$('stop').onclick = () => { if (state.controller) state.controller.abort(); };
$('new').onclick = () => {
  if (state.controller) state.controller.abort();
  state.messages = [];
  renderEmptyState();
  showToast('Started a new conversation.', false, 1600);
  inputEl.focus();
};

/* ---------- Local API Keys panel (write-only to ~/.config/word-poc/.env) ---------- */

function syncKeysFormSelection() {
  if (!state.keyInfo) return;
  const sel = $('keys-provider');
  const current = state.keyInfo.providers.find(p => p.id === sel.value);
  if (!current) return;
  $('keys-input-label').textContent = `${current.label} (${current.envKey})`;
  $('keys-input').placeholder = current.configured ? 'Paste a new key to replace…' : `Paste ${current.envKey}…`;
  $('keys-remove').hidden = current.source !== 'config';
  document.querySelectorAll('#keys-list .key-chip').forEach(chip => {
    chip.classList.toggle('selected', chip.dataset.id === current.id);
  });
}

function renderKeysPanel(data, focusProvider = null) {
  state.keyInfo = data;
  const list = $('keys-list');
  const sel = $('keys-provider');
  const prevChoice = focusProvider || sel.value || $('provider').value;
  list.innerHTML = '';
  sel.innerHTML = '';

  for (const p of data.providers) {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'key-chip';
    chip.dataset.id = p.id;
    chip.innerHTML = `<span class="key-dot${p.configured ? ' ok' : ''}"></span><span></span>`;
    chip.querySelector('span:last-child').textContent = p.label;
    chip.title = p.configured ? `${p.label}: configured (${p.source})` : `${p.label}: not saved in .env`;
    chip.onclick = () => {
      sel.value = p.id;
      syncKeysFormSelection();
      $('keys-input').focus();
    };
    list.appendChild(chip);

    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = `${p.label} · ${p.envKey}`;
    sel.appendChild(opt);
  }

  if (prevChoice && data.providers.some(p => p.id === prevChoice)) {
    sel.value = prevChoice;
  }
  syncKeysFormSelection();
}

async function openKeysPanel(focusProvider = null) {
  closeActionsMenu();
  try {
    const res = await fetch('/api/keys');
    const data = await res.json();
    renderKeysPanel(data, focusProvider || $('provider').value);
    $('keys-panel').hidden = false;
    $('keys-input').value = '';
    $('keys-input').focus();
  } catch {
    showToast('Could not load key settings from local server.', true);
  }
}

function closeKeysPanel() {
  $('keys-panel').hidden = true;
  $('keys-input').value = '';
}

$('keys-btn').onclick = () => {
  if ($('keys-panel').hidden) openKeysPanel();
  else closeKeysPanel();
};
$('keys-close').onclick = closeKeysPanel;
$('keys-provider').onchange = syncKeysFormSelection;

$('keys-form').onsubmit = async e => {
  e.preventDefault();
  const provider = $('keys-provider').value;
  const apiKey = $('keys-input').value.trim();
  if (!apiKey) {
    showToast('Paste an API key first.', true);
    return;
  }
  try {
    const res = await fetch('/api/keys', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider, apiKey }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Could not save key');
    $('keys-input').value = '';
    renderKeysPanel(data, provider);
    showToast('API key saved locally to ~/.config/word-poc/.env');
  } catch (err) {
    showToast(err.message || String(err), true);
  }
};

$('keys-remove').onclick = async () => {
  const provider = $('keys-provider').value;
  try {
    const res = await fetch('/api/keys', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider, apiKey: '' }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Could not remove key');
    renderKeysPanel(data, provider);
    showToast('Removed key from ~/.config/word-poc/.env');
  } catch (err) {
    showToast(err.message || String(err), true);
  }
};

document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && !$('keys-panel').hidden) closeKeysPanel();
});

/* ---------- Boot ---------- */

function boot(inWord) {
  state.inWord = inWord;
  loadProviders();
  updateContextUI();
  renderEmptyState();
  updateSendState();
  if (inWord) {
    watchSelection();
  } else {
    showToast('Preview mode (open inside Word for live document editing).', false, 4000);
  }
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(info => {
    boot(info && info.host === Office.HostType.Word);
  });
} else {
  boot(false);
}
