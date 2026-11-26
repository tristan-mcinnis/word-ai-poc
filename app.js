/* global Office */
import {
  parseAssistantReply,
  actionLabel,
  primaryButtonLabel,
  isDestructiveAction,
  renderMarkdown,
  renderDiffHtml,
  escapeHtml,
} from './lib/format.js';
import { parseCellGrid, parseFormatSpec } from './lib/excel-utils.js';
import { parseSlideSpec, parseSlideRef } from './lib/ppt-utils.js';
import { detectOfficeHost, createHostAdapter } from './lib/hosts/index.js';

const $ = id => document.getElementById(id);
const CONTEXT_MODES = ['auto', 'selection', 'document'];

const state = {
  inOffice: false,
  hostId: 'word',
  adapter: createHostAdapter('word'),
  messages: [],          // { role, content } sent to the model
  selectionText: '',
  documentText: '',
  comments: [],
  hostContext: {},       // host-specific structured metadata (sheets/errors for Excel, slides/layouts for PPT)
  contextMode: 'auto',   // 'auto' | 'selection' | 'document'
  trackChanges: false,   // Word native Track Changes
  highlightEdits: false, // Excel yellow tint on AI-edited cells
  autoApply: false,      // false = Ask before edits, true = Always allow (non-destructive)
  controller: null,
  readTimer: null,
  reading: false,
  toastTimer: null,
  menuOpen: false,
  menuIndex: 0,
  filteredActions: [],
  userScrolledUp: false,
  keyInfo: null,
  skillsInfo: null,
  pinnedSkills: [],
  lastPrompt: '',
};

/* ---------- Utilities ---------- */

async function copyTextSafe(text) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {}
  // Fallback for sandboxed WKWebView inside Office on macOS
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    ta.remove();
    return ok;
  } catch {
    return false;
  }
}

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

/* ---------- Host UI & Context reading ---------- */

function updateHostChrome() {
  const badge = $('host-badge');
  const previewSelect = $('preview-host');
  if (state.inOffice) {
    badge.textContent = state.adapter.label;
    badge.className = `host-badge host-${state.hostId}`;
    badge.hidden = false;
    previewSelect.hidden = true;
  } else {
    badge.hidden = true;
    previewSelect.value = state.hostId;
    previewSelect.hidden = false;
  }
  $('input').placeholder = state.adapter.placeholder;
  state.filteredActions = state.adapter.quickActions;
  updateToolbarToggles();
  updateContextUI();
}

function updateToolbarToggles() {
  const secBtn = $('secondary-btn');
  const toggleCfg = state.adapter.secondaryToggle;
  if (toggleCfg) {
    const isOn = Boolean(state[toggleCfg.key]);
    secBtn.hidden = false;
    secBtn.classList.toggle('active', isOn);
    secBtn.querySelector('span').textContent = toggleCfg.label;
    secBtn.title = isOn ? toggleCfg.titleOn : toggleCfg.titleOff;
  } else {
    secBtn.hidden = true;
  }

  const modeBtn = $('mode-btn');
  modeBtn.classList.toggle('active', state.autoApply);
  modeBtn.querySelector('span').textContent = state.autoApply ? 'Auto' : 'Ask';
  modeBtn.title = state.autoApply
    ? `Auto-apply ON: non-destructive edits apply directly to ${state.adapter.label}`
    : 'Ask before edits: review and click Allow on each edit';
}

function updateContextUI() {
  const info = state.adapter.formatContextBadge(
    {
      selectionText: state.selectionText,
      documentText: state.documentText,
      comments: state.comments,
      hostContext: state.hostContext,
    },
    state.contextMode,
  );
  const badge = $('context-badge');
  const preview = $('context-preview');
  badge.textContent = info.badge;
  badge.classList.toggle('has-selection', Boolean(info.hasSelection));
  preview.textContent = info.preview;

  if (state.messages.length === 0) renderEmptyState();
}

async function readHostContext() {
  if (!state.inOffice || state.reading) return;
  state.reading = true;
  try {
    const snapshot = await state.adapter.readContext();
    if (!snapshot) return;

    const prevSig = `${state.selectionText.length}|${state.documentText.length}|${state.comments.length}|${JSON.stringify(state.hostContext)}`;
    state.selectionText = snapshot.selectionText || '';
    state.documentText = snapshot.documentText || '';
    state.comments = snapshot.comments || [];
    state.hostContext = snapshot.hostContext || {};

    if (typeof snapshot.trackChanges === 'boolean' && snapshot.trackChanges !== state.trackChanges) {
      state.trackChanges = snapshot.trackChanges;
      updateToolbarToggles();
    }

    const nextSig = `${state.selectionText.length}|${state.documentText.length}|${state.comments.length}|${JSON.stringify(state.hostContext)}`;
    if (prevSig !== nextSig) updateContextUI();
  } catch {
    // Office selection changes can fire mid-sync; the next event catches up.
  } finally {
    state.reading = false;
  }
}

function scheduleRead() {
  clearTimeout(state.readTimer);
  state.readTimer = setTimeout(readHostContext, 250);
}

$('refresh').onclick = async () => {
  await readHostContext();
  showToast(`${state.adapter.label} context refreshed.`, false, 1800);
};

$('context-pill').onclick = () => {
  const idx = CONTEXT_MODES.indexOf(state.contextMode);
  state.contextMode = CONTEXT_MODES[(idx + 1) % CONTEXT_MODES.length];
  updateContextUI();
  showToast(state.adapter.modeLabels[state.contextMode] || `Context: ${state.contextMode}`, false, 2000);
};

$('secondary-btn').onclick = async () => {
  const toggleCfg = state.adapter.secondaryToggle;
  if (!toggleCfg || !state.adapter.toggleSecondary) return;
  const next = !state[toggleCfg.key];
  try {
    const res = await state.adapter.toggleSecondary(next);
    state[toggleCfg.key] = Boolean(res.value);
    updateToolbarToggles();
    if (res.message) showToast(res.message, false, 2200);
  } catch (e) {
    showToast(e.message || String(e), true);
  }
};

$('mode-btn').onclick = () => {
  state.autoApply = !state.autoApply;
  updateToolbarToggles();
  showToast(
    state.autoApply ? `Auto-apply ON · edits apply directly to ${state.adapter.label}` : 'Ask before edits ON',
    false,
    2000,
  );
};

$('preview-host').onchange = () => {
  if (state.inOffice) return;
  const nextHost = $('preview-host').value;
  state.hostId = nextHost;
  state.adapter = createHostAdapter(nextHost);
  state.selectionText = '';
  state.documentText = '';
  state.comments = [];
  state.hostContext = {};
  closeActionsMenu();
  updateHostChrome();
  loadSkills();
  if (state.messages.length === 0) renderEmptyState();
  showToast(`Previewing ${state.adapter.label} mode`, false, 1800);
};

/* ---------- Smart auto-scroll & citation navigation ---------- */

const transcriptEl = $('transcript');
transcriptEl.addEventListener('scroll', () => {
  const distFromBottom = transcriptEl.scrollHeight - transcriptEl.scrollTop - transcriptEl.clientHeight;
  state.userScrolledUp = distFromBottom > 60;
});

transcriptEl.addEventListener('click', async e => {
  const chip = e.target.closest('.ref-chip');
  if (!chip) return;
  const ref = chip.dataset.ref;
  if (!ref || !state.adapter.navigateToRef) return;
  try {
    const res = await state.adapter.navigateToRef(ref);
    if (res?.message) showToast(res.message, !res.ok, 2000);
    if (res?.ok) scheduleRead();
  } catch (err) {
    showToast(err.message || String(err), true);
  }
});

function scrollToBottom(force = false) {
  if (force || !state.userScrolledUp) {
    transcriptEl.scrollTop = transcriptEl.scrollHeight;
  }
}

/* ---------- Empty state ---------- */

function renderEmptyState() {
  if (state.messages.length > 0) return;
  const { heading, subtitle, starters } = state.adapter.getStarters({
    selectionText: state.selectionText,
    documentText: state.documentText,
    comments: state.comments,
    hostContext: state.hostContext,
  });

  transcriptEl.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'empty-state';
  wrap.innerHTML = `<div class="empty-head"><h2></h2><p></p></div><div class="starter-grid"></div>`;
  wrap.querySelector('h2').textContent = heading;
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

/* ---------- Applying edits via active host adapter ---------- */

async function applyEdit(edit, expectedSelection = '', overrideAction = null) {
  if (!state.inOffice) {
    showToast(`Open inside Microsoft ${state.adapter.label} to apply edits directly.`, true);
    return false;
  }

  try {
    const res = await state.adapter.applyEdit(edit, {
      expectedSelection,
      currentSelection: state.selectionText,
      overrideAction,
      trackChanges: state.trackChanges,
      highlightEdits: state.highlightEdits,
    });
    if (res?.ok) {
      showToast(res.message || 'Applied');
      scheduleRead();
      return true;
    }
    showToast(res?.message || 'Could not apply edit.', true);
    return false;
  } catch (e) {
    showToast(e.message || String(e), true);
    return false;
  }
}

/* ---------- Rendering conversation turns & rich edit cards ---------- */

function addUserTurn(text, attachedSelection, hostContextSnapshot) {
  if (state.messages.length === 0) transcriptEl.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'turn user';

  if (state.contextMode !== 'document') {
    let tagLabel = '';
    if (state.hostId === 'excel' && hostContextSnapshot?.selectionAddress) {
      tagLabel = `Range: ${hostContextSnapshot.selectionAddress}`;
    } else if (state.hostId === 'powerpoint' && hostContextSnapshot?.currentSlideIndex) {
      tagLabel = hostContextSnapshot.selectedShapeName
        ? `Slide ${hostContextSnapshot.currentSlideIndex} · ${hostContextSnapshot.selectedShapeName}`
        : `Slide ${hostContextSnapshot.currentSlideIndex}`;
    } else {
      const cleanSel = (attachedSelection || '').replace(/\s+/g, ' ').trim();
      if (cleanSel) {
        tagLabel = `Selection: "${cleanSel.slice(0, 48)}${cleanSel.length > 48 ? '…' : ''}"`;
      }
    }
    if (tagLabel) {
      const tag = document.createElement('div');
      tag.className = 'user-context-tag';
      tag.textContent = tagLabel;
      tag.title = attachedSelection || tagLabel;
      wrap.appendChild(tag);
    }
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

function lookupShapeBaseline(edit, hostContextSnapshot) {
  const slides = hostContextSnapshot?.slides;
  const attrs = edit.attrs || {};
  if (Array.isArray(slides) && slides.length) {
    const slideNum = parseSlideRef(attrs.slide) || hostContextSnapshot.currentSlideIndex || 1;
    const slide = slides.find(s => s.index === slideNum);
    if (slide && Array.isArray(slide.shapes)) {
      if (attrs.shape) {
        const q = attrs.shape.trim().toLowerCase();
        const found = slide.shapes.find(sh =>
          String(sh.id).toLowerCase() === q ||
          String(sh.name || '').trim().toLowerCase() === q ||
          String(sh.name || '').toLowerCase().includes(q),
        );
        if (found?.text) return found.text;
      }
    }
  }
  return hostContextSnapshot?.selectedShapeText || '';
}

function renderGridPreviewHtml(grid) {
  const maxRows = 8;
  const maxCols = 6;
  const rows = grid.slice(0, maxRows);
  const trs = rows.map(row => {
    const cols = row.slice(0, maxCols);
    const tds = cols.map(cell => {
      const str = String(cell ?? '');
      const cls = str.startsWith('=') ? ' class="is-formula"' : '';
      return `<td${cls} title="${escapeHtml(str)}">${escapeHtml(str)}</td>`;
    }).join('');
    return `<tr>${tds}</tr>`;
  }).join('');
  const truncated = grid.length > maxRows || (grid[0]?.length || 0) > maxCols
    ? `<div class="excel-preview-note">Showing ${rows.length}×${Math.min(maxCols, grid[0].length)} of ${grid.length}×${grid[0].length} cells</div>`
    : '';
  return `<div class="excel-preview-wrap"><table class="excel-preview-table"><tbody>${trs}</tbody></table>${truncated}</div>`;
}

function renderStructuredEditBody(bodyEl, edit) {
  const action = edit.action;
  const attrs = edit.attrs || {};

  if (action === 'set_cells' || action === 'add_sheet' || action === 'add_table') {
    const grid = parseCellGrid(edit.content);
    if (grid.length > 0 && grid[0].length > 0) {
      bodyEl.innerHTML = renderGridPreviewHtml(grid);
      return true;
    }
    if (action === 'add_sheet') {
      bodyEl.textContent = `Creates empty worksheet "${attrs.name || 'Summary'}"`;
      return true;
    }
  }

  if (action === 'set_formula') {
    const formula = (edit.content || '').trim();
    const pills = attrs.number_format ? `<div class="excel-preview-note">Format: ${escapeHtml(attrs.number_format)}</div>` : '';
    bodyEl.innerHTML = `<div class="formula-preview">${escapeHtml(formula)}</div>${pills}`;
    return true;
  }

  if (action === 'format_cells') {
    const spec = parseFormatSpec(edit.content);
    const entries = Object.entries(spec);
    if (entries.length) {
      bodyEl.innerHTML = `<div class="spec-pills">${entries.map(([k, v]) => `<span class="spec-pill">${escapeHtml(k)}: ${escapeHtml(String(v))}</span>`).join('')}</div>`;
      return true;
    }
  }

  if (action === 'add_chart') {
    const pills = [
      `Type: ${attrs.chart_type || 'ColumnClustered'}`,
      attrs.range ? `Data: ${attrs.range}` : 'Data: selection',
      attrs.title ? `Title: ${attrs.title}` : '',
    ].filter(Boolean);
    bodyEl.innerHTML = `<div class="spec-pills">${pills.map(p => `<span class="spec-pill">${escapeHtml(p)}</span>`).join('')}</div>`;
    return true;
  }

  if (action === 'sort_range') {
    const pills = [
      attrs.range ? `Range: ${attrs.range}` : 'Selection',
      `Column: ${attrs.column ?? '0'}`,
      `Order: ${String(attrs.order || 'asc').toUpperCase()}`,
    ];
    bodyEl.innerHTML = `<div class="spec-pills">${pills.map(p => `<span class="spec-pill">${escapeHtml(p)}</span>`).join('')}</div>`;
    return true;
  }

  if (action === 'clear_range') {
    bodyEl.textContent = `Clears cell contents in ${attrs.range || attrs.cell || 'selected range'}.`;
    return true;
  }

  if (action === 'add_slide') {
    const spec = parseSlideSpec(edit.content);
    const pieces = [];
    if (spec.title) pieces.push(`<div class="slide-preview-title">${escapeHtml(spec.title)}</div>`);
    if (spec.subtitle) pieces.push(`<div class="slide-preview-sub">${escapeHtml(spec.subtitle)}</div>`);
    if (spec.bullets.length) {
      pieces.push(`<ul class="slide-preview-bullets">${spec.bullets.map(b => `<li>${escapeHtml(b)}</li>`).join('')}</ul>`);
    } else if (spec.body) {
      pieces.push(`<div class="slide-preview-sub">${escapeHtml(spec.body)}</div>`);
    }
    if (pieces.length) {
      bodyEl.innerHTML = `<div class="slide-preview">${pieces.join('')}</div>`;
      return true;
    }
  }

  if (action === 'delete_slide') {
    bodyEl.textContent = `Permanently deletes Slide ${attrs.slide || '(current)'} from the presentation (Cmd+Z to undo).`;
    return true;
  }

  return false;
}

function buildEditCard(edit, index, turn, expectedSelection, hostContextSnapshot) {
  const card = document.createElement('div');
  const destructive = isDestructiveAction(edit.action);
  card.className = 'edit-card' + (destructive ? ' danger' : '');
  if (turn.appliedSet.has(index)) card.classList.add('applied');

  const head = document.createElement('div');
  head.className = 'edit-card-head';

  const badge = document.createElement('span');
  badge.className = 'edit-badge';
  badge.textContent = actionLabel(edit.action, edit.label, edit.attrs);
  head.appendChild(badge);

  let baselineText = '';
  if (edit.action === 'replace_match' || edit.action === 'replace_slide_text') {
    baselineText = edit.find;
  } else if (edit.action === 'replace_selection') {
    baselineText = expectedSelection;
  } else if (edit.action === 'update_shape') {
    baselineText = lookupShapeBaseline(edit, hostContextSnapshot);
  }
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

  if ((edit.action === 'replace_match' || edit.action === 'replace_slide_text') && edit.find && !turn.diffStates.get(index)) {
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
    } else if (!edit.streaming && renderStructuredEditBody(body, edit)) {
      // Rendered rich Excel/PowerPoint preview
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

    let alwaysBtn = null;

    const primary = document.createElement('button');
    primary.type = 'button';
    primary.className = 'btn-primary' + (destructive ? ' btn-danger' : '') + (turn.appliedSet.has(index) ? ' done' : '');
    primary.textContent = turn.appliedSet.has(index) ? 'Applied ✓' : primaryButtonLabel(edit.action);
    primary.onclick = async () => {
      const ok = await applyEdit(edit, expectedSelection);
      if (ok) {
        turn.appliedSet.add(index);
        card.classList.add('applied');
        primary.classList.remove('btn-danger');
        primary.classList.add('done');
        primary.textContent = 'Applied ✓';
        if (alwaysBtn) alwaysBtn.remove();
      }
    };
    foot.appendChild(primary);

    if (!state.autoApply && !destructive && !turn.appliedSet.has(index)) {
      alwaysBtn = document.createElement('button');
      alwaysBtn.type = 'button';
      alwaysBtn.className = 'btn-secondary';
      alwaysBtn.textContent = 'Always allow';
      alwaysBtn.onclick = async () => {
        state.autoApply = true;
        updateToolbarToggles();
        const ok = await applyEdit(edit, expectedSelection);
        if (ok) {
          turn.appliedSet.add(index);
          card.classList.add('applied');
          primary.classList.add('done');
          primary.textContent = 'Applied ✓';
          alwaysBtn.remove();
        }
      };
      foot.appendChild(alwaysBtn);
    }

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
          if (alwaysBtn) alwaysBtn.remove();
        }
      };
      foot.appendChild(insertBtn);
    }

    if (edit.content && edit.content.trim()) {
      const copyBtn = document.createElement('button');
      copyBtn.type = 'button';
      copyBtn.className = 'btn-secondary';
      copyBtn.textContent = 'Copy';
      copyBtn.onclick = async () => {
        const ok = await copyTextSafe(edit.content);
        showToast(ok ? 'Copied to clipboard.' : 'Copy failed.', !ok, 1800);
      };
      foot.appendChild(copyBtn);
    }

    card.appendChild(foot);
  }

  return card;
}

function renderAssistantTurn(turn, text, expectedSelection, hostContextSnapshot, isFinal = false, metaLabel = '', canRetry = false) {
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
      prose.innerHTML = renderMarkdown(part.text, { host: state.hostId });
      turn.body.appendChild(prose);
    } else if (part.kind === 'edit') {
      const idx = editIndex++;
      edits.push({ edit: part, idx });
      turn.body.appendChild(buildEditCard(part, idx, turn, expectedSelection, hostContextSnapshot));
    }
  }

  const safeEdits = edits.filter(({ edit }) => !isDestructiveAction(edit.action));
  if (isFinal && safeEdits.length >= 2) {
    const remaining = safeEdits.filter(({ idx }) => !turn.appliedSet.has(idx));
    if (remaining.length > 0) {
      const bar = document.createElement('div');
      bar.className = 'apply-all-bar';
      const allBtn = document.createElement('button');
      allBtn.type = 'button';
      allBtn.className = 'btn-primary';
      allBtn.textContent = `Apply all (${remaining.length})`;
      allBtn.onclick = async () => {
        let appliedCount = 0;
        let failedCount = 0;
        for (const { edit, idx } of remaining) {
          if (turn.appliedSet.has(idx)) continue;
          const ok = await applyEdit(edit, expectedSelection);
          if (ok) {
            turn.appliedSet.add(idx);
            appliedCount++;
          } else {
            failedCount++;
          }
        }
        renderAssistantTurn(turn, text, expectedSelection, hostContextSnapshot, true, metaLabel);
        if (appliedCount > 0 && failedCount === 0) {
          showToast(`Applied ${appliedCount} edit${appliedCount === 1 ? '' : 's'} · Cmd+Z to undo`);
        } else if (appliedCount > 0 && failedCount > 0) {
          showToast(`Applied ${appliedCount}, ${failedCount} failed — check individual cards.`, true);
        }
      };
      bar.appendChild(allBtn);
      turn.body.appendChild(bar);
    }
  }

  if (isFinal) {
    turn.meta.hidden = false;
    turn.meta.innerHTML = '';
    const left = document.createElement('span');
    left.textContent = metaLabel;
    const actions = document.createElement('div');
    actions.className = 'meta-actions';

    if (canRetry && state.lastPrompt) {
      const retryBtn = document.createElement('button');
      retryBtn.type = 'button';
      retryBtn.className = 'meta-btn';
      retryBtn.textContent = 'Retry';
      retryBtn.onclick = () => sendTurn(state.lastPrompt);
      actions.appendChild(retryBtn);
    }

    const copyAll = document.createElement('button');
    copyAll.type = 'button';
    copyAll.className = 'meta-btn';
    copyAll.textContent = 'Copy reply';
    copyAll.onclick = async () => {
      const plain = text.replace(/<\/?(?:edit[^>]*|replacement|find|content)>/gi, '').trim();
      const ok = await copyTextSafe(plain);
      showToast(ok ? 'Reply copied.' : 'Copy failed.', !ok, 1600);
    };
    actions.appendChild(copyAll);
    turn.meta.append(left, actions);
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
  $('send').disabled = busy || (!inputEl.value.trim() && state.pinnedSkills.length === 0);
}

function combinedSlashActions() {
  const hostActions = (state.adapter.quickActions || []).map(a => ({ ...a, isSkill: false }));
  const globalSkills = (state.skillsInfo?.skills || [])
    .filter(s => s.enabled)
    .map(s => ({
      cmd: s.cmd,
      title: s.description ? `${s.name} — ${s.description.slice(0, 60)}` : s.name,
      skillName: s.name,
      prompt: `Apply the /${s.name} skill to the current ${state.adapter.noun}.`,
      isSkill: true,
    }));
  return [...hostActions, ...globalSkills];
}

function chooseSlashAction(chosen) {
  closeActionsMenu();
  inputEl.value = '';
  autoResizeInput();
  if (chosen.isSkill) {
    pinSkill(chosen.skillName);
    updateSendState();
    inputEl.focus();
    showToast(`Pinned /${chosen.skillName} — type your request or press Enter to run.`, false, 2600);
    return;
  }
  updateSendState();
  sendTurn(chosen.prompt);
}

function renderActionsMenu(filter = '') {
  const actions = combinedSlashActions();
  const q = filter.trim().toLowerCase();
  state.filteredActions = q
    ? actions.filter(a => a.cmd.includes(q) || a.title.toLowerCase().includes(q))
    : actions;

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
    btn.innerHTML = `<span class="action-item-title"></span><span class="action-item-meta">${action.isSkill ? '<span class="action-skill-tag">Skill</span>' : ''}<span class="action-item-cmd"></span></span>`;
    btn.querySelector('.action-item-title').textContent = action.title;
    btn.querySelector('.action-item-cmd').textContent = action.cmd;
    btn.onclick = () => chooseSlashAction(action);
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
      if (chosen) chooseSlashAction(chosen);
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

document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && !$('keys-panel').hidden) {
    closeKeysPanel();
  } else if (e.key === 'Escape' && !$('skills-panel').hidden) {
    closeSkillsPanel();
  } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
    e.preventDefault();
    inputEl.focus();
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
  let text = String(userText || '').trim();
  const pinnedNow = state.pinnedSkills.slice();
  if (!text && pinnedNow.length > 0) {
    text = `Apply ${pinnedNow.map(s => `/${s}`).join(', ')} to the current ${state.adapter.noun}.`;
  }
  if (!text || state.controller) return;

  closeActionsMenu();
  inputEl.value = '';
  autoResizeInput();
  state.lastPrompt = text;

  // Pull fresh host context right before sending so the model sees recent edits
  await readHostContext();

  const usedSelection = state.selectionText;
  const usedDoc = state.documentText;
  const usedHostContext = state.hostContext;

  addUserTurn(text, usedSelection, usedHostContext);
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
        host: state.hostId,
        messages: state.messages,
        selection: usedSelection,
        documentText: usedDoc,
        contextMode: state.contextMode,
        comments: state.comments,
        hostContext: usedHostContext,
        activeSkills: pinnedNow,
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
          renderAssistantTurn(turn, reply, usedSelection, usedHostContext, false);
        }
        if (payload.error) throw new Error(payload.error);
        if (payload.done) {
          const baseMeta = payload.providerLabel ? `${payload.providerLabel} · ${payload.model}` : payload.model;
          const skillTag = Array.isArray(payload.activeSkills) && payload.activeSkills.length
            ? ` · ${payload.activeSkills.map(s => `/${s}`).join(' ')}`
            : '';
          metaLabel = `${baseMeta}${skillTag}`;
        }
      }
    }

    if (!reply.trim()) throw new Error('The model returned nothing.');
    state.messages.push({ role: 'assistant', content: reply });
    renderAssistantTurn(turn, reply, usedSelection, usedHostContext, true, metaLabel);

    // In Auto-apply mode, apply all non-destructive proposed edits automatically once finished
    if (state.autoApply && state.inOffice) {
      const finalParts = parseAssistantReply(reply).filter(p => p.kind === 'edit');
      let appliedCount = 0;
      for (let idx = 0; idx < finalParts.length; idx++) {
        if (isDestructiveAction(finalParts[idx].action)) continue;
        const ok = await applyEdit(finalParts[idx], usedSelection);
        if (ok) {
          turn.appliedSet.add(idx);
          appliedCount++;
        }
      }
      if (appliedCount > 0) {
        renderAssistantTurn(turn, reply, usedSelection, usedHostContext, true, metaLabel);
      }
    }
  } catch (e) {
    if (e.name === 'AbortError') {
      showToast('Stopped.', false, 1800);
      if (reply.trim()) {
        state.messages.push({ role: 'assistant', content: reply });
        renderAssistantTurn(turn, reply, usedSelection, usedHostContext, true, 'Stopped');
      } else {
        turn.wrap.remove();
        state.messages.pop();
        if (state.messages.length === 0) renderEmptyState();
      }
    } else {
      showToast(e.message || String(e), true);
      if (reply.trim()) {
        renderAssistantTurn(turn, reply, usedSelection, usedHostContext, true, 'Error', true);
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
  closeSkillsPanel();
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

/* ---------- Global AGENTS.md & Skills panel ---------- */

function renderActiveSkillChips() {
  const row = $('active-skills-row');
  row.innerHTML = '';
  if (!state.pinnedSkills.length) {
    row.hidden = true;
    return;
  }
  row.hidden = false;
  for (const name of state.pinnedSkills) {
    const chip = document.createElement('span');
    chip.className = 'active-skill-chip';
    chip.innerHTML = `<span></span><button type="button" aria-label="Remove skill">&times;</button>`;
    chip.querySelector('span').textContent = `/${name}`;
    chip.querySelector('button').onclick = () => unpinSkill(name);
    row.appendChild(chip);
  }
}

function pinSkill(name) {
  if (!name || state.pinnedSkills.includes(name)) return;
  state.pinnedSkills.push(name);
  renderActiveSkillChips();
  updateSkillsButtonBadge();
  updateSendState();
}

function unpinSkill(name) {
  state.pinnedSkills = state.pinnedSkills.filter(s => s !== name);
  renderActiveSkillChips();
  updateSkillsButtonBadge();
  updateSendState();
}

function updateSkillsButtonBadge() {
  const btn = $('skills-btn');
  const label = btn.querySelector('span');
  const enabledCount = (state.skillsInfo?.skills || []).filter(s => s.enabled).length;
  const hasAgents = Boolean(state.skillsInfo?.agentsEnabled && state.skillsInfo?.agentsFiles?.length);
  const pinned = state.pinnedSkills.length;
  btn.classList.toggle('has-active', hasAgents || enabledCount > 0 || pinned > 0);
  label.textContent = pinned > 0 ? `Skills (${pinned})` : 'Skills';
}

function renderSkillsPanel(data) {
  state.skillsInfo = data;
  // Drop any pinned skills that were disabled or removed on disk
  const enabledNames = new Set((data.skills || []).filter(s => s.enabled).map(s => s.name));
  state.pinnedSkills = state.pinnedSkills.filter(n => enabledNames.has(n));
  renderActiveSkillChips();
  updateSkillsButtonBadge();

  const agentsSec = $('agents-section');
  agentsSec.innerHTML = '';
  const agentsTitle = document.createElement('div');
  agentsTitle.className = 'skills-section-title';
  agentsTitle.innerHTML = `<span>Global AGENTS.md</span>`;
  const toggleAgents = document.createElement('button');
  toggleAgents.type = 'button';
  toggleAgents.className = 'skill-mini-btn' + (data.agentsEnabled ? ' active' : '');
  toggleAgents.textContent = data.agentsEnabled ? 'Active' : 'Off';
  toggleAgents.onclick = async () => {
    await updateSkillSettings({ agentsEnabled: !data.agentsEnabled });
  };
  agentsTitle.appendChild(toggleAgents);
  agentsSec.appendChild(agentsTitle);

  if (data.agentsFiles && data.agentsFiles.length > 0) {
    for (const f of data.agentsFiles) {
      const row = document.createElement('div');
      row.className = 'skill-row' + (data.agentsEnabled ? '' : ' disabled');
      row.innerHTML = `<div class="skill-row-main"><div class="skill-row-name"></div><div class="skill-row-source"></div></div>`;
      row.querySelector('.skill-row-name').textContent = f.path;
      row.querySelector('.skill-row-source').textContent = `${f.chars.toLocaleString()} chars loaded into system prompt`;
      agentsSec.appendChild(row);
    }
  } else {
    const empty = document.createElement('div');
    empty.className = 'skills-empty';
    empty.textContent = 'No global AGENTS.md found yet.';
    agentsSec.appendChild(empty);
  }

  const skillsSec = $('skills-section');
  skillsSec.innerHTML = '';
  const skillsTitle = document.createElement('div');
  skillsTitle.className = 'skills-section-title';
  skillsTitle.innerHTML = `<span>Global Skills (${(data.skills || []).length})</span>`;
  skillsSec.appendChild(skillsTitle);

  if (data.skills && data.skills.length > 0) {
    for (const s of data.skills) {
      const isPinned = state.pinnedSkills.includes(s.name);
      const row = document.createElement('div');
      row.className = 'skill-row' + (s.enabled ? '' : ' disabled');
      row.innerHTML = `<div class="skill-row-main"><div class="skill-row-name"></div><div class="skill-row-desc"></div><div class="skill-row-source"></div></div><div class="skill-row-actions"></div>`;
      row.querySelector('.skill-row-name').textContent = s.cmd;
      row.querySelector('.skill-row-desc').textContent = s.description || 'Global skill';
      row.querySelector('.skill-row-source').textContent = s.source;

      const actions = row.querySelector('.skill-row-actions');
      if (s.enabled) {
        const useBtn = document.createElement('button');
        useBtn.type = 'button';
        useBtn.className = 'skill-mini-btn' + (isPinned ? ' active' : '');
        useBtn.textContent = isPinned ? 'Pinned' : 'Use';
        useBtn.onclick = () => {
          if (state.pinnedSkills.includes(s.name)) unpinSkill(s.name);
          else {
            pinSkill(s.name);
            closeSkillsPanel();
            inputEl.focus();
          }
          renderSkillsPanel(state.skillsInfo);
        };
        actions.appendChild(useBtn);
      }

      const enBtn = document.createElement('button');
      enBtn.type = 'button';
      enBtn.className = 'skill-mini-btn';
      enBtn.textContent = s.enabled ? 'Disable' : 'Enable';
      enBtn.onclick = async () => {
        await updateSkillSettings({ toggleSkill: s.name, enabled: !s.enabled });
      };
      actions.appendChild(enBtn);

      skillsSec.appendChild(row);
    }
  } else {
    const empty = document.createElement('div');
    empty.className = 'skills-empty';
    empty.textContent = 'No global SKILL.md folders found yet.';
    skillsSec.appendChild(empty);
  }

  const dirs = (data.searchPaths?.skillsDirs || []).slice(0, 4).join(', ');
  $('skills-paths-note').textContent = `Scans AGENTS.md (~/.agents, ~/.claude, ~/.codex, ~/.config/word-poc) and skills (${dirs}). Invoke any skill with /skill-name.`;
}

async function loadSkills(showFeedback = false) {
  try {
    const res = await fetch(`/api/skills?host=${encodeURIComponent(state.hostId)}`);
    if (!res.ok) return;
    const data = await res.json();
    renderSkillsPanel(data);
    if (showFeedback) {
      showToast(`Loaded ${data.agentsFiles.length} AGENTS.md & ${data.skills.length} global skill${data.skills.length === 1 ? '' : 's'}.`, false, 2000);
    }
  } catch {
    // Non-fatal if server not reachable yet
  }
}

async function updateSkillSettings(patch) {
  try {
    const res = await fetch('/api/skills', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...patch, host: state.hostId }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Could not update skills');
    renderSkillsPanel(data);
  } catch (err) {
    showToast(err.message || String(err), true);
  }
}

async function openSkillsPanel() {
  closeActionsMenu();
  closeKeysPanel();
  await loadSkills();
  $('skills-panel').hidden = false;
}

function closeSkillsPanel() {
  $('skills-panel').hidden = true;
}

$('skills-btn').onclick = () => {
  if ($('skills-panel').hidden) openSkillsPanel();
  else closeSkillsPanel();
};
$('skills-close').onclick = closeSkillsPanel;
$('skills-refresh').onclick = () => loadSkills(true);

/* ---------- Boot ---------- */

function boot(detectedHost) {
  state.inOffice = Boolean(detectedHost);
  state.hostId = detectedHost || 'word';
  state.adapter = createHostAdapter(state.hostId);

  loadProviders();
  loadSkills();
  updateHostChrome();
  renderEmptyState();
  updateSendState();

  if (state.inOffice) {
    state.adapter.watchSelection(scheduleRead, msg => showToast(msg));
    readHostContext();
  } else {
    showToast('Preview mode (open inside Word, Excel, or PowerPoint for live editing).', false, 4000);
  }
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(info => {
    boot(detectOfficeHost(info, Office));
  });
} else {
  boot(null);
}
