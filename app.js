/* global Office, Word */
const $ = id => document.getElementById(id);

let range = null;          // tracked range in the document
let originalText = '';     // what that range said when captured
let lastResult = '';       // streamed text
let controller = null;     // AbortController for the run
let mode = 'rewrite';

const setStatus = (text, isError = false) => {
  $('status').textContent = text;
  $('status').classList.toggle('err', isError);
};

const refreshButtons = () => {
  const hasSelection = Boolean(range) && !controller;
  $('run').disabled = !hasSelection;
  $('stop').disabled = !controller;
  const hasResult = Boolean(lastResult.trim()) && !controller;
  $('apply').disabled = !(hasResult && range);
  $('insert').disabled = !(hasResult && range);
};

function showOriginal(text) {
  const el = $('original');
  el.textContent = text || 'Nothing captured. Select a paragraph in the document, then press Read selection.';
  el.classList.toggle('empty', !text);
}

function showPreview(text, thinking = false) {
  const el = $('preview');
  el.textContent = text || (thinking ? 'Thinking…' : 'No result yet.');
  el.classList.toggle('thinking', thinking);
}

/* ---------- Word side ---------- */

$('read').onclick = async () => {
  try {
    await Word.run(async ctx => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      if (!sel.text.trim()) throw new Error('Select some text in the document first.');
      if (sel.text.length > 12000) throw new Error('Selection is over 12,000 characters. Select less.');
      sel.track();
      await ctx.sync();
      range = sel;
      originalText = sel.text;
    });
    showOriginal(originalText);
    lastResult = '';
    showPreview('');
    setStatus(`Captured ${originalText.length} characters. Pick an action, then Run.`);
  } catch (e) {
    setStatus(e.message || String(e), true);
  }
  refreshButtons();
};

$('apply').onclick = async () => {
  const text = lastResult;
  if (!text.trim() || !range) return;
  try {
    await Word.run(range, async ctx => {
      range.load('text');
      await ctx.sync();
      if (range.text !== originalText) throw new Error('The document text changed since you read it. Read the selection again.');
      range.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
      range.untrack();
      await ctx.sync();
    });
    range = null;
    setStatus('Replaced. Word Undo (Cmd+Z) puts the original back.');
  } catch (e) {
    setStatus(e.message || String(e), true);
  }
  refreshButtons();
};

$('insert').onclick = async () => {
  const text = lastResult;
  if (!text.trim() || !range) return;
  try {
    await Word.run(range, async ctx => {
      range.insertText('\n' + text, Word.InsertLocation.after);
      await ctx.sync();
      range.untrack();
      await ctx.sync();
    });
    range = null;
    setStatus('Inserted below the selection. The original is untouched.');
  } catch (e) {
    setStatus(e.message || String(e), true);
  }
  refreshButtons();
};

/* ---------- settings ---------- */

async function loadProviders() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    const sel = $('provider');
    sel.innerHTML = '';
    for (const p of data.providers) {
      const opt = document.createElement('option');
      opt.value = p.id;
      opt.textContent = `${p.label} · ${p.model}`;
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
      if (cfg.error) setStatus(cfg.error, true);
      else setStatus(`Model: ${cfg.active.provider} · ${cfg.active.model}`);
    };
  } catch (e) {
    setStatus('Cannot reach the local pane server. Is the login agent running?', true);
  }
}

/* ---------- run ---------- */

document.querySelectorAll('[data-mode]').forEach(btn => {
  btn.onclick = () => {
    mode = btn.dataset.mode;
    $('instruction').value = '';
    document.querySelectorAll('[data-mode]').forEach(b => b.classList.toggle('primary', b === btn));
    if (range) $('run').focus();
  };
});

$('run').onclick = async () => {
  if (!range) return;
  controller = new AbortController();
  lastResult = '';
  showPreview('', true);
  refreshButtons();
  setStatus('Asking the model…');
  try {
    const res = await fetch('/api/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        selection: originalText,
        mode,
        instruction: $('instruction').value,
        language: $('language').value,
      }),
      signal: controller.signal,
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: `Server returned ${res.status}` }));
      throw new Error(err.error || `Server returned ${res.status}`);
    }
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';
    let done = false;
    while (!done) {
      const { value, done: end } = await reader.read();
      if (end) break;
      buf += decoder.decode(value, { stream: true });
      let i;
      while ((i = buf.indexOf('\n\n')) >= 0) {
        const frame = buf.slice(0, i);
        buf = buf.slice(i + 2);
        if (!frame.startsWith('data:')) continue;
        const payload = JSON.parse(frame.slice(5).trim());
        if (payload.delta) {
          lastResult += payload.delta;
          showPreview(lastResult);
        }
        if (payload.error) throw new Error(payload.error);
        if (payload.done) {
          done = true;
          setStatus(`Done · ${payload.provider} · ${payload.model}. Review before replacing.`);
        }
      }
    }
    if (!lastResult.trim()) throw new Error('The model returned nothing.');
  } catch (e) {
    if (e.name === 'AbortError') setStatus('Stopped. Nothing was changed in the document.');
    else setStatus(e.message || String(e), true);
  } finally {
    controller = null;
    refreshButtons();
  }
};

$('stop').onclick = () => { if (controller) controller.abort(); };

/* ---------- start ---------- */

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    setStatus('Open this pane inside Word.', true);
    return;
  }
  setStatus('Ready. Select a paragraph, then Read selection.');
  loadProviders();
  refreshButtons();
});
