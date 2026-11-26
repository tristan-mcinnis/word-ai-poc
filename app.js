/* global Office, Word */
const $ = id => document.getElementById(id);

const state = {
  messages: [],        // {role, content} sent to the model
  contextText: '',     // latest selection text, attached to every turn
  controller: null,
  readTimer: null,
  reading: false,
};

const setStatus = (text, isError = false) => {
  $('status').textContent = text;
  $('status').classList.toggle('err', isError);
};

const refreshControls = () => {
  const busy = Boolean(state.controller);
  $('send').disabled = busy;
  $('stop').hidden = !busy;
  $('input').disabled = false;
};

/* ---------- selection tracking ---------- */

function showContext() {
  const el = $('context');
  const text = state.contextText.replace(/\s+/g, ' ').trim();
  if (!text) {
    el.textContent = 'no selection';
    el.classList.add('none');
    return;
  }
  el.classList.remove('none');
  el.textContent = `${text.length} chars · ${text.slice(0, 90)}${text.length > 90 ? '…' : ''}`;
  el.title = state.contextText;
}

async function readSelection() {
  if (state.reading) return;
  state.reading = true;
  try {
    await Word.run(async ctx => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      if (sel.text !== state.contextText) {
        state.contextText = sel.text;
        showContext();
      }
    });
  } catch {
    // A selection change can land mid-sync; the next event fixes it.
  } finally {
    state.reading = false;
  }
}

function scheduleRead() {
  clearTimeout(state.readTimer);
  state.readTimer = setTimeout(readSelection, 400);
}

function watchSelection() {
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRead, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus('Selection changes are not reported here. Use Refresh after selecting.');
      }
    });
    setStatus('Ready. Select text in the document, it attaches itself.');
  } catch {
    setStatus('Selection changes are not reported here. Use Refresh after selecting.');
  }
  readSelection();
}

$('refresh').onclick = () => readSelection();

/* ---------- rendering ---------- */

function parseReply(text) {
  const parts = [];
  const re = /<replacement>([\s\S]*?)<\/replacement>/gi;
  let last = 0;
  let match;
  while ((match = re.exec(text)) !== null) {
    if (match.index > last) parts.push({ kind: 'prose', text: text.slice(last, match.index) });
    parts.push({ kind: 'replacement', text: match[1].replace(/^\n+|\n+$/g, '') });
    last = re.lastIndex;
  }
  const tail = text.slice(last);
  if (tail.trim()) parts.push({ kind: 'prose', text: tail });
  if (!parts.length) parts.push({ kind: 'prose', text: text });
  return parts;
}

function addTurn(role, text = '') {
  const wrap = document.createElement('div');
  wrap.className = `turn ${role}`;
  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  wrap.appendChild(bubble);
  $('transcript').appendChild(wrap);
  $('transcript').scrollTop = $('transcript').scrollHeight;
  return { wrap, bubble, role };
}

function renderAssistant(turn, text, expectedContext = '') {
  const bubble = turn.bubble;
  bubble.innerHTML = '';
  bubble.classList.toggle('empty', !text.trim());
  if (!text.trim()) { bubble.textContent = 'Thinking…'; return; }
  for (const part of parseReply(text)) {
    if (part.kind === 'prose') {
      const p = document.createElement('div');
      p.textContent = part.text.trim();
      bubble.appendChild(p);
    } else {
      const box = document.createElement('div');
      box.className = 'repl';
      box.textContent = part.text;
      bubble.appendChild(box);
      const acts = document.createElement('div');
      acts.className = 'acts';
      const replace = document.createElement('button');
      replace.textContent = 'Replace selection';
      replace.onclick = () => applyText(part.text, 'replace', expectedContext);
      const insert = document.createElement('button');
      insert.textContent = 'Insert below';
      insert.onclick = () => applyText(part.text, 'insert', expectedContext);
      const copy = document.createElement('button');
      copy.textContent = 'Copy';
      copy.onclick = async () => {
        try { await navigator.clipboard.writeText(part.text); setStatus('Copied.'); }
        catch { setStatus('Copy failed.', true); }
      };
      acts.append(replace, insert, copy);
      bubble.appendChild(acts);
    }
  }
  $('transcript').scrollTop = $('transcript').scrollHeight;
}

/* ---------- writing back to Word ---------- */

async function applyText(text, how, expectedContext = '') {
  if (!text.trim()) return;
  try {
    await Word.run(async ctx => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      if (!sel.text.trim()) throw new Error('Nothing is selected. Select the passage again.');
      // The answer was written for the passage attached to its turn. Refuse if that is no longer what is selected.
      if (how === 'replace' && expectedContext.trim() && sel.text !== expectedContext) {
        throw new Error('The selection changed since this answer. Select the passage again, or use Insert below.');
      }
      if (how === 'replace') {
        sel.insertText(text, Word.InsertLocation.replace);
      } else {
        sel.insertText('\n' + text, Word.InsertLocation.after);
      }
      await ctx.sync();
    });
    setStatus(how === 'replace' ? 'Replaced. Cmd+Z puts the original back.' : 'Inserted below the selection.');
  } catch (e) {
    setStatus(e.message || String(e), true);
  }
}

/* ---------- model ---------- */

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
      setStatus(cfg.error ? cfg.error : `Model: ${cfg.active.provider} · ${cfg.active.model}`, Boolean(cfg.error));
    };
  } catch {
    setStatus('Cannot reach the local pane server. Is the login agent running?', true);
  }
}

function pushMessage(role, content) {
  state.messages.push({ role, content });
}

async function sendTurn(userText) {
  const text = userText.trim();
  if (!text || state.controller) return;
  $('input').value = '';
  addTurn('user', text).bubble.textContent = text;
  pushMessage('user', text);

  const turn = addTurn('assistant', '');
  const usedContext = state.contextText;
  let reply = '';
  const controller = new AbortController();
  state.controller = controller;
  refreshControls();
  setStatus('Working…');

  try {
    const res = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: state.messages, selection: state.contextText }),
      signal: controller.signal,
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: `Server returned ${res.status}` }));
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
          renderAssistant(turn, reply, usedContext);
        }
        if (payload.error) throw new Error(payload.error);
        if (payload.done) setStatus(`${payload.provider} · ${payload.model}`);
      }
    }
    if (!reply.trim()) throw new Error('The model returned nothing.');
    pushMessage('assistant', reply);
    renderAssistant(turn, reply, usedContext);
  } catch (e) {
    if (e.name === 'AbortError') {
      setStatus('Stopped. The document was not touched.');
      if (reply.trim()) pushMessage('assistant', reply);
    } else {
      setStatus(e.message || String(e), true);
      renderAssistant(turn, reply || '(no answer)', usedContext);
    }
  } finally {
    state.controller = null;
    refreshControls();
  }
}

$('send').onclick = () => sendTurn($('input').value);
$('input').onkeydown = e => {
  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); sendTurn($('input').value); }
};
$('stop').onclick = () => { if (state.controller) state.controller.abort(); };
document.querySelectorAll('#chips [data-send]').forEach(btn => {
  btn.onclick = () => sendTurn(btn.dataset.send);
});
$('new').onclick = () => {
  state.messages = [];
  $('transcript').innerHTML = '';
  addTurn('assistant', '').bubble.textContent = 'New conversation. Select text in the document, then ask for a change.';
  setStatus('Ready.');
};

/* ---------- start ---------- */

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    setStatus('Open this pane inside Word.', true);
    return;
  }
  loadProviders();
  watchSelection();
  refreshControls();
});
