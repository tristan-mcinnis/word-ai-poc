/* global Office, Word */
const $ = id => document.getElementById(id);
let originalRange, originalText = '', busy = false;
function status(text) { $('status').textContent = text; }
async function run(fn) {
  if (busy) return;
  busy = true;
  for (const id of ['read','demo','ai','apply']) $(id).disabled = true;
  try { await fn(); } catch(e) { status(e.message || String(e)); }
  finally {
    busy = false; $('read').disabled = false;
    $('demo').disabled = $('ai').disabled = $('apply').disabled = !originalRange;
  }
}
Office.onReady(info => {
  if(info.host !== Office.HostType.Word) { status('Open this add-in inside Word.'); return; }
  status('Ready. Select a short paragraph in a scratch document.'); $('read').disabled = false;
});
$('read').onclick = () => run(async () => {
  originalRange = null; originalText = ''; $('original').value = $('preview').value = '';
  await Word.run(async ctx => {
    const range = ctx.document.getSelection(); range.load('text'); await ctx.sync();
    if (!range.text.trim()) throw Error('Select some text first.');
    if (range.text.length > 12000) throw Error('Use a shorter selection (under 12,000 characters).');
    range.track(); await ctx.sync(); originalRange = range; originalText = range.text;
  });
  $('original').value = originalText; status('Selection captured. Choose demo or local AI.');
});
$('demo').onclick = () => run(async () => {
  $('preview').value = originalText.split('\n').map(line => line.replace(/[\t ]+/g,' ').trim()).join('\n');
  status('Demo only: whitespace cleanup. No AI used. Review before applying.');
});
$('ai').onclick = () => run(async () => {
  if(!$('model').value.trim()) throw Error('Enter an installed Ollama model name.');
  status('Asking local AI…');
  const response = await fetch('/rewrite', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({text:originalText,model:$('model').value.trim()})});
  const result = await response.json(); if(!response.ok) throw Error(result.error);
  $('preview').value = result.text; status('Local AI draft ready. Review before applying.');
});
$('apply').onclick = () => run(async () => {
  const text = $('preview').value; if(!text.trim()) throw Error('Create a preview first.');
  await Word.run(originalRange, async ctx => {
    originalRange.load('text'); await ctx.sync();
    if(originalRange.text !== originalText) throw Error('Original text changed. Read the selection again.');
    originalRange.insertText(text,Word.InsertLocation.replace); await ctx.sync();
    originalRange.untrack(); await ctx.sync();
  });
  originalRange = null; status('Applied. Use Word Undo to revert.');
});
