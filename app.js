/* global Office, Word */
const $ = id => document.getElementById(id);
let originalRange, originalText = '', busy = false;
function status(text) { $('status').textContent = text; }
async function run(fn) {
  if (busy) return;
  busy = true;
  for (const id of ['read','demo','apply']) $(id).disabled = true;
  try { await fn(); } catch(e) { status(e.message || String(e)); }
  finally {
    busy = false; $('read').disabled = false;
    $('demo').disabled = $('apply').disabled = !originalRange;
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
  $('original').value = originalText; status('Selection captured. Tidy the whitespace, review, then apply.');
});
$('demo').onclick = () => run(async () => {
  $('preview').value = originalText.split('\n').map(line => line.replace(/[\t ]+/g,' ').trim()).join('\n');
  status('Whitespace cleaned. Review the preview before applying.');
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
