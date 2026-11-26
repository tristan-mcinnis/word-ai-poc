import test from 'node:test';
import assert from 'node:assert/strict';
import { detectOfficeHost, createHostAdapter } from '../lib/hosts/index.js';
import {
  createMockWordEnv,
  createMockExcelEnv,
  createMockPowerPointEnv,
} from './mock-office.mjs';

test('detectOfficeHost accurately detects Word, Excel, and PowerPoint', () => {
  const Office = { HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' } };
  assert.equal(detectOfficeHost({ host: 'Word' }, Office), 'word');
  assert.equal(detectOfficeHost({ host: 'Excel' }, Office), 'excel');
  assert.equal(detectOfficeHost({ host: 'PowerPoint' }, Office), 'powerpoint');
  assert.equal(detectOfficeHost(null, Office), null);
});

test('Word adapter reads context, toggles Track Changes, and executes all Word edits', async () => {
  const env = createMockWordEnv({
    paragraphs: ['Intro paragraph.', 'Original target sentence.', 'Closing paragraph.'],
    selectionIndex: 1,
    comments: [{ id: 'c-1', authorName: 'Alex', content: 'Make sharper', anchor: 'Original target sentence.', resolved: false }],
  });
  const word = createHostAdapter('word', env);

  const ctx = await word.readContext();
  assert.equal(ctx.selectionText, 'Original target sentence.');
  assert.match(ctx.documentText, /Intro paragraph/);
  assert.equal(ctx.comments.length, 1);
  assert.equal(ctx.trackChanges, false);

  // Toggle native Track Changes
  const toggled = await word.toggleSecondary(true);
  assert.equal(toggled.ok, true);
  assert.equal(env.state.changeTrackingMode, 'TrackAll');

  // Replace selection with multi-paragraph output
  const r1 = await word.applyEdit({ action: 'replace_selection', content: 'Crisp paragraph A.\nCrisp paragraph B.' }, { trackChanges: true });
  assert.equal(r1.ok, true);
  assert.deepEqual(env.state.paragraphs, ['Intro paragraph.', 'Crisp paragraph A.', 'Crisp paragraph B.', 'Closing paragraph.']);

  // Replace match + reply to comment
  const r2 = await word.applyEdit({ action: 'replace_match', find: 'Intro paragraph.', content: 'Executive summary.' });
  assert.equal(r2.ok, true);
  assert.equal(env.state.paragraphs[0], 'Executive summary.');

  const r3 = await word.applyEdit({ action: 'reply_comment', commentId: 'c-1', content: 'Updated as suggested.' });
  assert.equal(r3.ok, true);
  assert.deepEqual(env.state.comments[0].replies, ['Updated as suggested.']);

  // Insert after + append + comment
  await word.applyEdit({ action: 'insert_after', content: 'Chinese translation.' });
  await word.applyEdit({ action: 'append_doc', content: 'Appendix.' });
  await word.applyEdit({ action: 'comment', content: 'Double check metric.' });
  assert.equal(env.state.paragraphs[env.state.paragraphs.length - 1], 'Appendix.');
  assert.equal(env.state.comments.length, 2);
});

test('Excel adapter reads formulas/errors, navigates citations, and executes all Excel edits', async () => {
  const env = createMockExcelEnv({
    sheets: {
      Model: {
        values: [
          ['Item', 'Q1', 'Q2'],
          ['Revenue', 100, 140],
          ['Cost', 60, 80],
          ['Margin', 0.4, '#DIV/0!'],
        ],
        formulas: [
          ['Item', 'Q1', 'Q2'],
          ['Revenue', 100, 140],
          ['Cost', 60, 80],
          ['Margin', '=(B2-B3)/B2', '=(C2-C3)/0'],
        ],
      },
      Assumptions: {
        values: [['Tax Rate', 0.21]],
        formulas: [['Tax Rate', 0.21]],
      },
    },
    activeSheet: 'Model',
    selectionAddress: 'Model!C4',
  });
  const excel = createHostAdapter('excel', env);

  const ctx = await excel.readContext();
  assert.equal(ctx.hostContext.activeSheet, 'Model');
  assert.equal(ctx.hostContext.sheets.length, 2);
  assert.equal(ctx.hostContext.errors.length, 1);
  assert.match(ctx.hostContext.errors[0], /#DIV\/0!/);
  assert.match(ctx.documentText, /Assumptions/);

  // Fix formula in C4 with highlight ON (even if model omits leading '=')
  const fRes = await excel.applyEdit(
    { action: 'set_formula', attrs: { cell: 'C4', sheet: 'Model', number_format: '0.0%' }, content: 'IFERROR((C2-C3)/C2,0)' },
    { highlightEdits: true },
  );
  assert.equal(fRes.ok, true);
  assert.equal(env.state.sheets.get('Model').formulas[3][2], '=IFERROR((C2-C3)/C2,0)');
  assert.ok(env.state.sheets.get('Model').formats.some(f => f.fill.color === '#FEF9C3'));

  // 1x1 set_cells write without resize
  const oneRes = await excel.applyEdit({
    action: 'set_cells',
    attrs: { range: 'E1', sheet: 'Model' },
    content: 'Notes',
  });
  assert.equal(oneRes.ok, true);
  assert.equal(env.state.sheets.get('Model').values[0][4], 'Notes');

  // Write multi-cell grid starting from anchor cell D1 (auto-resizes cleanly)
  const cRes = await excel.applyEdit({
    action: 'set_cells',
    attrs: { range: 'D1', sheet: 'Model' },
    content: JSON.stringify([['Total'], ['=SUM(B2:C2)'], ['=SUM(B3:C3)']]),
  });
  assert.equal(cRes.ok, true);
  assert.equal(env.state.sheets.get('Model').formulas[0][3], 'Total');
  assert.equal(env.state.sheets.get('Model').formulas[1][3], '=SUM(B2:C2)');

  // Format header row
  const fmtRes = await excel.applyEdit({
    action: 'format_cells',
    attrs: { range: 'A1:D1', sheet: 'Model' },
    content: JSON.stringify({ bold: true, fill: '#F3F4F6', autoFit: true }),
  });
  assert.equal(fmtRes.ok, true);
  assert.ok(env.state.sheets.get('Model').formats.some(f => f.font.bold === true && f.autofitted === true));

  // Add chart
  const chRes = await excel.applyEdit({
    action: 'add_chart',
    attrs: { range: 'A1:C3', sheet: 'Model', chart_type: 'ColumnClustered', title: 'Q1 vs Q2' },
    content: '',
  });
  assert.equal(chRes.ok, true);
  assert.equal(env.state.sheets.get('Model').charts.length, 1);
  assert.equal(env.state.sheets.get('Model').charts[0].title.text, 'Q1 vs Q2');

  // Add new populated sheet
  const shRes = await excel.applyEdit({
    action: 'add_sheet',
    attrs: { name: 'Summary' },
    content: JSON.stringify([['KPI', 'Value'], ['Total Rev', '=Model!D2']]),
  });
  assert.equal(shRes.ok, true);
  assert.ok(env.state.sheets.has('Summary'));
  assert.equal(env.state.sheets.get('Summary').formulas[1][1], '=Model!D2');

  // Sort, comment, clear, and navigateToRef
  await excel.applyEdit({
    action: 'sort_range',
    attrs: { range: 'A1:C3', sheet: 'Model', column: 'B', order: 'desc', has_headers: 'true' },
    content: '',
  });
  assert.equal(env.state.sheets.get('Model').sorts.length, 1);

  await excel.applyEdit({
    action: 'excel_comment',
    attrs: { cell: 'C4', sheet: 'Model' },
    content: 'Verified margin formula.',
  });
  assert.equal(env.state.sheets.get('Model').comments.length, 1);

  const nav = await excel.navigateToRef('Assumptions!B1');
  assert.equal(nav.ok, true);
  assert.equal(env.state.activeSheet, 'Assumptions');
  assert.equal(env.state.selectionAddress, 'Assumptions!B1');

  await excel.applyEdit({ action: 'clear_range', attrs: { range: 'D1:D3', sheet: 'Model' }, content: '' });
  assert.equal(env.state.sheets.get('Model').values[0][3], '');
});

test('PowerPoint adapter reads deck/layouts, navigates slides, and executes all PowerPoint edits', async () => {
  const env = createMockPowerPointEnv();
  const ppt = createHostAdapter('powerpoint', env);

  const ctx = await ppt.readContext();
  assert.equal(ctx.hostContext.slideCount, 2);
  assert.equal(ctx.hostContext.currentSlideIndex, 2);
  assert.equal(ctx.hostContext.selectedShapeName, 'Content Placeholder 2');
  assert.equal(ctx.hostContext.layouts.length, 4);
  assert.match(ctx.documentText, /Available Slide Master layouts/);

  // Update shape text
  const uRes = await ppt.applyEdit({
    action: 'update_shape',
    attrs: { slide: '2', shape: 'Title 1' },
    content: 'Enterprise growth drove +22% Q3 revenue',
  });
  assert.equal(uRes.ok, true);
  assert.equal(env.state.slides[1].shapes[0].text, 'Enterprise growth drove +22% Q3 revenue');

  // Deck-wide replace_slide_text
  const rRes = await ppt.applyEdit({
    action: 'replace_slide_text',
    find: 'Old Product Name',
    content: 'Nova Platform',
    attrs: {},
  });
  assert.equal(rRes.ok, true);
  assert.match(env.state.slides[1].shapes[1].text, /Nova Platform/);

  // Add on-brand slide matching Slide Master layout + filling native placeholders
  const addRes = await ppt.applyEdit({
    action: 'add_slide',
    attrs: { layout: 'Title and Content', after_slide: '2' },
    content: 'Title: Q4 Priorities\n- Close 15 enterprise deals\n- Launch self-serve onboarding',
  });
  assert.equal(addRes.ok, true);
  assert.equal(env.state.slides.length, 3);
  assert.equal(env.state.slides[2].layoutName, 'Title and Content');
  assert.equal(env.state.slides[2].shapes[0].text, 'Q4 Priorities');
  assert.match(env.state.slides[2].shapes[1].text, /Close 15 enterprise deals/);

  // Add blank slide fallback (creates text boxes when layout has no placeholders)
  const blankRes = await ppt.applyEdit({
    action: 'add_slide',
    attrs: { layout: 'Blank' },
    content: 'Title: Appendix\n- Backup metrics',
  });
  assert.equal(blankRes.ok, true);
  assert.equal(env.state.slides.length, 4);
  assert.equal(env.state.slides[3].shapes.length, 2);

  // Add textbox and native table
  const tbRes = await ppt.applyEdit({
    action: 'add_textbox',
    attrs: { slide: '2', left: '60', top: '440', font_size: '11', bold: 'true' },
    content: 'Source: Finance close',
  });
  assert.equal(tbRes.ok, true);
  assert.ok(env.state.slides[1].shapes.some(s => s.text === 'Source: Finance close' && s.fontSize === 11 && s.bold === true));

  const tblRes = await ppt.applyEdit({
    action: 'add_table',
    attrs: { slide: '2' },
    content: '| Metric | Q2 | Q3 |\n|---|---|---|\n| ARR | $10M | $12.2M |',
  });
  assert.equal(tblRes.ok, true);
  assert.equal(env.state.slides[1].tables.length, 1);
  assert.deepEqual(env.state.slides[1].tables[0].values, [['Metric', 'Q2', 'Q3'], ['ARR', '$10M', '$12.2M']]);

  // Insert a cover slide at the very beginning (after_slide="0")
  const coverRes = await ppt.applyEdit({
    action: 'add_slide',
    attrs: { layout: 'Title Slide', after_slide: '0' },
    content: 'Title: Executive Cover',
  });
  assert.equal(coverRes.ok, true);
  assert.equal(env.state.slides[0].shapes[0].text, 'Executive Cover');

  // Navigate to Slide 1
  const nav = await ppt.navigateToRef('Slide 1');
  assert.equal(nav.ok, true);
  assert.deepEqual(env.state.selectedSlideIds, [env.state.slides[0].id]);

  // Delete Slide 5
  const delRes = await ppt.applyEdit({ action: 'delete_slide', attrs: { slide: '5' }, content: '' });
  assert.equal(delRes.ok, true);
  assert.equal(env.state.slides.length, 4);
});
