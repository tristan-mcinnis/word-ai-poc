import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseSlideSpec,
  matchSlideLayout,
  classifyPlaceholders,
  formatPresentationSnapshot,
  parseSlideRef,
} from '../lib/ppt-utils.js';

test('parseSlideSpec parses structured text and JSON slide specs', () => {
  const textSpec = parseSlideSpec([
    'Title: Q3 Commercial Momentum',
    'Subtitle: Enterprise expansion led growth',
    '- ARR grew 38% YoY',
    '- Net retention hit 124%',
  ].join('\n'));

  assert.equal(textSpec.title, 'Q3 Commercial Momentum');
  assert.equal(textSpec.subtitle, 'Enterprise expansion led growth');
  assert.deepEqual(textSpec.bullets, ['ARR grew 38% YoY', 'Net retention hit 124%']);
  assert.equal(textSpec.body, 'ARR grew 38% YoY\nNet retention hit 124%');

  const jsonSpec = parseSlideSpec(JSON.stringify({
    title: 'Comparison',
    columns: [['Pro A', 'Pro B'], ['Risk A', 'Risk B']],
  }));
  assert.equal(jsonSpec.title, 'Comparison');
  assert.deepEqual(jsonSpec.columns, ['Pro A\nPro B', 'Risk A\nRisk B']);
});

test('matchSlideLayout matches exact, normalized, token-overlap, and default layouts', () => {
  const layouts = [
    { slideMasterId: 'm1', layoutId: 'l1', layoutName: 'Title Slide' },
    { slideMasterId: 'm1', layoutId: 'l2', layoutName: 'Title and Content' },
    { slideMasterId: 'm1', layoutId: 'l3', layoutName: 'Two Content' },
    { slideMasterId: 'm1', layoutId: 'l4', layoutName: 'Blank' },
  ];

  assert.equal(matchSlideLayout(layouts, 'Title Slide').layoutId, 'l1');
  assert.equal(matchSlideLayout(layouts, 'title-and-content').layoutId, 'l2');
  assert.equal(matchSlideLayout(layouts, 'two column').layoutId, 'l3');
  assert.equal(matchSlideLayout(layouts, 'blank').layoutId, 'l4');
  assert.equal(matchSlideLayout(layouts, '').layoutId, 'l2');
  assert.equal(matchSlideLayout([], 'Blank'), null);
});

test('classifyPlaceholders identifies Title, Subtitle, and Body shapes and ignores footers', () => {
  const shapes = [
    { id: 's4', name: 'Footer Placeholder 4', top: 500, left: 40 },
    { id: 's3', name: 'Slide Number Placeholder 5', top: 500, left: 650 },
    { id: 's2', name: 'Content Placeholder 2', top: 140, left: 50 },
    { id: 's1', name: 'Title 1', top: 30, left: 50 },
  ];
  const { titleShape, subtitleShape, bodyShapes } = classifyPlaceholders(shapes);
  assert.equal(titleShape.id, 's1');
  assert.equal(subtitleShape, null);
  assert.deepEqual(bodyShapes.map(s => s.id), ['s2']);

  // Positional fallback for localized shape names
  const localized = [
    { id: 'b', name: '形状 2', top: 150, left: 50 },
    { id: 'a', name: '形状 1', top: 30, left: 50 },
  ];
  const fallback = classifyPlaceholders(localized);
  assert.equal(fallback.titleShape.id, 'a');
  assert.deepEqual(fallback.bodyShapes.map(s => s.id), ['b']);
});

test('formatPresentationSnapshot and parseSlideRef format deck cleanly', () => {
  const out = formatPresentationSnapshot({
    layouts: [{ layoutName: 'Title Slide' }, { layoutName: 'Title and Content' }],
    selectedSlideIndices: [2],
    slides: [
      { index: 1, layoutName: 'Title Slide', shapes: [{ name: 'Title 1', text: 'Board Update' }] },
      { index: 2, layoutName: 'Title and Content', shapes: [{ name: 'Title 1', text: 'Q3 Revenue' }, { name: 'Chart 1', type: 'Chart', text: '' }] },
    ],
  });

  assert.match(out, /Available Slide Master layouts: Title Slide, Title and Content/);
  assert.match(out, /--- Slide 1 \(Title Slide\) ---/);
  assert.match(out, /--- Slide 2 \(Title and Content\) \[SELECTED\] ---/);
  assert.match(out, /\[Chart 1 \(Chart\)\]/);

  assert.equal(parseSlideRef('Slide 3'), 3);
  assert.equal(parseSlideRef('slide #12'), 12);
  assert.equal(parseSlideRef('2'), 2);
  assert.equal(parseSlideRef('0'), null);
  assert.equal(parseSlideRef('abc'), null);
});
