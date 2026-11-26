import test from 'node:test';
import assert from 'node:assert/strict';
import {
  indexToCol,
  colToIndex,
  parseExcelRef,
  formatExcelAddress,
  coerceCellValue,
  parseCellGrid,
  formatSheetSnapshot,
  parseFormatSpec,
  diffCellGrids,
} from '../lib/excel-utils.js';

test('indexToCol and colToIndex round-trip single and multi-letter columns', () => {
  assert.equal(indexToCol(0), 'A');
  assert.equal(indexToCol(25), 'Z');
  assert.equal(indexToCol(26), 'AA');
  assert.equal(indexToCol(27), 'AB');
  assert.equal(colToIndex('A'), 0);
  assert.equal(colToIndex('Z'), 25);
  assert.equal(colToIndex('AA'), 26);
  assert.equal(colToIndex('AB'), 27);
  assert.equal(colToIndex(''), -1);
});

test('parseExcelRef handles cells, ranges, absolute refs, and quoted sheet names', () => {
  assert.deepEqual(parseExcelRef('B4'), {
    sheet: '',
    range: 'B4',
    startCell: 'B4',
    endCell: 'B4',
    startCol: 1,
    startRow: 3,
    endCol: 1,
    endRow: 3,
    isRange: false,
  });

  const quoted = parseExcelRef("'Q1 P&L'!$B$2:$D$10");
  assert.equal(quoted.sheet, 'Q1 P&L');
  assert.equal(quoted.range, 'B2:D10');
  assert.equal(quoted.startCol, 1);
  assert.equal(quoted.startRow, 1);
  assert.equal(quoted.endCol, 3);
  assert.equal(quoted.endRow, 9);
  assert.equal(quoted.isRange, true);

  assert.equal(formatExcelAddress('Q1 P&L', 'B2:D10'), "'Q1 P&L'!B2:D10");
  assert.equal(formatExcelAddress('Sheet1', 'A1'), 'Sheet1!A1');
  assert.equal(parseExcelRef('not-a-cell'), null);
});

test('coerceCellValue preserves formulas, numbers, booleans, and leading-zero IDs', () => {
  assert.equal(coerceCellValue('=SUM(B2:B5)'), '=SUM(B2:B5)');
  assert.equal(coerceCellValue('1450.5'), 1450.5);
  assert.equal(coerceCellValue('-42'), -42);
  assert.equal(coerceCellValue('0'), 0);
  assert.equal(coerceCellValue('00412'), '00412');
  assert.equal(coerceCellValue('true'), true);
  assert.equal(coerceCellValue('FALSE'), false);
  assert.equal(coerceCellValue('Revenue'), 'Revenue');
});

test('parseCellGrid parses JSON 2D arrays, Markdown tables, TSV, and CSV with formulas', () => {
  const fromJson = parseCellGrid('```json\n[["Item","Q1","Q2"],["Rev",100,120],["Total","=SUM(B2:B2)","=SUM(C2:C2)"]]\n```');
  assert.deepEqual(fromJson, [
    ['Item', 'Q1', 'Q2'],
    ['Rev', 100, 120],
    ['Total', '=SUM(B2:B2)', '=SUM(C2:C2)'],
  ]);

  const fromMd = parseCellGrid('| Metric | Value |\n|---|---|\n| ARR | 1200 |\n| Margin | =B2/100 |');
  assert.deepEqual(fromMd, [
    ['Metric', 'Value'],
    ['ARR', 1200],
    ['Margin', '=B2/100'],
  ]);

  const fromTsv = parseCellGrid('A\tB\n1\t=A2*2');
  assert.deepEqual(fromTsv, [['A', 'B'], [1, '=A2*2']]);

  // CSV with comma inside formula parentheses must NOT split inside IFERROR(...)
  const fromCsv = parseCellGrid('Metric,Formula\nMargin,=IFERROR(B2/C2,0)');
  assert.deepEqual(fromCsv, [
    ['Metric', 'Formula'],
    ['Margin', '=IFERROR(B2/C2,0)'],
  ]);
});

test('formatSheetSnapshot serializes formulas alongside values and flags Excel errors', () => {
  const snap = formatSheetSnapshot({
    sheetName: 'Model',
    address: 'Model!A1:C3',
    values: [
      ['Item', 'Rev', 'Margin'],
      ['Alpha', 100, 0.25],
      ['Beta', 0, '#DIV/0!'],
    ],
    formulas: [
      ['Item', 'Rev', 'Margin'],
      ['Alpha', 100, '=25/B2'],
      ['Beta', 0, '=10/B3'],
    ],
  });

  assert.match(snap.text, /\| Row \| A \| B \| C \|/);
  assert.match(snap.text, /=25\/B2 \[0\.25\]/);
  assert.match(snap.text, /=10\/B3 \[#DIV\/0!\]/);
  assert.equal(snap.errors.length, 1);
  assert.equal(snap.errors[0], 'Model!C3: #DIV/0! (=10/B3)');
});

test('parseFormatSpec and diffCellGrids work cleanly', () => {
  assert.deepEqual(parseFormatSpec('```json\n{"bold":true,"fill":"#F3F4F6"}\n```'), {
    bold: true,
    fill: '#F3F4F6',
  });
  assert.deepEqual(parseFormatSpec('not json'), {});

  const changes = diffCellGrids(
    [['Rev', 100], ['Cost', 60]],
    [['Rev', 120], ['Cost', 60]],
    'B2',
  );
  assert.deepEqual(changes, [{ cell: 'C2', before: 100, after: 120 }]);
});
