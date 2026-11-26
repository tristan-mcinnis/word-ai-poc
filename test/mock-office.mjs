/* Realistic in-memory Office.js mock environments for Word, Excel, and PowerPoint.
   Used by both unit/integration tests and the multi-host evaluation suite. */

import { parseExcelRef, indexToCol } from '../lib/excel-utils.js';

export function createMockWordEnv({
  paragraphs = ['First paragraph.', 'Second paragraph with target text.'],
  selectionIndex = 1,
  trackChanges = false,
  comments = [],
} = {}) {
  const state = {
    paragraphs: paragraphs.slice(),
    selectionIndex,
    changeTrackingMode: trackChanges ? 'TrackAll' : 'Off',
    comments: comments.map(c => ({ ...c, replies: [] })),
    selectionHandlers: [],
  };

  function makeRange(idx) {
    return {
      _idx: idx,
      get text() { return state.paragraphs[this._idx] ?? ''; },
      load() {},
      expandTo(other) {
        const endIdx = other._idx;
        return {
          _start: idx,
          _end: endIdx,
          load() {},
          insertText(val) {
            state.paragraphs.splice(idx, endIdx - idx + 1, val);
          },
          insertParagraph(val, loc) {
            if (loc === 'After') state.paragraphs.splice(idx + 1, 0, val);
            else state.paragraphs.splice(idx, 0, val);
          },
        };
      },
      insertText(val) {
        if (this._idx >= 0 && this._idx < state.paragraphs.length) {
          state.paragraphs[this._idx] = val;
        }
      },
      insertParagraph(val, loc) {
        if (loc === 'After') state.paragraphs.splice(this._idx + 1, 0, val);
        else if (loc === 'Before') state.paragraphs.splice(this._idx, 0, val);
        else state.paragraphs.push(val);
      },
      insertComment(text) {
        state.comments.push({
          id: `c-${state.comments.length + 1}`,
          authorName: 'User',
          content: text,
          anchor: state.paragraphs[this._idx] || '',
          resolved: false,
          replies: [],
        });
      },
    };
  }

  const Office = {
    HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' },
    EventType: { DocumentSelectionChanged: 'DocumentSelectionChanged' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        addHandlerAsync(_ev, fn, cb) {
          state.selectionHandlers.push(fn);
          if (cb) cb({ status: 'succeeded' });
        },
      },
    },
  };

  const Word = {
    InsertLocation: { replace: 'Replace', after: 'After', before: 'Before', end: 'End' },
    ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' },
    async run(callback) {
      const ctx = {
        document: {
          get changeTrackingMode() { return state.changeTrackingMode; },
          set changeTrackingMode(v) { state.changeTrackingMode = v; },
          load() {},
          getSelection() { return makeRange(state.selectionIndex); },
          body: {
            get text() { return state.paragraphs.join('\n'); },
            load() {},
            insertParagraph(val, loc) {
              if (loc === 'End') state.paragraphs.push(val);
              else state.paragraphs.unshift(val);
            },
            search(snippet) {
              const items = [];
              state.paragraphs.forEach((p, idx) => {
                if (p.includes(snippet)) {
                  const r = makeRange(idx);
                  r.insertText = val => {
                    state.paragraphs[idx] = state.paragraphs[idx].replace(snippet, val);
                  };
                  items.push(r);
                }
              });
              return { items, load() {} };
            },
            paragraphs: {
              get items() {
                return state.paragraphs.map((text, idx) => ({
                  text,
                  getRange: () => makeRange(idx),
                }));
              },
              load() {},
            },
            getComments() {
              return {
                items: state.comments.map(c => ({
                  ...c,
                  getRange: () => ({ text: c.anchor || '', load() {} }),
                  reply(replyText) { c.replies.push(replyText); },
                })),
                load() {},
              };
            },
          },
        },
        async sync() {},
      };
      return callback(ctx);
    },
  };

  return { Office, Word, state };
}

export function createMockExcelEnv({
  sheets = {
    Sheet1: {
      values: [
        ['Metric', 'Q1', 'Q2'],
        ['Revenue', 100, 120],
        ['Cost', 60, 70],
        ['Margin', 0.4, '#DIV/0!'],
      ],
      formulas: [
        ['Metric', 'Q1', 'Q2'],
        ['Revenue', 100, 120],
        ['Cost', 60, 70],
        ['Margin', '=(B2-B3)/B2', '=(C2-C3)/0'],
      ],
    },
  },
  activeSheet = 'Sheet1',
  selectionAddress = 'Sheet1!C4',
} = {}) {
  const sheetStore = new Map();
  let pos = 0;
  for (const [name, data] of Object.entries(sheets)) {
    sheetStore.set(name, {
      name,
      position: pos++,
      values: (data.values || []).map(r => r.slice()),
      formulas: (data.formulas || data.values || []).map(r => r.slice()),
      numberFormats: {},
      formats: [],
      charts: [],
      comments: [],
      sorts: [],
    });
  }

  const state = {
    sheets: sheetStore,
    activeSheet,
    selectionAddress,
    activatedEvents: [],
  };

  function ensureCell(sheetData, r, c) {
    while (sheetData.values.length <= r) sheetData.values.push([]);
    while (sheetData.formulas.length <= r) sheetData.formulas.push([]);
    while (sheetData.values[r].length <= c) sheetData.values[r].push('');
    while (sheetData.formulas[r].length <= c) sheetData.formulas[r].push('');
  }

  function makeRange(sheetData, startRow, startCol, rowCount, colCount) {
    const endRow = startRow + rowCount - 1;
    const endCol = startCol + colCount - 1;
    const refPart = rowCount === 1 && colCount === 1
      ? `${indexToCol(startCol)}${startRow + 1}`
      : `${indexToCol(startCol)}${startRow + 1}:${indexToCol(endCol)}${endRow + 1}`;
    const fullAddr = `${sheetData.name}!${refPart}`;

    const rangeFormat = {
      font: {},
      fill: {},
      horizontalAlignment: '',
      wrapText: false,
      autofitted: false,
      autofitColumns() { this.autofitted = true; },
    };

    return {
      address: fullAddr,
      rowCount,
      columnCount: colCount,
      columnIndex: startCol,
      format: rangeFormat,
      load() {},
      get values() {
        const out = [];
        for (let r = startRow; r <= endRow; r++) {
          const row = [];
          for (let c = startCol; c <= endCol; c++) {
            row.push(sheetData.values[r]?.[c] ?? '');
          }
          out.push(row);
        }
        return out;
      },
      set values(grid) {
        this.formulas = grid;
      },
      get formulas() {
        const out = [];
        for (let r = startRow; r <= endRow; r++) {
          const row = [];
          for (let c = startCol; c <= endCol; c++) {
            row.push(sheetData.formulas[r]?.[c] ?? '');
          }
          out.push(row);
        }
        return out;
      },
      set formulas(grid) {
        for (let r = 0; r < grid.length; r++) {
          for (let c = 0; c < grid[r].length; c++) {
            ensureCell(sheetData, startRow + r, startCol + c);
            sheetData.formulas[startRow + r][startCol + c] = grid[r][c];
            sheetData.values[startRow + r][startCol + c] = grid[r][c];
          }
        }
      },
      set numberFormat(grid) {
        sheetData.numberFormats[refPart] = grid;
      },
      getResizedRange(deltaRows, deltaCols) {
        return makeRange(sheetData, startRow, startCol, rowCount + deltaRows, colCount + deltaCols);
      },
      select() {
        state.activeSheet = sheetData.name;
        state.selectionAddress = fullAddr;
      },
      clear() {
        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            if (sheetData.values[r]) sheetData.values[r][c] = '';
            if (sheetData.formulas[r]) sheetData.formulas[r][c] = '';
          }
        }
      },
      sort: {
        apply(fields, hasHeaders) {
          sheetData.sorts.push({ range: refPart, fields, hasHeaders });
          const headerRow = hasHeaders ? [sheetData.values[startRow].slice()] : [];
          const bodyStart = hasHeaders ? startRow + 1 : startRow;
          const bodyRows = [];
          for (let r = bodyStart; r <= endRow; r++) {
            bodyRows.push((sheetData.values[r] || []).slice(startCol, endCol + 1));
          }
          const { key, ascending } = fields[0];
          bodyRows.sort((a, b) => {
            const av = a[key];
            const bv = b[key];
            if (av === bv) return 0;
            return (av > bv ? 1 : -1) * (ascending ? 1 : -1);
          });
          const merged = headerRow.concat(bodyRows);
          merged.forEach((row, idx) => {
            for (let c = 0; c < row.length; c++) {
              sheetData.values[startRow + idx][startCol + c] = row[c];
              sheetData.formulas[startRow + idx][startCol + c] = row[c];
            }
          });
        },
      },
      _recordFormat: () => {
        sheetData.formats.push({
          range: refPart,
          font: { ...rangeFormat.font },
          fill: { ...rangeFormat.fill },
          horizontalAlignment: rangeFormat.horizontalAlignment,
          wrapText: rangeFormat.wrapText,
          autofitted: rangeFormat.autofitted,
        });
      },
    };
  }

  function makeSheetProxy(sheetData) {
    const activeRanges = [];
    const proxy = {
      name: sheetData.name,
      position: sheetData.position,
      isNullObject: false,
      load() {},
      activate() { state.activeSheet = sheetData.name; },
      getRange(addr) {
        const parsed = parseExcelRef(addr) || { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
        const r = makeRange(
          sheetData,
          parsed.startRow,
          parsed.startCol,
          parsed.endRow - parsed.startRow + 1,
          parsed.endCol - parsed.startCol + 1,
        );
        activeRanges.push(r);
        return r;
      },
      getUsedRangeOrNullObject() {
        const rows = sheetData.values.length;
        const cols = sheetData.values.reduce((m, r) => Math.max(m, r.length), 0);
        if (!rows || !cols) return { isNullObject: true, load() {} };
        const r = makeRange(sheetData, 0, 0, rows, cols);
        r.isNullObject = false;
        return r;
      },
      charts: {
        add(chartType, dataRange, seriesBy) {
          const chart = {
            chartType,
            dataAddress: dataRange.address,
            seriesBy,
            title: { text: '', visible: false },
          };
          sheetData.charts.push(chart);
          return chart;
        },
      },
      comments: {
        get items() {
          return sheetData.comments.map(c => ({
            ...c,
            getLocation: () => ({ address: c.cell, load() {} }),
          }));
        },
        load() {},
        add(range, content) {
          sheetData.comments.push({
            id: `ec-${sheetData.comments.length + 1}`,
            authorName: 'Analyst',
            cell: typeof range === 'string' ? range : range.address,
            content,
            resolved: false,
          });
        },
      },
      _flushFormats() {
        for (const r of activeRanges) {
          if (Object.keys(r.format.font).length || Object.keys(r.format.fill).length || r.format.autofitted || r.format.horizontalAlignment || r.format.wrapText) {
            r._recordFormat();
          }
        }
      },
    };
    return proxy;
  }

  const Office = {
    HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' },
    EventType: { DocumentSelectionChanged: 'DocumentSelectionChanged' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        addHandlerAsync(_ev, _fn, cb) { if (cb) cb({ status: 'succeeded' }); },
      },
    },
  };

  const Excel = {
    ChartSeriesBy: { auto: 'Auto' },
    ClearApplyTo: { contents: 'Contents' },
    async run(callback) {
      const createdProxies = [];
      const wrapSheet = data => {
        const p = makeSheetProxy(data);
        createdProxies.push(p);
        return p;
      };
      const ctx = {
        workbook: {
          worksheets: {
            get items() { return [...sheetStore.values()].map(wrapSheet); },
            load() {},
            getActiveWorksheet() { return wrapSheet(sheetStore.get(state.activeSheet)); },
            getItemOrNullObject(name) {
              const found = sheetStore.get(name);
              if (!found) return { isNullObject: true, load() {} };
              return wrapSheet(found);
            },
            getItem(name) {
              const found = sheetStore.get(name);
              if (!found) throw new Error(`Sheet not found: ${name}`);
              return wrapSheet(found);
            },
            add(name) {
              const created = {
                name,
                position: sheetStore.size,
                values: [],
                formulas: [],
                numberFormats: {},
                formats: [],
                charts: [],
                comments: [],
                sorts: [],
              };
              sheetStore.set(name, created);
              return wrapSheet(created);
            },
            onActivated: {
              add(fn) { state.activatedEvents.push(fn); },
            },
          },
          getSelectedRange() {
            const parsed = parseExcelRef(state.selectionAddress) || { sheet: state.activeSheet, startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
            const sData = sheetStore.get(parsed.sheet || state.activeSheet);
            const p = wrapSheet(sData);
            return p.getRange(parsed.range || 'A1');
          },
        },
        async sync() {
          for (const p of createdProxies) p._flushFormats();
        },
      };
      return callback(ctx);
    },
  };

  return { Office, Excel, state };
}

export function createMockPowerPointEnv({
  layouts = [
    { id: 'layout-title', name: 'Title Slide' },
    { id: 'layout-content', name: 'Title and Content' },
    { id: 'layout-two', name: 'Two Content' },
    { id: 'layout-blank', name: 'Blank' },
  ],
  slides = [
    {
      id: 'sl-1',
      layoutName: 'Title Slide',
      shapes: [
        { id: 'sh-1', name: 'Title 1', type: 'Placeholder', top: 120, left: 60, text: 'Q3 Board Update' },
        { id: 'sh-2', name: 'Subtitle 2', type: 'Placeholder', top: 240, left: 60, text: 'Prepared by Finance' },
      ],
    },
    {
      id: 'sl-2',
      layoutName: 'Title and Content',
      shapes: [
        { id: 'sh-3', name: 'Title 1', type: 'Placeholder', top: 30, left: 50, text: 'Revenue Overview' },
        { id: 'sh-4', name: 'Content Placeholder 2', type: 'Placeholder', top: 130, left: 50, text: 'Old Product Name grew 22%.\nWe need to tighten this text.' },
      ],
    },
  ],
  selectedSlideIds = ['sl-2'],
  selectedShapeIds = ['sh-4'],
} = {}) {
  let idCounter = 100;
  const state = {
    layouts: layouts.map(l => ({ ...l })),
    slides: slides.map(sl => ({
      ...sl,
      shapes: sl.shapes.map(sh => ({ ...sh })),
      tables: [],
    })),
    selectedSlideIds: selectedSlideIds.slice(),
    selectedShapeIds: selectedShapeIds.slice(),
    navigatedSlides: [],
  };

  function makeShapeProxy(sh) {
    return {
      id: sh.id,
      name: sh.name,
      type: sh.type,
      left: sh.left ?? 50,
      top: sh.top ?? 50,
      width: sh.width ?? 500,
      height: sh.height ?? 100,
      load() {},
      textFrame: sh.type === 'Image' || sh.type === 'Chart' ? undefined : {
        get hasText() { return true; },
        load() {},
        textRange: {
          get text() { return sh.text; },
          set text(v) { sh.text = v; },
          font: {
            set size(v) { sh.fontSize = v; },
            set bold(v) { sh.bold = v; },
          },
          load() {},
        },
      },
    };
  }

  function makeSlideProxy(sl) {
    return {
      id: sl.id,
      layout: { name: sl.layoutName, load() {} },
      load() {},
      delete() {
        const idx = state.slides.findIndex(s => s.id === sl.id);
        if (idx >= 0) state.slides.splice(idx, 1);
      },
      shapes: {
        get items() { return sl.shapes.map(makeShapeProxy); },
        load() {},
        addTextBox(text, opts = {}) {
          const sh = {
            id: `sh-${++idCounter}`,
            name: `TextBox ${idCounter}`,
            type: 'TextBox',
            left: opts.left ?? 60,
            top: opts.top ?? 120,
            width: opts.width ?? 600,
            height: opts.height ?? 100,
            text,
          };
          sl.shapes.push(sh);
          return makeShapeProxy(sh);
        },
        addTable(rows, cols, opts = {}) {
          const tbl = {
            id: `tbl-${++idCounter}`,
            name: `Table ${idCounter}`,
            type: 'Table',
            rows,
            cols,
            values: opts.values || [],
          };
          sl.tables.push(tbl);
          sl.shapes.push({ id: tbl.id, name: tbl.name, type: 'Table', text: '' });
          return tbl;
        },
      },
    };
  }

  const Office = {
    HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' },
    EventType: { DocumentSelectionChanged: 'DocumentSelectionChanged' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    GoToType: { Slide: 'slide' },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        addHandlerAsync(_ev, _fn, cb) { if (cb) cb({ status: 'succeeded' }); },
        goToByIdAsync(id, _type, cb) {
          state.navigatedSlides.push(id);
          if (cb) cb({ status: 'succeeded' });
        },
      },
    },
  };

  const PowerPoint = {
    async run(callback) {
      const ctx = {
        presentation: {
          slideMasters: {
            items: [
              {
                id: 'master-1',
                name: 'Corporate Master',
                layouts: {
                  items: state.layouts.map(l => ({ id: l.id, name: l.name })),
                  load() {},
                },
              },
            ],
            load() {},
          },
          slides: {
            get items() { return state.slides.map(makeSlideProxy); },
            load() {},
            add(opts = {}) {
              const matchedLayout = state.layouts.find(l => l.id === opts?.layoutId) || state.layouts[1];
              const layoutName = matchedLayout?.name || 'Blank';
              const newShapes = /blank/i.test(layoutName)
                ? []
                : /two/i.test(layoutName)
                  ? [
                      { id: `sh-${++idCounter}`, name: 'Title 1', type: 'Placeholder', top: 30, left: 50, text: '' },
                      { id: `sh-${++idCounter}`, name: 'Content Placeholder 2', type: 'Placeholder', top: 130, left: 50, text: '' },
                      { id: `sh-${++idCounter}`, name: 'Content Placeholder 3', type: 'Placeholder', top: 130, left: 370, text: '' },
                    ]
                  : [
                      { id: `sh-${++idCounter}`, name: 'Title 1', type: 'Placeholder', top: 30, left: 50, text: '' },
                      { id: `sh-${++idCounter}`, name: 'Content Placeholder 2', type: 'Placeholder', top: 130, left: 50, text: '' },
                    ];
              const newSlide = {
                id: `sl-${++idCounter}`,
                layoutName,
                shapes: newShapes,
                tables: [],
              };
              if (typeof opts?.index === 'number' && opts.index >= 0 && opts.index <= state.slides.length) {
                state.slides.splice(opts.index, 0, newSlide);
              } else {
                state.slides.push(newSlide);
              }
            },
          },
          getSelectedSlides() {
            const items = state.slides.filter(s => state.selectedSlideIds.includes(s.id)).map(makeSlideProxy);
            return { items, load() {} };
          },
          getSelectedShapes() {
            const active = state.slides.find(s => state.selectedSlideIds.includes(s.id));
            const items = (active?.shapes || []).filter(sh => state.selectedShapeIds.includes(sh.id)).map(makeShapeProxy);
            return { items, load() {} };
          },
          setSelectedSlides(ids) {
            state.selectedSlideIds = ids.slice();
          },
        },
        async sync() {},
      };
      return callback(ctx);
    },
  };

  return { Office, PowerPoint, state };
}
