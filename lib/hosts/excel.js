/* Microsoft Excel host adapter (Office.js ExcelApi), inspired by Claude in Excel. */

import {
  colToIndex,
  parseExcelRef,
  formatExcelAddress,
  parseCellGrid,
  parseFormatSpec,
  formatSheetSnapshot,
} from '../excel-utils.js';

export const EXCEL_QUICK_ACTIONS = [
  { cmd: '/explain-formula', title: 'Explain formula', prompt: 'Explain the formula in the selected cell(s) step-by-step, trace every cell reference it depends on, and flag any subtle assumptions or risks.' },
  { cmd: '/audit-model', title: 'Audit workbook', prompt: 'Audit this workbook for formula errors (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A), hardcoded overrides inside formula rows/columns, broken totals, and structural risks. Cite exact cells before making edits.' },
  { cmd: '/fix-errors', title: 'Fix formula errors', prompt: 'Inspect every formula error in this workbook, explain the root cause with exact cell citations, and propose clean fixes using set_formula or set_cells.' },
  { cmd: '/write-formula', title: 'Write formula', prompt: 'Write the right dynamic Excel formula for the selected cell or range using set_formula (and set an appropriate number_format if helpful).' },
  { cmd: '/summarize-data', title: 'Summarize data', prompt: 'Summarize the key metrics, trends, totals, and outliers in this sheet with exact cell citations and a concise Markdown summary table.' },
  { cmd: '/clean-data', title: 'Clean & standardize', prompt: 'Inspect the selected range for blanks, duplicates, inconsistent labels, or mixed types, and propose a cleaned grid using set_cells.' },
  { cmd: '/build-table', title: 'Build summary table', prompt: 'Build a clean dynamic summary table using Excel formulas (=SUM, =AVERAGE, =SUMIFS, etc.) and propose it via set_cells or add_sheet.' },
  { cmd: '/add-chart', title: 'Create native chart', prompt: 'Recommend the clearest chart for this data and propose creating it natively with add_chart.' },
  { cmd: '/format-table', title: 'Format selection', prompt: 'Propose clean, executive-ready formatting for the selected table (header fill/bold, currency/percent number formats, alignment, autoFit) using format_cells.' },
  { cmd: '/comment', title: 'Add cell comment', prompt: 'Write a concise audit or review comment on the selected cell using excel_comment.' },
];

const HIGHLIGHT_FILL = '#FEF9C3';

function stripSheetFromAddress(address) {
  const s = String(address || '');
  const bang = s.lastIndexOf('!');
  return bang >= 0 ? s.slice(bang + 1) : s;
}

async function resolveSheet(ctx, sheetName) {
  if (!sheetName) return ctx.workbook.worksheets.getActiveWorksheet();
  const sheet = ctx.workbook.worksheets.getItemOrNullObject
    ? ctx.workbook.worksheets.getItemOrNullObject(sheetName)
    : ctx.workbook.worksheets.getItem(sheetName);
  sheet.load('name');
  await ctx.sync();
  if (sheet.isNullObject) {
    throw new Error(`Worksheet "${sheetName}" was not found in this workbook.`);
  }
  return sheet;
}

/**
 * Two-phase bounded range reader: loads only dimensions first, then clips via
 * getAbsoluteResizedRange (or anchor+getResizedRange) BEFORE transferring values/formulas
 * across the Office.js IPC bridge. Prevents OOM/payload errors on 100k-row sheets or Ctrl+A.
 */
async function readBoundedRange(ctx, range, { maxRows = 70, maxCols = 22 } = {}) {
  range.load('address,rowCount,columnCount');
  await ctx.sync();
  if (range.isNullObject) return null;

  const totalRows = Math.max(1, range.rowCount || 1);
  const totalCols = Math.max(1, range.columnCount || 1);
  const clipRows = Math.min(totalRows, maxRows);
  const clipCols = Math.min(totalCols, maxCols);
  const clipped = totalRows > clipRows || totalCols > clipCols;

  let target = range;
  if (clipped) {
    if (typeof range.getAbsoluteResizedRange === 'function') {
      target = range.getAbsoluteResizedRange(clipRows, clipCols);
    } else if (typeof range.getCell === 'function') {
      const topLeft = range.getCell(0, 0);
      target = (clipRows === 1 && clipCols === 1)
        ? topLeft
        : topLeft.getResizedRange(clipRows - 1, clipCols - 1);
    }
  }

  target.load('address,values,formulas');
  await ctx.sync();
  return {
    address: range.address,
    clippedAddress: target.address || range.address,
    values: target.values || [],
    formulas: target.formulas || [],
    totalRows,
    totalCols,
    clipped,
  };
}

async function readUsedRangeSafe(ctx, sheet, limits = { maxRows: 70, maxCols: 22 }) {
  if (typeof sheet.getUsedRangeOrNullObject === 'function') {
    const used = sheet.getUsedRangeOrNullObject(true);
    return readBoundedRange(ctx, used, limits);
  }
  try {
    const used = sheet.getUsedRange(true);
    return await readBoundedRange(ctx, used, limits);
  } catch {
    return null;
  }
}

async function readSelectedRangeSafe(ctx) {
  try {
    const sel = ctx.workbook.getSelectedRange();
    return await readBoundedRange(ctx, sel, { maxRows: 40, maxCols: 16 });
  } catch {
    // Fallback for multi-area (Cmd+Click) selections where getSelectedRange() throws InvalidSelection
    try {
      if (typeof ctx.workbook.getSelectedRanges === 'function') {
        const areas = ctx.workbook.getSelectedRanges().areas;
        const first = areas.getItemAt(0);
        return await readBoundedRange(ctx, first, { maxRows: 40, maxCols: 16 });
      }
    } catch {}
    return null;
  }
}

export function sanitizeSheetName(raw, existingNames = []) {
  let clean = String(raw || 'Summary')
    .replace(/[:\\/?*\[\]]/g, ' ')
    .replace(/^'+|'+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 31);
  if (!clean) clean = 'Summary';

  const existingLower = new Set((existingNames || []).map(n => String(n).toLowerCase()));
  if (!existingLower.has(clean.toLowerCase())) return clean;

  for (let n = 2; n <= 99; n++) {
    const suffix = ` (${n})`;
    const candidate = `${clean.slice(0, 31 - suffix.length)}${suffix}`;
    if (!existingLower.has(candidate.toLowerCase())) return candidate;
  }
  return `${clean.slice(0, 26)}_${Date.now().toString(36).slice(-4)}`;
}

export function createExcelAdapter(env = globalThis) {
  const getExcel = () => env.Excel;
  const getOffice = () => env.Office;

  return {
    id: 'excel',
    label: 'Excel',
    noun: 'workbook',
    placeholder: 'Ask about formulas, audit numbers, or type / for Excel actions…',
    quickActions: EXCEL_QUICK_ACTIONS,
    secondaryToggle: {
      key: 'highlightEdits',
      label: 'Highlight',
      titleOn: 'Highlight edits ON (modified cells are tinted light yellow for review)',
      titleOff: 'Highlight edits OFF (click to tint AI-modified cells for easy review)',
    },
    modeLabels: {
      auto: 'Context: Auto (active sheet + selection)',
      selection: 'Context: Selected range only',
      document: 'Context: Full workbook',
    },

    formatContextBadge(ctx, contextMode) {
      const meta = ctx.hostContext || {};
      const selAddr = meta.selectionAddress || '';
      const activeSheet = meta.activeSheet || 'Sheet1';
      const sheetCount = Array.isArray(meta.sheets) ? meta.sheets.length : 0;
      const errCount = Array.isArray(meta.errors) ? meta.errors.length : 0;
      const errSuffix = errCount ? ` · ${errCount} error${errCount === 1 ? '' : 's'}` : '';

      if (contextMode === 'selection') {
        return {
          badge: selAddr ? stripSheetFromAddress(selAddr) : 'Selection only',
          preview: selAddr ? `${selAddr} selected` : 'Select cells in Excel',
          hasSelection: Boolean(selAddr),
        };
      }
      if (contextMode === 'document') {
        return {
          badge: sheetCount ? `Workbook · ${sheetCount} tab${sheetCount === 1 ? '' : 's'}` : 'Workbook',
          preview: `${activeSheet} active${errSuffix}`,
          hasSelection: false,
        };
      }
      if (selAddr) {
        const shortSel = stripSheetFromAddress(selAddr);
        return {
          badge: `${activeSheet}!${shortSel}`,
          preview: meta.selectionFormula
            ? `${meta.selectionFormula}${errSuffix}`
            : `${meta.selectionSummary || 'Range attached'}${errSuffix}`,
          hasSelection: true,
        };
      }
      return {
        badge: activeSheet,
        preview: sheetCount ? `${sheetCount} worksheet${sheetCount === 1 ? '' : 's'} attached${errSuffix}` : 'Workbook ready',
        hasSelection: false,
      };
    },

    getStarters(ctx) {
      const meta = ctx.hostContext || {};
      const hasFormula = Boolean(meta.selectionFormula);
      const errCount = Array.isArray(meta.errors) ? meta.errors.length : 0;
      const selAddr = meta.selectionAddress ? stripSheetFromAddress(meta.selectionAddress) : '';

      const starters = hasFormula
        ? [
            { title: `Explain formula in ${selAddr || 'selection'}`, sub: 'Trace dependencies and logic cell by cell', prompt: 'Explain the formula in the selected cell(s) step-by-step, cite every cell it depends on, and note any edge cases.' },
            { title: 'Audit this workbook', sub: errCount ? `${errCount} formula error${errCount === 1 ? '' : 's'} detected` : 'Check for broken links, errors, and overrides', prompt: 'Audit this workbook for formula errors, hardcoded overrides, and structural issues. Cite exact cells before editing.' },
            { title: 'Improve or simplify formula', sub: 'Cleaner formula with error handling', prompt: 'Review the selected formula and propose a cleaner, safer version (e.g. with IFERROR or dynamic references) using set_formula.' },
            { title: 'Build a sensitivity or summary table', sub: 'Dynamic formula-driven view', prompt: 'Summarize the key metrics on this sheet and propose a clean formula-driven summary table.' },
          ]
        : [
            { title: 'Audit this workbook', sub: errCount ? `${errCount} formula error${errCount === 1 ? '' : 's'} spotted` : 'Find errors, hardcoded cells, and risks', prompt: 'Audit this workbook for formula errors, hardcoded overrides, and inconsistencies. Cite exact cells in your findings.' },
            { title: 'Summarize key numbers', sub: 'Metrics, trends, and outliers with cell citations', prompt: 'Summarize the key numbers, trends, and outliers in this sheet. Cite exact cells for every figure.' },
            { title: 'Chart this data', sub: 'Create a native Excel chart directly on the sheet', prompt: 'Recommend the best chart for this dataset and propose creating it with add_chart.' },
            { title: 'Format & polish selection', sub: 'Headers, number formats, alignment, auto-fit', prompt: 'Clean up the formatting of the selected table with bold headers, proper number formats, and auto-fitted columns using format_cells.' },
          ];

      const activeSheet = meta.activeSheet || 'Sheet1';
      const subtitle = selAddr
        ? `Selected ${activeSheet}!${selAddr}${errCount ? ` · ${errCount} formula error${errCount === 1 ? '' : 's'}` : ''} · Ask anything or edit directly.`
        : `Reading ${activeSheet} · Ask about formulas, audit numbers, or select a range.`;

      return { heading: 'What are we analyzing?', subtitle, starters };
    },

    async readContext() {
      const Excel = getExcel();
      if (!Excel || !Excel.run) return null;

      let result = null;
      await Excel.run(async ctx => {
        const sheetsCol = ctx.workbook.worksheets;
        const activeSheet = sheetsCol.getActiveWorksheet();

        sheetsCol.load('items/name,items/position');
        activeSheet.load('name');
        await ctx.sync();

        const activeName = activeSheet.name;
        const sheetsMeta = (sheetsCol.items || []).map(s => ({
          name: s.name,
          active: s.name === activeName,
        }));

        const allErrors = [];
        const docSections = [];

        // Active sheet (two-phase bounded read: max 70x22 cells transferred over bridge)
        const activeUsed = await readUsedRangeSafe(ctx, activeSheet, { maxRows: 70, maxCols: 22 });
        if (activeUsed) {
          const snap = formatSheetSnapshot({
            sheetName: activeName,
            address: activeUsed.address,
            values: activeUsed.values,
            formulas: activeUsed.formulas,
            maxRows: 70,
            maxCols: 22,
          });
          allErrors.push(...snap.errors);
          const clipNote = activeUsed.clipped
            ? `\n(showing top ${activeUsed.values.length}×${activeUsed.values[0]?.length || 0} of ${activeUsed.totalRows}×${activeUsed.totalCols} used cells)`
            : '';
          docSections.push(`### Active Sheet: ${activeName} (${stripSheetFromAddress(activeUsed.address)})\n${snap.text}${clipNote}`);
        } else {
          docSections.push(`### Active Sheet: ${activeName}\n(empty worksheet)`);
        }

        // Secondary sheets (up to 6 tabs, bounded 30x14 transfer)
        const otherSheets = (sheetsCol.items || []).filter(s => s.name !== activeName).slice(0, 6);
        for (const sheet of otherSheets) {
          const used = await readUsedRangeSafe(ctx, sheet, { maxRows: 30, maxCols: 14 });
          if (!used) continue;
          const snap = formatSheetSnapshot({
            sheetName: sheet.name,
            address: used.address,
            values: used.values,
            formulas: used.formulas,
            maxRows: 30,
            maxCols: 14,
          });
          allErrors.push(...snap.errors);
          const clipNote = used.clipped
            ? `\n(showing top ${used.values.length}×${used.values[0]?.length || 0} of ${used.totalRows}×${used.totalCols} used cells)`
            : '';
          docSections.push(`### Sheet: ${sheet.name} (${stripSheetFromAddress(used.address)})\n${snap.text}${clipNote}`);
        }

        // Bounded selection read (handles Ctrl+A full-column selections & multi-area selections)
        const selData = await readSelectedRangeSafe(ctx);
        const selSnap = selData
          ? formatSheetSnapshot({
              sheetName: activeName,
              address: selData.address,
              values: selData.values,
              formulas: selData.formulas,
              maxRows: 40,
              maxCols: 16,
            })
          : { text: '' };

        const topLeftFormula = selData?.formulas?.[0]?.[0];
        const selectionFormula = typeof topLeftFormula === 'string' && topLeftFormula.startsWith('=')
          ? topLeftFormula
          : '';
        const selectionSummary = !selData
          ? 'No range selected'
          : selData.totalRows === 1 && selData.totalCols === 1
            ? (selData.values?.[0]?.[0] !== '' && selData.values?.[0]?.[0] !== undefined
                ? `Value: ${selData.values[0][0]}`
                : 'Empty cell')
            : `${selData.totalRows}×${selData.totalCols} cells`;
        const selectionClipped = selData?.clipped
          ? `clipped to top ${selData.values.length}×${selData.values[0]?.length || 0} of ${selData.totalRows}×${selData.totalCols} selected cells`
          : '';

        // Open comments on active sheet (ExcelApi 1.10+)
        let comments = [];
        try {
          if (activeSheet.comments && typeof activeSheet.comments.load === 'function') {
            activeSheet.comments.load('items/id,items/content,items/authorName,items/resolved');
            await ctx.sync();
            const openItems = (activeSheet.comments.items || []).filter(c => !c.resolved).slice(0, 20);
            const locs = openItems.map(c => {
              const loc = c.getLocation();
              loc.load('address');
              return loc;
            });
            if (locs.length) await ctx.sync();
            comments = openItems.map((c, idx) => ({
              id: c.id,
              author: c.authorName || 'Reviewer',
              content: c.content || '',
              anchor: locs[idx]?.address || '',
            }));
          }
        } catch {
          // Comments API unavailable on older Excel builds.
        }

        result = {
          selectionText: selSnap.text,
          documentText: docSections.join('\n\n'),
          comments,
          hostContext: {
            activeSheet: activeName,
            sheets: sheetsMeta,
            selectionAddress: selData?.address || '',
            selectionFormula,
            selectionSummary,
            selectionClipped,
            errors: allErrors.slice(0, 40),
          },
        };
      });
      return result;
    },

    watchSelection(onChange, onFallback) {
      const Office = getOffice();
      const Excel = getExcel();
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange, result => {
          if (result.status !== Office.AsyncResultStatus.Succeeded && onFallback) {
            onFallback('Auto-selection unavailable here. Use the refresh icon after selecting.');
          }
        });
      } catch {
        if (onFallback) onFallback('Auto-selection unavailable here. Use the refresh icon after selecting.');
      }

      // Also refresh when switching worksheet tabs
      if (Excel && Excel.run) {
        Excel.run(async ctx => {
          if (ctx.workbook.worksheets.onActivated && typeof ctx.workbook.worksheets.onActivated.add === 'function') {
            ctx.workbook.worksheets.onActivated.add(() => onChange());
            await ctx.sync();
          }
        }).catch(() => {});
      }
    },

    async toggleSecondary(enabled) {
      return {
        ok: true,
        value: enabled,
        message: enabled
          ? 'Highlight edits ON · modified cells will be tinted yellow'
          : 'Highlight edits OFF',
      };
    },

    async navigateToRef(ref) {
      const Excel = getExcel();
      if (!Excel || !Excel.run) {
        return { ok: false, message: 'Open inside Microsoft Excel to jump to cells.' };
      }
      const parsed = parseExcelRef(ref);
      if (!parsed) return { ok: false, message: `Invalid Excel reference: ${ref}` };

      await Excel.run(async ctx => {
        const sheet = await resolveSheet(ctx, parsed.sheet);
        if (parsed.sheet && typeof sheet.activate === 'function') {
          sheet.activate();
        }
        const range = sheet.getRange(parsed.range);
        range.select();
        await ctx.sync();
      });

      const label = formatExcelAddress(parsed.sheet, parsed.range);
      return { ok: true, message: `Jumped to ${label}` };
    },

    async applyEdit(edit, { highlightEdits = false } = {}) {
      const Excel = getExcel();
      if (!Excel || !Excel.run) {
        return { ok: false, message: 'Open inside Microsoft Excel to apply edits directly.' };
      }

      const action = edit.action;
      const attrs = edit.attrs || {};
      const content = (edit.content || '').trim();
      let summaryMsg = 'Applied to Excel · Cmd+Z to undo';

      await Excel.run(async ctx => {
        if (action === 'add_sheet') {
          const sheetsCol = ctx.workbook.worksheets;
          sheetsCol.load('items/name');
          await ctx.sync();
          const existingNames = (sheetsCol.items || []).map(s => s.name);
          const sheetName = sanitizeSheetName(attrs.name || 'Summary', existingNames);
          const sheet = ctx.workbook.worksheets.add(sheetName);
          if (typeof sheet.activate === 'function') sheet.activate();

          const grid = parseCellGrid(content);
          if (grid.length > 0 && grid[0].length > 0) {
            const base = sheet.getRange('A1');
            const target = (grid.length === 1 && grid[0].length === 1)
              ? base
              : base.getResizedRange(grid.length - 1, grid[0].length - 1);
            target.formulas = grid;
            if (highlightEdits) target.format.fill.color = HIGHLIGHT_FILL;
            if (target.format?.autofitColumns) target.format.autofitColumns();
          }
          await ctx.sync();
          summaryMsg = `Created sheet "${sheetName}" · Cmd+Z to undo`;
          return;
        }

        const rawTarget = attrs.range || attrs.cell || '';
        const parsedRef = parseExcelRef(rawTarget);
        const sheetName = parsedRef?.sheet || attrs.sheet || '';
        const sheet = await resolveSheet(ctx, sheetName);

        if (action === 'set_cells') {
          const grid = parseCellGrid(content);
          if (!grid.length || !grid[0].length) {
            throw new Error('Could not parse cell values to write.');
          }
          const isSingle = grid.length === 1 && grid[0].length === 1;
          let target;
          let targetLabel;
          if (parsedRef) {
            const base = sheet.getRange(parsedRef.startCell);
            target = isSingle ? base : base.getResizedRange(grid.length - 1, grid[0].length - 1);
            targetLabel = formatExcelAddress(sheetName, parsedRef.range);
          } else {
            const sel = ctx.workbook.getSelectedRange();
            sel.load('address');
            await ctx.sync();
            const selParsed = parseExcelRef(sel.address);
            const anchor = selParsed ? selParsed.startCell : 'A1';
            const base = sheet.getRange(anchor);
            target = isSingle ? base : base.getResizedRange(grid.length - 1, grid[0].length - 1);
            targetLabel = formatExcelAddress(sheetName, anchor);
          }
          target.formulas = grid;
          if (attrs.number_format) {
            target.numberFormat = Array.from({ length: grid.length }, () => Array(grid[0].length).fill(attrs.number_format));
          }
          if (highlightEdits) target.format.fill.color = HIGHLIGHT_FILL;
          await ctx.sync();
          summaryMsg = `Wrote ${grid.length}×${grid[0].length} cells at ${targetLabel}${highlightEdits ? ' (highlighted)' : ''} · Cmd+Z to undo`;
          return;
        }

        if (action === 'set_formula') {
          if (!content) throw new Error('Formula is empty.');
          let formula = content.replace(/^```[a-z]*\s*\n?/i, '').replace(/\n?```\s*$/, '').trim();
          if (formula && !formula.startsWith('=')) formula = `=${formula}`;
          const target = parsedRef ? sheet.getRange(parsedRef.range) : ctx.workbook.getSelectedRange();
          target.load('rowCount,columnCount,address');
          await ctx.sync();
          const rows = Math.max(1, target.rowCount || 1);
          const cols = Math.max(1, target.columnCount || 1);
          target.formulas = Array.from({ length: rows }, () => Array(cols).fill(formula));
          if (attrs.number_format) {
            target.numberFormat = Array.from({ length: rows }, () => Array(cols).fill(attrs.number_format));
          }
          if (highlightEdits) target.format.fill.color = HIGHLIGHT_FILL;
          await ctx.sync();
          summaryMsg = `Set formula on ${stripSheetFromAddress(target.address)}${highlightEdits ? ' (highlighted)' : ''} · Cmd+Z to undo`;
          return;
        }

        if (action === 'format_cells') {
          const spec = parseFormatSpec(content);
          const target = parsedRef ? sheet.getRange(parsedRef.range) : ctx.workbook.getSelectedRange();
          target.load('rowCount,columnCount,address');
          await ctx.sync();
          if (typeof spec.bold === 'boolean') target.format.font.bold = spec.bold;
          if (typeof spec.italic === 'boolean') target.format.font.italic = spec.italic;
          if (typeof spec.fontSize === 'number') target.format.font.size = spec.fontSize;
          if (typeof spec.fontColor === 'string' && spec.fontColor) target.format.font.color = spec.fontColor;
          if (typeof spec.fill === 'string' && spec.fill) target.format.fill.color = spec.fill;
          if (typeof spec.horizontalAlignment === 'string' && spec.horizontalAlignment) {
            target.format.horizontalAlignment = spec.horizontalAlignment;
          }
          if (typeof spec.wrapText === 'boolean') target.format.wrapText = spec.wrapText;
          if (typeof spec.numberFormat === 'string' && spec.numberFormat) {
            const rows = Math.max(1, target.rowCount || 1);
            const cols = Math.max(1, target.columnCount || 1);
            target.numberFormat = Array.from({ length: rows }, () => Array(cols).fill(spec.numberFormat));
          }
          if (spec.autoFit && typeof target.format.autofitColumns === 'function') {
            target.format.autofitColumns();
          }
          await ctx.sync();
          summaryMsg = `Formatted ${stripSheetFromAddress(target.address)} · Cmd+Z to undo`;
          return;
        }

        if (action === 'add_chart') {
          const chartType = (attrs.chart_type || 'ColumnClustered').trim();
          const dataRange = parsedRef ? sheet.getRange(parsedRef.range) : ctx.workbook.getSelectedRange();
          const seriesBy = Excel.ChartSeriesBy ? Excel.ChartSeriesBy.auto : 'Auto';
          const chart = sheet.charts.add(chartType, dataRange, seriesBy);
          const title = (attrs.title || content || '').trim();
          if (title && chart.title) {
            chart.title.text = title;
            chart.title.visible = true;
          }
          await ctx.sync();
          summaryMsg = `Added ${chartType} chart${title ? ` "${title}"` : ''} · Cmd+Z to undo`;
          return;
        }

        if (action === 'sort_range') {
          const target = parsedRef ? sheet.getRange(parsedRef.range) : ctx.workbook.getSelectedRange();
          target.load('address,columnIndex');
          await ctx.sync();
          const rawCol = String(attrs.column ?? '0').trim();
          let key = 0;
          if (/^[A-Za-z]+$/.test(rawCol)) {
            const absIdx = colToIndex(rawCol);
            const startIdx = typeof target.columnIndex === 'number'
              ? target.columnIndex
              : (parseExcelRef(target.address)?.startCol ?? 0);
            key = Math.max(0, absIdx - startIdx);
          } else {
            key = Math.max(0, parseInt(rawCol, 10) || 0);
          }
          const ascending = String(attrs.order || 'asc').toLowerCase() !== 'desc';
          const hasHeaders = String(attrs.has_headers ?? 'true').toLowerCase() !== 'false';
          target.sort.apply([{ key, ascending }], hasHeaders);
          await ctx.sync();
          summaryMsg = `Sorted ${stripSheetFromAddress(target.address)} (${ascending ? 'A→Z' : 'Z→A'}) · Cmd+Z to undo`;
          return;
        }

        if (action === 'excel_comment') {
          if (!content) throw new Error('Comment text is empty.');
          const target = parsedRef ? sheet.getRange(parsedRef.startCell) : ctx.workbook.getSelectedRange();
          target.load('address');
          await ctx.sync();
          if (sheet.comments && typeof sheet.comments.add === 'function') {
            sheet.comments.add(target, content);
          } else if (ctx.workbook.comments && typeof ctx.workbook.comments.add === 'function') {
            ctx.workbook.comments.add(target.address, content);
          } else {
            throw new Error('This version of Excel does not support adding comments from add-ins.');
          }
          await ctx.sync();
          summaryMsg = `Added comment on ${stripSheetFromAddress(target.address)}`;
          return;
        }

        if (action === 'clear_range') {
          const target = parsedRef ? sheet.getRange(parsedRef.range) : ctx.workbook.getSelectedRange();
          target.load('address');
          await ctx.sync();
          const clearMode = Excel.ClearApplyTo ? Excel.ClearApplyTo.contents : 'Contents';
          target.clear(clearMode);
          await ctx.sync();
          summaryMsg = `Cleared ${stripSheetFromAddress(target.address)} · Cmd+Z to undo`;
          return;
        }

        throw new Error(`Unsupported Excel action: ${action}`);
      });

      return { ok: true, message: summaryMsg };
    },
  };
}
