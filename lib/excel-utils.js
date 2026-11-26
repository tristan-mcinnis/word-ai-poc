/* Pure Excel helpers for A1 addresses, formula/value grids, error detection,
   and compact workbook serialization. Runs in both browser ES modules and Node. */

export const EXCEL_ERRORS = new Set([
  '#REF!',
  '#DIV/0!',
  '#VALUE!',
  '#NAME?',
  '#N/A',
  '#NUM!',
  '#NULL!',
  '#SPILL!',
  '#CALC!',
]);

export function indexToCol(index) {
  let n = Number(index);
  if (!Number.isFinite(n) || n < 0) return 'A';
  let out = '';
  n = Math.floor(n);
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

export function colToIndex(col) {
  const clean = String(col || '').trim().toUpperCase().replace(/[^A-Z]/g, '');
  if (!clean) return -1;
  let result = 0;
  for (let i = 0; i < clean.length; i++) {
    result = result * 26 + (clean.charCodeAt(i) - 64);
  }
  return result - 1;
}

/**
 * Parse an Excel reference like "B2", "B2:D10", "Sheet2!A1:C5", or "'Q1 P&L'!$B$2:$D$10".
 * Returns { sheet, range, startCell, endCell, startCol, startRow, endCol, endRow } or null.
 */
export function parseExcelRef(raw) {
  const input = String(raw || '').trim();
  if (!input) return null;

  let sheet = '';
  let addr = input;
  const bang = input.lastIndexOf('!');
  if (bang > 0) {
    sheet = input.slice(0, bang).trim();
    if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
    addr = input.slice(bang + 1).trim();
  }

  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(addr);
  if (!m) return null;

  const c1 = m[1].toUpperCase();
  const r1 = parseInt(m[2], 10);
  const c2 = (m[3] || m[1]).toUpperCase();
  const r2 = parseInt(m[4] || m[2], 10);
  if (!r1 || !r2) return null;

  const startCell = `${c1}${r1}`;
  const endCell = `${c2}${r2}`;
  const range = m[3] ? `${startCell}:${endCell}` : startCell;

  return {
    sheet,
    range,
    startCell,
    endCell,
    startCol: colToIndex(c1),
    startRow: r1 - 1,
    endCol: colToIndex(c2),
    endRow: r2 - 1,
    isRange: Boolean(m[3]),
  };
}

export function formatExcelAddress(sheet, range) {
  const cleanRange = String(range || '').trim();
  if (!sheet) return cleanRange;
  const needsQuotes = /[^A-Za-z0-9_]/.test(sheet);
  const safeSheet = needsQuotes ? `'${sheet.replace(/'/g, "''")}'` : sheet;
  return `${safeSheet}!${cleanRange}`;
}

export function isExcelErrorValue(val) {
  return typeof val === 'string' && EXCEL_ERRORS.has(val.trim().toUpperCase());
}

/**
 * Coerce a string cell token from LLM output into an appropriate Excel value:
 * - Formulas starting with "=" stay strings
 * - Explicit quoted strings stay strings
 * - Clean numbers become JS numbers
 * - Booleans become JS booleans
 */
export function coerceCellValue(raw) {
  if (raw === null || raw === undefined) return '';
  if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
  const s = String(raw).trim();
  if (!s) return '';
  if (s.startsWith('=')) return s;
  if (/^true$/i.test(s)) return true;
  if (/^false$/i.test(s)) return false;
  // Preserve leading-zero codes like "00123" as strings
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(s)) {
    const num = Number(s);
    if (Number.isFinite(num)) return num;
  }
  return s;
}

/**
 * Parse cell grid text from an <edit action="set_cells"> or <edit action="add_sheet"> body.
 * Supports:
 * 1) JSON 2D array (e.g. [["Item","Q1"],["Rev","=SUM(B2:B4)"]]) or 1D array
 * 2) Markdown pipe tables
 * 3) Tab-separated rows (TSV)
 * 4) Comma-separated rows (CSV) when no tabs are present
 * Always returns a rectangular 2D array of coerced values.
 */
export function parseCellGrid(raw) {
  const text = String(raw || '').trim();
  if (!text) return [];

  // Strip optional markdown code fences around JSON or TSV
  const unwrapped = text.replace(/^```(?:json|tsv|csv)?\s*\n?/i, '').replace(/\n?```\s*$/, '').trim();

  // 1. JSON array
  if (unwrapped.startsWith('[')) {
    try {
      const parsed = JSON.parse(unwrapped);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const rows = Array.isArray(parsed[0])
          ? parsed.map(r => (Array.isArray(r) ? r.map(coerceCellValue) : [coerceCellValue(r)]))
          : [parsed.map(coerceCellValue)];
        return normalizeRect(rows);
      }
    } catch {
      // Fall through to line-based parsing
    }
  }

  const lines = unwrapped.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (!lines.length) return [];

  // 2. Markdown pipe table
  if (lines.every(l => l.trim().includes('|'))) {
    const rows = [];
    for (const line of lines) {
      const stripped = line.trim().replace(/^\|/, '').replace(/\|$/, '');
      const cells = stripped.split('|').map(c => c.trim());
      // Skip separator rows like |---|---|
      if (cells.every(c => /^:?-{2,}:?$/.test(c))) continue;
      rows.push(cells.map(coerceCellValue));
    }
    if (rows.length) return normalizeRect(rows);
  }

  // 3. TSV
  if (lines.some(l => l.includes('\t'))) {
    return normalizeRect(lines.map(l => l.split('\t').map(coerceCellValue)));
  }

  // 4. Simple CSV (only if consistent commas and not a single formula with commas)
  if (lines.length >= 1 && lines.some(l => !l.trim().startsWith('=') && l.includes(','))) {
    return normalizeRect(lines.map(splitCsvLine).map(row => row.map(coerceCellValue)));
  }

  // Single-column values
  return normalizeRect(lines.map(l => [coerceCellValue(l)]));
}

function splitCsvLine(line) {
  const cells = [];
  let cur = '';
  let inQuotes = false;
  let parenDepth = 0;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (!inQuotes && ch === '(') {
      parenDepth++;
      cur += ch;
    } else if (!inQuotes && ch === ')' && parenDepth > 0) {
      parenDepth--;
      cur += ch;
    } else if (!inQuotes && parenDepth === 0 && ch === ',') {
      cells.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  cells.push(cur);
  return cells;
}

function normalizeRect(rows) {
  const width = rows.reduce((max, r) => Math.max(max, r.length), 0);
  if (width === 0) return [];
  return rows.map(r => {
    const copy = r.slice();
    while (copy.length < width) copy.push('');
    return copy;
  });
}

/**
 * Format a worksheet's values and formulas into a clean, token-efficient markdown-like grid
 * with explicit column letters & row numbers so the model can cite exact cells and trace formulas.
 */
export function formatSheetSnapshot({
  sheetName,
  address,
  values = [],
  formulas = [],
  maxRows = 60,
  maxCols = 20,
}) {
  const parsed = parseExcelRef(address || 'A1');
  const baseRow = parsed ? parsed.startRow : 0;
  const baseCol = parsed ? parsed.startCol : 0;

  const totalRows = values.length;
  const totalCols = values.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 0), 0);
  if (!totalRows || !totalCols) {
    return { text: `Sheet "${sheetName}" is empty.`, errors: [], cellCount: 0 };
  }

  const rowLimit = Math.min(totalRows, maxRows);
  const colLimit = Math.min(totalCols, maxCols);
  const errors = [];
  let nonEmpty = 0;

  // Scan full sheet for formula errors even beyond rowLimit
  for (let r = 0; r < totalRows; r++) {
    const row = values[r] || [];
    const fRow = formulas[r] || [];
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v !== '' && v !== null && v !== undefined) nonEmpty++;
      if (isExcelErrorValue(v)) {
        const ref = `${indexToCol(baseCol + c)}${baseRow + r + 1}`;
        const f = typeof fRow[c] === 'string' && fRow[c].startsWith('=') ? ` (${fRow[c]})` : '';
        errors.push(`${formatExcelAddress(sheetName, ref)}: ${v}${f}`);
      }
    }
  }

  const colHeaders = [];
  for (let c = 0; c < colLimit; c++) {
    colHeaders.push(indexToCol(baseCol + c));
  }

  const lines = [];
  lines.push(`| Row | ${colHeaders.join(' | ')} |`);
  lines.push(`|---|${colHeaders.map(() => '---').join('|')}|`);

  for (let r = 0; r < rowLimit; r++) {
    const row = values[r] || [];
    const fRow = formulas[r] || [];
    const cells = [];
    let rowHasContent = false;
    for (let c = 0; c < colLimit; c++) {
      const val = row[c];
      const formula = fRow[c];
      let cellStr = '';
      if (typeof formula === 'string' && formula.startsWith('=')) {
        const displayVal = val === '' || val === null || val === undefined ? '' : ` [${val}]`;
        cellStr = `${formula}${displayVal}`;
      } else if (val !== '' && val !== null && val !== undefined) {
        cellStr = String(val);
      }
      if (cellStr) rowHasContent = true;
      cells.push(cellStr.replace(/\|/g, '\\|').replace(/\r?\n/g, ' ').slice(0, 80));
    }
    if (rowHasContent) {
      lines.push(`| ${baseRow + r + 1} | ${cells.join(' | ')} |`);
    }
  }

  if (totalRows > rowLimit || totalCols > colLimit) {
    lines.push(`... (showing ${rowLimit}x${colLimit} of ${totalRows}x${totalCols} used cells in ${address})`);
  }

  return {
    text: lines.join('\n'),
    errors,
    cellCount: nonEmpty,
    totalRows,
    totalCols,
  };
}

/**
 * Parse JSON formatting options safely for format_cells.
 */
export function parseFormatSpec(raw) {
  const text = String(raw || '').trim()
    .replace(/^```(?:json)?\s*\n?/i, '')
    .replace(/\n?```\s*$/, '')
    .trim();
  if (!text) return {};
  try {
    const obj = JSON.parse(text);
    return obj && typeof obj === 'object' && !Array.isArray(obj) ? obj : {};
  } catch {
    return {};
  }
}

/**
 * Compute cell-level change summary between existing 2D values and proposed 2D grid.
 */
export function diffCellGrids(oldGrid, newGrid, startRef = 'A1') {
  const parsed = parseExcelRef(startRef) || { startRow: 0, startCol: 0 };
  const changes = [];
  for (let r = 0; r < newGrid.length; r++) {
    for (let c = 0; c < (newGrid[r] || []).length; c++) {
      const before = oldGrid?.[r]?.[c] ?? '';
      const after = newGrid[r][c] ?? '';
      if (String(before) !== String(after)) {
        changes.push({
          cell: `${indexToCol(parsed.startCol + c)}${parsed.startRow + r + 1}`,
          before,
          after,
        });
      }
    }
  }
  return changes;
}
