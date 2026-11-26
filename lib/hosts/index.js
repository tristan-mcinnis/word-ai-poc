/* Host adapter registry & Office host detection (Word, Excel, PowerPoint). */

import { createWordAdapter } from './word.js';
import { createExcelAdapter } from './excel.js';
import { createPowerPointAdapter } from './powerpoint.js';

export const HOST_IDS = ['word', 'excel', 'powerpoint'];

export function detectOfficeHost(officeInfo, OfficeGlobal = globalThis.Office) {
  const host = officeInfo?.host;
  if (!host) return null;
  if (OfficeGlobal?.HostType) {
    if (host === OfficeGlobal.HostType.Word) return 'word';
    if (host === OfficeGlobal.HostType.Excel) return 'excel';
    if (host === OfficeGlobal.HostType.PowerPoint) return 'powerpoint';
  }
  const lower = String(host).toLowerCase();
  if (lower.includes('word')) return 'word';
  if (lower.includes('excel')) return 'excel';
  if (lower.includes('powerpoint')) return 'powerpoint';
  return null;
}

export function createHostAdapter(hostId, env = globalThis) {
  switch (String(hostId || '').toLowerCase()) {
    case 'excel': return createExcelAdapter(env);
    case 'powerpoint': return createPowerPointAdapter(env);
    default: return createWordAdapter(env);
  }
}
