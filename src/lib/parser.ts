import type { ParsingInstruction, TenantObject } from './types';
import { colLetterToIndex, getCellValue } from './col-utils';

/**
 * Raw-row parser.
 * Groups rows into tenant blocks based on the suite_id column.
 * Stores raw Excel rows directly — no data transformation or loss.
 */
export function parseSheet(
  data: (string | number | null)[][],
  instruction: ParsingInstruction,
  addLog?: (type: 'system' | 'flag', msg: string) => void
): TenantObject[] {
  const { column_map: cm, data_starts_at_row, skip_row_patterns } = instruction;
  const startRow = (data_starts_at_row ?? 1) - 1;
  const suiteColIdx = colLetterToIndex(cm.suite_id);
  const tenantColIdx = colLetterToIndex(cm.tenant_name);

  const log = addLog || (() => {});
  log('system', `Parser: start_row=${data_starts_at_row}, suite_col=${cm.suite_id}(${suiteColIdx}), ${data.length} total rows`);

  const tenants: TenantObject[] = [];
  let current: TenantObject | null = null;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === null || c === undefined || String(c).trim() === '')) continue;

    const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();

    // Skip patterns
    if (skip_row_patterns.length > 0 && skip_row_patterns.some(p => {
      try { return new RegExp(p, 'i').test(rowStr); } catch { return rowStr.includes(p.toLowerCase()); }
    })) continue;

    const suiteVal = suiteColIdx >= 0 ? getCellValue(row, suiteColIdx) : '';
    const tenantVal = tenantColIdx >= 0 ? getCellValue(row, tenantColIdx) : '';

    // NEW TENANT — suite_id column has a value
    if (suiteVal) {
      if (current) tenants.push(current);
      current = {
        suite_id: suiteVal,
        tenant_name: tenantVal,
        rawRows: [row],
        notes: '',
      };
      continue;
    }

    // CONTINUATION ROW — append to current tenant
    if (current) {
      current.rawRows.push(row);
    }
  }

  if (current) tenants.push(current);

  log('system', `${tenants.length} tenant blocks found.`);
  return tenants;
}
