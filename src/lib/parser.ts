import type { ParsingInstruction, TenantObject, GroupSpan } from './types';

function indexToColLetter(idx: number): string {
  let letter = '';
  let n = idx;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

function colLetterToIndex(letter: string): number {
  if (!letter) return -1;
  const upper = letter.toUpperCase().trim().replace(/[^A-Z]/g, '');
  if (!upper) return -1;
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

function getCellValue(row: (string | number | null)[], colIdx: number): string {
  if (colIdx < 0 || colIdx >= row.length) return '';
  const val = row[colIdx];
  return val !== null && val !== undefined ? String(val).trim() : '';
}

/**
 * Group-based parser.
 * - Scalar groups: first non-empty values win (lease dates, space, base rent)
 * - Collection groups: every row with data accumulates (charges, future rent)
 */
export function parseSheet(
  data: (string | number | null)[][],
  instruction: ParsingInstruction,
  groupSpans: GroupSpan[],
  columnLabels: Record<number, string>,
  addLog?: (type: 'system' | 'flag', msg: string) => void
): TenantObject[] {
  const { column_map: cm, data_starts_at_row, skip_row_patterns, addon_space_patterns } = instruction;
  const startRow = (data_starts_at_row ?? 1) - 1;
  const suiteColIdx = colLetterToIndex(cm.suite_id);
  const tenantColIdx = colLetterToIndex(cm.tenant_name);

  const log = addLog || (() => {});
  log('system', `Parser: start_row=${data_starts_at_row}, suite_col=${cm.suite_id}(${suiteColIdx}), ${data.length} total rows, ${groupSpans.length} groups`);

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
          scalars: {},
          collections: {},
          notes: '',
        };

        collectRow(row, groupSpans, columnLabels, current);
        continue;
      }
    }

    // CONTINUATION ROW
    if (current) {
      collectRow(row, groupSpans, columnLabels, current);
    }
  }

  if (current) tenants.push(current);

  log('system', `${tenants.length} tenant blocks found.`);
  return tenants;
}

/** Collect values from a row into the tenant's scalars/collections based on group type */
function collectRow(
  row: (string | number | null)[],
  groupSpans: GroupSpan[],
  columnLabels: Record<number, string>,
  tenant: TenantObject
) {
  for (const span of groupSpans) {
    if (span.groupId === 'identity') continue;

    const entry: Record<string, string | number | null> = {};
    let hasValue = false;

    for (let col = span.startCol; col <= span.endCol; col++) {
      const label = columnLabels[col] || indexToColLetter(col);
      const val = col < row.length ? row[col] : null;
      const cleaned = val !== null && val !== undefined ? String(val).trim() : '';
      entry[label] = cleaned || null;
      if (cleaned) hasValue = true;
    }

    if (!hasValue) continue;

    if (span.collection) {
      // Collection group: accumulate every row
      if (!tenant.collections[span.groupId]) tenant.collections[span.groupId] = [];
      tenant.collections[span.groupId].push(entry);
    } else {
      // Scalar group: only set if not already populated
      if (!tenant.scalars[span.groupId]) {
        tenant.scalars[span.groupId] = entry;
      }
    }
  }
}
