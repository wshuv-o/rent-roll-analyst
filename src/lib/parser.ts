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
 * 1. Uses suite_id column to detect new tenants
 * 2. For each row, collects all columns within each group span as {label → value}
 * 3. Continuation rows append entries to the current tenant's groups
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

  console.log('[PARSER] groupSpans:', JSON.stringify(groupSpans));
  console.log('[PARSER] columnLabels:', JSON.stringify(columnLabels));

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

    // Addon space — attach as continuation with a note
    if (current && (rowStr.includes("add'l space") || rowStr.includes('addl space') || rowStr.includes('additional space') ||
      (addon_space_patterns.length > 0 && addon_space_patterns.some(p => {
        try { return new RegExp(p, 'i').test(rowStr); } catch { return rowStr.includes(p.toLowerCase()); }
      })))) {
      collectGroupValues(row, groupSpans, columnLabels, current);
      current.notes += (current.notes ? '; ' : '') + `Add'l space row ${i + 1}`;
      continue;
    }

    // NEW TENANT — suite_id column has a value
    if (suiteVal) {
      if (tenantVal.toLowerCase().startsWith('psf') || tenantVal.startsWith('(')) {
        // sub-line — treat as continuation
      } else {
        if (current) tenants.push(current);

        current = {
          suite_id: suiteVal,
          tenant_name: tenantVal,
          groups: {},
          notes: '',
        };

        collectGroupValues(row, groupSpans, columnLabels, current);
        continue;
      }
    }

    // CONTINUATION ROW
    if (current) {
      collectGroupValues(row, groupSpans, columnLabels, current);
    }
  }

  if (current) tenants.push(current);

  log('system', `${tenants.length} tenant blocks found.`);
  console.log('[PARSER] Found', tenants.length, 'tenants');

  return tenants;
}

/** Collect all columns within each group span as a {label→value} entry */
function collectGroupValues(
  row: (string | number | null)[],
  groupSpans: GroupSpan[],
  columnLabels: Record<number, string>,
  tenant: TenantObject
) {
  for (const span of groupSpans) {
    // Skip identity group — suite_id and tenant_name are already top-level
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

    if (hasValue) {
      if (!tenant.groups[span.groupId]) tenant.groups[span.groupId] = [];
      tenant.groups[span.groupId].push(entry);
    }
  }
}
