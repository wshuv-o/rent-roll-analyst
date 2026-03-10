import type { ParsingInstruction, TenantObject } from './types';

function colLetterToIndex(letter: string): number {
  if (!letter) return -1;
  const upper = letter.toUpperCase().trim();
  if (!upper) return -1;
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    index = index * 26 + (code - 64);
  }
  return index - 1;
}

function getCellValue(row: (string | number | null)[], colLetter: string): string {
  const idx = colLetterToIndex(colLetter);
  if (idx < 0 || idx >= row.length) return '';
  const val = row[idx];
  return val !== null && val !== undefined ? String(val).trim() : '';
}

function getNumericValue(row: (string | number | null)[], colLetter: string): number | null {
  const str = getCellValue(row, colLetter);
  if (!str) return null;
  const cleaned = str.replace(/[\s$,]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

function shouldSkipRow(row: (string | number | null)[], patterns: string[]): boolean {
  if (patterns.length === 0) return false;
  const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();
  return patterns.some(p => {
    try {
      return new RegExp(p, 'i').test(rowStr);
    } catch {
      return rowStr.includes(p.toLowerCase());
    }
  });
}

function isAddonSpace(row: (string | number | null)[], patterns: string[]): boolean {
  if (patterns.length === 0) return false;
  const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();
  // Also detect "Add'l Space" pattern directly
  if (rowStr.includes("add'l space") || rowStr.includes('addl space') || rowStr.includes('additional space')) {
    return true;
  }
  return patterns.some(p => {
    try {
      return new RegExp(p, 'i').test(rowStr);
    } catch {
      return rowStr.includes(p.toLowerCase());
    }
  });
}

/**
 * Determines if a row starts a new tenant block.
 * Uses a multi-strategy approach — doesn't rely on exact AI phrasing.
 */
function isNewTenant(
  row: (string | number | null)[],
  _rule: string,
  colMap: ParsingInstruction['column_map']
): boolean {
  const suiteVal = getCellValue(row, colMap.suite_id);
  const tenantVal = getCellValue(row, colMap.tenant_name);

  // Skip rows where both suite and tenant are empty — these are continuation rows
  if (!suiteVal && !tenantVal) return false;

  // Skip summary rows (Total SF, Total PSF, etc.)
  const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();
  if (rowStr.includes('total sf') || rowStr.includes('total psf') || rowStr.includes('total nra')) {
    return false;
  }

  // Skip addon space rows
  if (rowStr.includes("add'l space") || rowStr.includes('addl space') || rowStr.includes('additional space')) {
    return false;
  }

  // Skip PSF indicator rows
  if (tenantVal.toLowerCase().startsWith('psf') || tenantVal.startsWith('(')) {
    return false;
  }

  // A row with a suite ID is a new tenant
  if (suiteVal) return true;

  // A row with a tenant name (and no suite) could be a continuation or a new tenant
  // without a suite — for safety, don't count it as new unless suite is present
  return false;
}

export function parseSheet(
  data: (string | number | null)[][],
  instruction: ParsingInstruction,
  addLog?: (type: 'system' | 'flag', msg: string) => void
): TenantObject[] {
  const { column_map: cm, data_starts_at_row, skip_row_patterns, addon_space_patterns } = instruction;
  const startRow = (data_starts_at_row ?? 1) - 1;

  const log = addLog || (() => {});

  log('system', `Parser config: data starts at row ${data_starts_at_row}, suite_id=${cm.suite_id}, tenant_name=${cm.tenant_name}`);

  const tenants: TenantObject[] = [];
  let currentTenant: TenantObject | null = null;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === null || c === undefined || String(c).trim() === '')) continue;

    // Skip summary/total rows
    if (shouldSkipRow(row, skip_row_patterns)) continue;

    // Check if addon space — attach to current tenant
    if (currentTenant && isAddonSpace(row, addon_space_patterns)) {
      const addonSqft = getNumericValue(row, cm.gla_sqft);
      if (addonSqft !== null && currentTenant.gla_sqft !== null) {
        currentTenant.gla_sqft += addonSqft;
      }
      currentTenant.notes += (currentTenant.notes ? '; ' : '') + `Add'l space: ${addonSqft ?? 'N/A'} SF`;
      continue;
    }

    // Check for new tenant
    if (isNewTenant(row, instruction.new_tenant_rule, cm)) {
      if (currentTenant) {
        tenants.push(currentTenant);
      }
      currentTenant = {
        suite_id: getCellValue(row, cm.suite_id),
        tenant_name: getCellValue(row, cm.tenant_name),
        lease_start: getCellValue(row, cm.lease_start),
        lease_end: getCellValue(row, cm.lease_end),
        gla_sqft: getNumericValue(row, cm.gla_sqft),
        monthly_base_rent: getNumericValue(row, cm.monthly_base_rent),
        base_rent_psf: getNumericValue(row, cm.base_rent_psf),
        recurring_charges: [],
        future_rent_increases: [],
        notes: '',
      };

      // Check for recurring charge on same row
      const rcCode = getCellValue(row, cm.recurring_charge_code);
      const rcAmt = getNumericValue(row, cm.recurring_charge_amount);
      if (rcCode || rcAmt !== null) {
        currentTenant.recurring_charges.push({
          code: rcCode,
          amount: rcAmt,
          psf: getNumericValue(row, cm.recurring_charge_psf),
        });
      }

      // Check for future rent on same row
      const frDate = getCellValue(row, cm.future_rent_date);
      const frAmt = getNumericValue(row, cm.future_rent_amount);
      if (frDate || frAmt !== null) {
        currentTenant.future_rent_increases.push({
          effective_date: frDate,
          monthly_amount: frAmt,
          psf: getNumericValue(row, cm.future_rent_psf),
        });
      }
    } else if (currentTenant) {
      // Continuation row — collect recurring charges and future rents
      const rcCode = getCellValue(row, cm.recurring_charge_code);
      const rcAmt = getNumericValue(row, cm.recurring_charge_amount);
      if (rcCode || rcAmt !== null) {
        // Skip the "*" summary lines
        if (rcCode !== '*') {
          currentTenant.recurring_charges.push({
            code: rcCode,
            amount: rcAmt,
            psf: getNumericValue(row, cm.recurring_charge_psf),
          });
        }
      }

      const frDate = getCellValue(row, cm.future_rent_date);
      const frAmt = getNumericValue(row, cm.future_rent_amount);
      if (frDate || frAmt !== null) {
        currentTenant.future_rent_increases.push({
          effective_date: frDate,
          monthly_amount: frAmt,
          psf: getNumericValue(row, cm.future_rent_psf),
        });
      }
    }
  }

  if (currentTenant) {
    tenants.push(currentTenant);
  }

  log('system', `Parser found ${tenants.length} tenant blocks`);

  return tenants;
}
