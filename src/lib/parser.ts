import type { ParsingInstruction, TenantObject, RecurringCharge, FutureRentIncrease } from './types';

function colLetterToIndex(letter: string): number {
  if (!letter) return -1;
  const upper = letter.toUpperCase().trim();
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1; // 0-indexed
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
  const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();
  return patterns.some(p => {
    try {
      return new RegExp(p, 'i').test(rowStr);
    } catch {
      return rowStr.includes(p.toLowerCase());
    }
  });
}

function isNewTenant(row: (string | number | null)[], rule: string, colMap: ParsingInstruction['column_map']): boolean {
  const ruleLower = rule.toLowerCase();

  // Common rule: new tenant when tenant_name column is non-empty
  if (ruleLower.includes('tenant') && (ruleLower.includes('non-empty') || ruleLower.includes('not empty') || ruleLower.includes('populated') || ruleLower.includes('has a value'))) {
    const name = getCellValue(row, colMap.tenant_name);
    return name.length > 0;
  }

  // Common rule: new tenant when suite_id changes
  if (ruleLower.includes('suite') && (ruleLower.includes('non-empty') || ruleLower.includes('changes') || ruleLower.includes('new'))) {
    const suite = getCellValue(row, colMap.suite_id);
    return suite.length > 0;
  }

  // Fallback: tenant name is non-empty
  const name = getCellValue(row, colMap.tenant_name);
  return name.length > 0;
}

function isAddonSpace(row: (string | number | null)[], patterns: string[]): boolean {
  const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();
  return patterns.some(p => {
    try {
      return new RegExp(p, 'i').test(rowStr);
    } catch {
      return rowStr.includes(p.toLowerCase());
    }
  });
}

export function parseSheet(
  data: (string | number | null)[][],
  instruction: ParsingInstruction
): TenantObject[] {
  const { column_map: cm, data_starts_at_row, skip_row_patterns, addon_space_patterns, new_tenant_rule } = instruction;
  const startRow = (data_starts_at_row ?? 1) - 1; // convert to 0-indexed

  const tenants: TenantObject[] = [];
  let currentTenant: TenantObject | null = null;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === null || c === undefined || String(c).trim() === '')) continue;

    // Skip summary/total rows
    if (shouldSkipRow(row, skip_row_patterns)) continue;

    // Check if addon space
    if (currentTenant && isAddonSpace(row, addon_space_patterns)) {
      // Attach addon data to current tenant
      const addonSqft = getNumericValue(row, cm.gla_sqft);
      if (addonSqft !== null && currentTenant.gla_sqft !== null) {
        currentTenant.gla_sqft += addonSqft;
      }
      continue;
    }

    // Check for new tenant
    if (isNewTenant(row, new_tenant_rule, cm)) {
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
      // Continuation row — check for recurring charges and future rent increases
      const rcCode = getCellValue(row, cm.recurring_charge_code);
      const rcAmt = getNumericValue(row, cm.recurring_charge_amount);
      if (rcCode || rcAmt !== null) {
        currentTenant.recurring_charges.push({
          code: rcCode,
          amount: rcAmt,
          psf: getNumericValue(row, cm.recurring_charge_psf),
        });
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

  // Don't forget the last tenant
  if (currentTenant) {
    tenants.push(currentTenant);
  }

  return tenants;
}
