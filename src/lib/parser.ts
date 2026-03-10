import type { ParsingInstruction, TenantObject } from './types';

function colLetterToIndex(letter: string): number {
  if (!letter) return -1;
  // Strip any digits (AI sometimes returns "B6" instead of "B")
  const upper = letter.toUpperCase().trim().replace(/[^A-Z]/g, '');
  if (!upper) return -1;
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
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

/**
 * Dead-simple parser. Logic:
 * 1. Start at data_starts_at_row
 * 2. If suite_id column has a value → new tenant
 * 3. Otherwise → continuation row (attach charges/rents to current tenant)
 * 4. Skip rows matching skip patterns, handle addon space
 */
export function parseSheet(
  data: (string | number | null)[][],
  instruction: ParsingInstruction,
  addLog?: (type: 'system' | 'flag', msg: string) => void
): TenantObject[] {
  const { column_map: cm, data_starts_at_row, skip_row_patterns, addon_space_patterns, custom_columns } = instruction;
  const startRow = (data_starts_at_row ?? 1) - 1;

  const log = addLog || (() => {});

  console.log('[PARSER] Instruction received:', JSON.stringify(instruction, null, 2));
  console.log('[PARSER] Data has', data.length, 'rows. Starting at row', startRow, '(0-indexed)');
  
  for (let d = startRow; d < Math.min(startRow + 5, data.length); d++) {
    const row = data[d];
    if (row) {
      console.log(`[PARSER] Row ${d + 1}:`, 
        `suite_id(${cm.suite_id})="${getCellValue(row, cm.suite_id)}"`,
        `tenant(${cm.tenant_name})="${getCellValue(row, cm.tenant_name)}"`,
        `sqft(${cm.gla_sqft})="${getCellValue(row, cm.gla_sqft)}"`
      );
    }
  }

  log('system', `Parser: start_row=${data_starts_at_row}, suite=${cm.suite_id}, tenant=${cm.tenant_name}, ${data.length} total rows`);

  const tenants: TenantObject[] = [];
  let current: TenantObject | null = null;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === null || c === undefined || String(c).trim() === '')) continue;

    const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();

    if (skip_row_patterns.length > 0 && skip_row_patterns.some(p => {
      try { return new RegExp(p, 'i').test(rowStr); } catch { return rowStr.includes(p.toLowerCase()); }
    })) continue;

    const suiteVal = getCellValue(row, cm.suite_id);
    const tenantVal = getCellValue(row, cm.tenant_name);

    // Addon space
    if (current && (rowStr.includes("add'l space") || rowStr.includes('addl space') || rowStr.includes('additional space') ||
      (addon_space_patterns.length > 0 && addon_space_patterns.some(p => {
        try { return new RegExp(p, 'i').test(rowStr); } catch { return rowStr.includes(p.toLowerCase()); }
      })))) {
      const addonSqft = getNumericValue(row, cm.gla_sqft);
      if (addonSqft !== null && current.gla_sqft !== null) {
        current.gla_sqft += addonSqft;
      }
      current.notes += (current.notes ? '; ' : '') + `Add'l space: ${addonSqft ?? 'N/A'} SF`;
      continue;
    }

    // NEW TENANT
    if (suiteVal) {
      if (tenantVal.toLowerCase().startsWith('psf') || tenantVal.startsWith('(')) {
        // sub-line — treat as continuation
      } else {
        if (current) tenants.push(current);

        current = {
          suite_id: suiteVal,
          tenant_name: tenantVal,
          lease_start: getCellValue(row, cm.lease_start),
          lease_end: getCellValue(row, cm.lease_end),
          gla_sqft: getNumericValue(row, cm.gla_sqft),
          monthly_base_rent: getNumericValue(row, cm.monthly_base_rent),
          base_rent_psf: getNumericValue(row, cm.base_rent_psf),
          recurring_charges: [],
          future_rent_increases: [],
          notes: '',
          custom_fields: {},
        };

        // Extract custom columns on the primary row
        if (custom_columns) {
          for (const [fieldName, colLetter] of Object.entries(custom_columns)) {
            if (colLetter) {
              const val = getCellValue(row, colLetter);
              if (val) current.custom_fields![fieldName] = val;
            }
          }
        }

        collectCharges(row, cm, current);
        continue;
      }
    }

    // CONTINUATION ROW
    if (current) {
      collectCharges(row, cm, current);
      // Also collect custom column values from continuation rows (append if not already set)
      if (custom_columns) {
        for (const [fieldName, colLetter] of Object.entries(custom_columns)) {
          if (colLetter && !current.custom_fields?.[fieldName]) {
            const val = getCellValue(row, colLetter);
            if (val) current.custom_fields![fieldName] = val;
          }
        }
      }
    }
  }

  if (current) tenants.push(current);

  log('system', `${tenants.length} tenant blocks found.`);
  console.log('[PARSER] Found', tenants.length, 'tenants');

  return tenants;
}

function collectCharges(
  row: (string | number | null)[],
  cm: ParsingInstruction['column_map'],
  tenant: TenantObject
) {
  const rcCode = getCellValue(row, cm.recurring_charge_code);
  const rcAmt = getNumericValue(row, cm.recurring_charge_amount);
  if ((rcCode && rcCode !== '*') || rcAmt !== null) {
    if (rcCode !== '*') {
      tenant.recurring_charges.push({
        code: rcCode,
        amount: rcAmt,
        psf: getNumericValue(row, cm.recurring_charge_psf),
      });
    }
  }

  const frDate = getCellValue(row, cm.future_rent_date);
  const frAmt = getNumericValue(row, cm.future_rent_amount);
  if (frDate || frAmt !== null) {
    tenant.future_rent_increases.push({
      effective_date: frDate,
      monthly_amount: frAmt,
      psf: getNumericValue(row, cm.future_rent_psf),
    });
  }
}
