import type { AnonymizationMapping } from './types';

// Patterns for detecting sensitive data
const DOLLAR_PATTERN = /^\$?[\d,]+\.?\d*$/;
const NAME_PATTERN = /^[A-Z][a-z]+(?:\s[A-Z][a-z]+)+$/;
const COMPANY_PATTERN = /(?:LLC|Inc|Corp|Ltd|LP|Co\.|Associates|Group|Holdings|Properties|Realty|Management|Services|Partners|Enterprises|Company|Restaurant|Cafe|Salon|Shop|Store|Bank|Pharmacy|Clinic|Medical|Dental|Law|Office|Studio|Fitness|Gym|Hotel|Motel|Inn)\b/i;
const SUITE_PATTERN = /^(?:Suite|Ste|Unit|#|Apt|Space|Rm|Room)?\s*[A-Z0-9][-A-Z0-9]*$/i;
const PURE_NUMBER_PATTERN = /^\d+\.?\d*$/;

function isDate(value: string): boolean {
  const datePatterns = [
    /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/,
    /^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/,
    /^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i,
  ];
  return datePatterns.some(p => p.test(value.trim()));
}

function isHeader(value: string): boolean {
  const headerKeywords = [
    'tenant', 'suite', 'unit', 'lease', 'start', 'end', 'expire', 'rent',
    'sqft', 'sf', 'gla', 'nra', 'area', 'monthly', 'annual', 'psf',
    'charge', 'cam', 'tax', 'insurance', 'total', 'base', 'date',
    'name', 'occupant', 'description', 'code', 'amount', 'rate',
    'term', 'commence', 'expiration', 'escalation', 'increase',
    'floor', 'building', 'property', 'status', 'type', 'notes',
  ];
  const lower = value.toLowerCase().trim();
  return headerKeywords.some(k => lower.includes(k));
}

function looksLikeDollar(value: string): boolean {
  const cleaned = value.replace(/[\s$,]/g, '');
  return DOLLAR_PATTERN.test(cleaned) && parseFloat(cleaned) > 0;
}

function looksLikeName(value: string): boolean {
  const trimmed = value.trim();
  if (trimmed.length < 3 || trimmed.length > 100) return false;
  if (PURE_NUMBER_PATTERN.test(trimmed)) return false;
  if (isDate(trimmed)) return false;
  if (isHeader(trimmed)) return false;
  if (NAME_PATTERN.test(trimmed)) return true;
  if (COMPANY_PATTERN.test(trimmed)) return true;
  // Multi-word string that isn't a number or date — likely a name
  if (trimmed.includes(' ') && /[a-zA-Z]/.test(trimmed) && !PURE_NUMBER_PATTERN.test(trimmed.replace(/\s/g, ''))) return true;
  return false;
}

function looksLikeSuiteId(value: string, colHeader?: string): boolean {
  const trimmed = value.trim();
  if (trimmed.length < 1 || trimmed.length > 20) return false;
  if (isDate(trimmed)) return false;
  const headerHint = colHeader?.toLowerCase() || '';
  if (headerHint.includes('suite') || headerHint.includes('unit') || headerHint.includes('space')) {
    return SUITE_PATTERN.test(trimmed) || /^[A-Z0-9][-A-Z0-9]*$/i.test(trimmed);
  }
  return SUITE_PATTERN.test(trimmed);
}

/**
 * Auto-detect which rows are headers by scanning for header-keyword density.
 * Returns indices of rows that look like headers (metadata rows + actual column headers).
 */
export function detectHeaderRows(data: (string | number | null)[][]): number[] {
  const headerRows: number[] = [];
  const scanLimit = Math.min(20, data.length); // Only scan first 20 rows

  for (let i = 0; i < scanLimit; i++) {
    const row = data[i];
    if (!row) continue;
    const nonEmpty = row.filter(c => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonEmpty.length === 0) continue;

    const headerCells = nonEmpty.filter(c => isHeader(String(c)));
    const ratio = headerCells.length / nonEmpty.length;

    // If more than 40% of non-empty cells look like headers, it's a header row
    if (ratio >= 0.4 && headerCells.length >= 2) {
      headerRows.push(i);
    }
  }

  // Also include metadata rows (first few rows before headers that contain report info)
  if (headerRows.length > 0) {
    const firstHeader = headerRows[0];
    for (let i = 0; i < firstHeader; i++) {
      if (!headerRows.includes(i)) headerRows.push(i);
    }
    headerRows.sort((a, b) => a - b);
  }

  // Fallback: if no headers detected, use first 3 rows
  if (headerRows.length === 0) {
    return [0, 1, 2].filter(i => i < data.length);
  }

  return headerRows;
}


export function anonymizeSheet(
  data: (string | number | null)[][],
  headerRowIndices: number[]
): { anonymized: (string | number | null)[][]; mapping: AnonymizationMapping; stats: { names: number; suites: number; amounts: number } } {
  const mapping: AnonymizationMapping = {
    tenantNames: new Map(),
    suiteIds: new Map(),
    amounts: new Map(),
    reverseMap: new Map(),
  };

  let tenantCounter = 0;
  let suiteCounter = 0;
  let amtCounter = 0;
  let stats = { names: 0, suites: 0, amounts: 0 };

  // Identify headers for column context
  const headers: string[] = [];
  if (headerRowIndices.length > 0) {
    const hRow = data[headerRowIndices[0]];
    if (hRow) {
      for (let c = 0; c < hRow.length; c++) {
        headers[c] = String(hRow[c] || '');
      }
    }
  }

  const anonymized = data.map((row, rowIdx) => {
    if (headerRowIndices.includes(rowIdx)) return [...row]; // keep headers

    return row.map((cell, colIdx) => {
      if (cell === null || cell === undefined || cell === '') return cell;
      const str = String(cell).trim();
      if (str === '') return cell;

      // Keep dates
      if (isDate(str)) return cell;

      // Keep headers
      if (isHeader(str)) return cell;

      // Dollar amounts
      if (looksLikeDollar(str)) {
        if (mapping.amounts.has(str)) return mapping.amounts.get(str)!;
        amtCounter++;
        const placeholder = `AMT_${amtCounter}`;
        mapping.amounts.set(str, placeholder);
        mapping.reverseMap.set(placeholder, str);
        stats.amounts++;
        return placeholder;
      }

      // Suite IDs
      if (looksLikeSuiteId(str, headers[colIdx])) {
        if (mapping.suiteIds.has(str)) return mapping.suiteIds.get(str)!;
        suiteCounter++;
        const letter = String.fromCharCode(64 + ((suiteCounter - 1) % 26) + 1);
        const placeholder = `SUITE_${letter}${suiteCounter > 26 ? Math.ceil(suiteCounter / 26) : ''}`;
        mapping.suiteIds.set(str, placeholder);
        mapping.reverseMap.set(placeholder, str);
        stats.suites++;
        return placeholder;
      }

      // Tenant names
      if (looksLikeName(str)) {
        if (mapping.tenantNames.has(str)) return mapping.tenantNames.get(str)!;
        tenantCounter++;
        const placeholder = `TENANT_${String(tenantCounter).padStart(3, '0')}`;
        mapping.tenantNames.set(str, placeholder);
        mapping.reverseMap.set(placeholder, str);
        stats.names++;
        return placeholder;
      }

      return cell;
    });
  });

  return { anonymized, mapping, stats };
}

export function deanonymize(tenants: import('./types').TenantObject[], mapping: AnonymizationMapping): import('./types').TenantObject[] {
  const reverse = mapping.reverseMap;

  function restore(val: string): string {
    return reverse.get(val) || val;
  }

  return tenants.map(t => ({
    ...t,
    suite_id: restore(t.suite_id),
    tenant_name: restore(t.tenant_name),
    notes: restore(t.notes),
  }));
}

function restoreAmount(val: number | null, reverse: Map<string, string>): number | null {
  if (val === null) return null;
  const key = `AMT_${val}`;
  const restored = reverse.get(key);
  if (restored) {
    const parsed = parseFloat(restored.replace(/[\s$,]/g, ''));
    return isNaN(parsed) ? val : parsed;
  }
  return val;
}
