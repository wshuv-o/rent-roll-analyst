import * as XLSX from 'xlsx';
import type { TenantObject, ParsingInstruction, GroupSpan } from './types';
import { COLUMN_GROUPS } from './types';

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

function getVal(rec: Record<string, string | number | null> | undefined, label: string): string {
  if (!rec || !label) return '';
  const v = rec[label];
  return v !== null && v !== undefined ? String(v).trim() : '';
}

/**
 * Templatized rent roll export with structured columns,
 * =SUM() / =*12 formulas, charge code pivots, and future rent step columns.
 */
export function exportTemplatizedExcel(
  tenants: TenantObject[],
  fileName: string,
  instruction: ParsingInstruction,
  columnLabels: Record<number, string>
): void {
  const cm = instruction.column_map;

  // Resolve semantic field → label used in tenant data
  const resolveLabel = (field: keyof typeof cm): string => {
    const letter = cm[field];
    if (!letter) return '';
    const idx = colLetterToIndex(letter);
    return idx >= 0 && columnLabels[idx] ? columnLabels[idx] : '';
  };

  const lbl = {
    leaseStart: resolveLabel('lease_start'),
    leaseEnd: resolveLabel('lease_end'),
    gla: resolveLabel('gla_sqft'),
    monthlyRent: resolveLabel('monthly_base_rent'),
    rentPsf: resolveLabel('base_rent_psf'),
    chargeCode: resolveLabel('recurring_charge_code'),
    chargeAmount: resolveLabel('recurring_charge_amount'),
    futureDate: resolveLabel('future_rent_date'),
    futureAmount: resolveLabel('future_rent_amount'),
    futurePsf: resolveLabel('future_rent_psf'),
  };

  // ── 1. Collect unique charge codes from recurring charges ──
  const chargeCodeSet = new Set<string>();
  for (const t of tenants) {
    for (const e of (t.collections['charges'] || [])) {
      const code = getVal(e, lbl.chargeCode);
      if (code) chargeCodeSet.add(code);
    }
  }
  const chargeCodes = [...chargeCodeSet].sort();

  // Per-tenant charge map: code → total amount
  const tenantCharges: Record<string, number>[] = tenants.map(t => {
    const map: Record<string, number> = {};
    for (const e of (t.collections['charges'] || [])) {
      const code = getVal(e, lbl.chargeCode);
      const amt = parseFloat(getVal(e, lbl.chargeAmount)) || 0;
      if (code) map[code] = (map[code] || 0) + amt;
    }
    return map;
  });

  // ── 2. Future rent: identify code label & group by code ──
  // Find a label in future-rent entries that isn't date/amount/psf → treat as code
  let futureCodeLabel = '';
  const allFutureEntries = tenants.flatMap(t => t.collections['future-rent'] || []);
  if (allFutureEntries.length > 0) {
    const knownLabels = new Set([lbl.futureDate, lbl.futureAmount, lbl.futurePsf].filter(Boolean));
    const sampleKeys = Object.keys(allFutureEntries[0]);
    for (const key of sampleKeys) {
      if (knownLabels.has(key)) continue;
      // Check if this column has string values (likely codes)
      const hasValues = allFutureEntries.some(e => getVal(e, key));
      if (hasValues) { futureCodeLabel = key; break; }
    }
  }

  type FutureStep = { date: string; rate: string };
  const futureMaxPerCode: Record<string, number> = {};
  const tenantFutureRent: Record<string, FutureStep[]>[] = tenants.map(t => {
    const entries = t.collections['future-rent'] || [];
    const byCode: Record<string, FutureStep[]> = {};
    for (const e of entries) {
      const code = futureCodeLabel ? getVal(e, futureCodeLabel) : 'General';
      if (!code) continue;
      const date = getVal(e, lbl.futureDate);
      const rate = getVal(e, lbl.futurePsf) || getVal(e, lbl.futureAmount);
      if (!byCode[code]) byCode[code] = [];
      byCode[code].push({ date, rate });
    }
    for (const [code, steps] of Object.entries(byCode)) {
      futureMaxPerCode[code] = Math.max(futureMaxPerCode[code] || 0, steps.length);
    }
    return byCode;
  });

  const futureCodes = Object.keys(futureMaxPerCode).sort();
  // Each code gets (max * 2 + 4) columns = (max + 2) pairs
  const futureCodeCols = futureCodes.map(code => {
    const max = futureMaxPerCode[code];
    const numPairs = max + 2;
    return { code, numCols: numPairs * 2, numPairs };
  });

  // ── 3. Build header rows ──
  const BASE_HEADERS = [
    'Suite', 'Tenant', 'Lease Start', 'Lease End', 'Space (sqft)',
    'Monthly Base Rent', 'Annual Base Rent', 'Rent PSF',
    'Total Other Charges', 'Total Annual Other Charges',
  ];

  // Row 1: base headers + charge codes + (blank + future code repeated)
  const row1: (string | null)[] = [...BASE_HEADERS, ...chargeCodes];
  for (let fi = 0; fi < futureCodeCols.length; fi++) {
    row1.push(''); // blank separator
    const fc = futureCodeCols[fi];
    for (let j = 0; j < fc.numCols; j++) row1.push(fc.code);
  }

  // Row 2: sub-headers (empty for base+charges, Step Date/Rate for future)
  const row2: (string | null)[] = BASE_HEADERS.map(() => '');
  chargeCodes.forEach(() => row2.push(''));
  for (let fi = 0; fi < futureCodeCols.length; fi++) {
    row2.push(''); // blank separator
    const fc = futureCodeCols[fi];
    for (let p = 1; p <= fc.numPairs; p++) {
      row2.push(`Step Date ${p}`);
      row2.push(`Step Rate ${p}`);
    }
  }

  // ── 4. Data rows ──
  const dataRows: (string | number | null)[][] = [];
  const chargeStartCol = BASE_HEADERS.length; // 0-indexed

  tenants.forEach((t, idx) => {
    const monthlyStr = getVal(t.scalars['base-rent'], lbl.monthlyRent);
    const monthly = parseFloat(monthlyStr) || 0;
    const charges = tenantCharges[idx];

    const row: (string | number | null)[] = [
      t.suite_id,
      t.tenant_name,
      getVal(t.scalars['lease'], lbl.leaseStart),
      getVal(t.scalars['lease'], lbl.leaseEnd),
      getVal(t.scalars['space'], lbl.gla) || '',
      monthly || '',
      null, // Annual Base Rent — formula
      getVal(t.scalars['base-rent'], lbl.rentPsf) || '',
      null, // Total Other Charges — formula
      null, // Total Annual Other Charges — formula
    ];

    // Charge code amounts
    for (const code of chargeCodes) {
      row.push(charges[code] || 0);
    }

    // Future rent steps
    const futureData = tenantFutureRent[idx];
    for (let fi = 0; fi < futureCodeCols.length; fi++) {
      row.push(''); // blank separator
      const fc = futureCodeCols[fi];
      const steps = futureData[fc.code] || [];
      for (let p = 0; p < fc.numPairs; p++) {
        row.push(steps[p]?.date || '');
        row.push(steps[p]?.rate || '');
      }
    }

    dataRows.push(row);
  });

  // ── 5. Build worksheet ──
  const allRows = [row1, row2, ...dataRows];
  const ws = XLSX.utils.aoa_to_sheet(allRows);

  // Add formulas (Excel rows are 1-indexed; row 1&2 = headers, data starts row 3)
  tenants.forEach((_, idx) => {
    const r = idx + 3; // Excel row

    // Annual Base Rent = Monthly * 12  (col F = index 5, col G = index 6)
    const monthlyCol = XLSX.utils.encode_col(5);
    const annualCol = XLSX.utils.encode_col(6);
    const cellAnnual = `${annualCol}${r}`;
    ws[cellAnnual] = { t: 'n', f: `${monthlyCol}${r}*12` };

    // Total Other Charges = SUM(charge columns)
    if (chargeCodes.length > 0) {
      const firstCC = XLSX.utils.encode_col(chargeStartCol);
      const lastCC = XLSX.utils.encode_col(chargeStartCol + chargeCodes.length - 1);
      const cellTotal = `${XLSX.utils.encode_col(8)}${r}`;
      ws[cellTotal] = { t: 'n', f: `SUM(${firstCC}${r}:${lastCC}${r})` };
    }

    // Total Annual Other Charges = Total Other Charges * 12
    const totalCol = XLSX.utils.encode_col(8);
    const annTotalCol = XLSX.utils.encode_col(9);
    const cellAnnTotal = `${annTotalCol}${r}`;
    ws[cellAnnTotal] = { t: 'n', f: `${totalCol}${r}*12` };
  });

  // Column widths
  const totalCols = Math.max(...allRows.map(r => r.length), 1);
  ws['!cols'] = Array.from({ length: totalCols }, () => ({ wch: 15 }));
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: allRows.length - 1, c: totalCols - 1 } });

  // Merge header row 1 cells for future rent code names
  const merges: XLSX.Range[] = [];
  let col = BASE_HEADERS.length + chargeCodes.length;
  for (const fc of futureCodeCols) {
    col++; // blank separator
    if (fc.numCols > 1) {
      merges.push({ s: { r: 0, c: col }, e: { r: 0, c: col + fc.numCols - 1 } });
    }
    col += fc.numCols;
  }
  if (merges.length > 0) ws['!merges'] = merges;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Templatized Rent Roll');

  const outputName = fileName.replace(/\.(xlsx|xls)$/i, '') + '_template.xlsx';
  XLSX.writeFile(wb, outputName);
}
