import * as XLSX from 'xlsx';
import type { TenantObject } from './types';

/**
 * Templatized rent roll export.
 * 
 * Layout:
 * Fixed: Suite | Tenant | Lease Start | Lease End | Sqft | Monthly Base Rent | Annual Base Rent (formula) | Rent PSF | Total Other Charges (SUM formula) | Total Annual Other Charges (formula)
 * Current Charges: one column per unique charge code
 * Future Rent: per unique charge code, max_occurrences * 2 columns (Step Date N, Step Rate N), separated by blank columns
 */

// Color palette (ARGB hex strings for xlsx)
const COLORS = {
  headerBg: 'FF1B2A4A',       // dark navy
  headerFont: 'FFFFFFFF',     // white
  currentChargesBg: 'FF2D5F2D', // dark green
  futureRentBg: 'FF5C3D1A',   // dark brown/orange
  chargeCategoryBg: 'FF3A3A6A', // muted purple
  dataBg: 'FF0D1117',         // very dark
  altRowBg: 'FF161B22',       // slightly lighter
  borderColor: 'FF30363D',
  formulaFont: 'FF58A6FF',    // blue for formula cells
};

function colToRef(col: number, row: number): string {
  let letter = '';
  let n = col;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return `${letter}${row}`;
}

function makeHeaderStyle(): Record<string, unknown> {
  return {
    font: { bold: true, color: { rgb: COLORS.headerFont }, sz: 10 },
    fill: { fgColor: { rgb: COLORS.headerBg }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top: { style: 'thin', color: { rgb: COLORS.borderColor } },
      bottom: { style: 'thin', color: { rgb: COLORS.borderColor } },
      left: { style: 'thin', color: { rgb: COLORS.borderColor } },
      right: { style: 'thin', color: { rgb: COLORS.borderColor } },
    },
  };
}

function makeChargeHeaderStyle(): Record<string, unknown> {
  return {
    ...makeHeaderStyle(),
    fill: { fgColor: { rgb: COLORS.currentChargesBg }, patternType: 'solid' },
  };
}

function makeFutureHeaderStyle(): Record<string, unknown> {
  return {
    ...makeHeaderStyle(),
    fill: { fgColor: { rgb: COLORS.futureRentBg }, patternType: 'solid' },
  };
}

function makeCategoryStyle(bg: string): Record<string, unknown> {
  return {
    font: { bold: true, color: { rgb: COLORS.headerFont }, sz: 10 },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: COLORS.borderColor } },
      bottom: { style: 'thin', color: { rgb: COLORS.borderColor } },
      left: { style: 'thin', color: { rgb: COLORS.borderColor } },
      right: { style: 'thin', color: { rgb: COLORS.borderColor } },
    },
  };
}

/** Extract a numeric value from a label→value record for a given label substring match, with positional fallback */
function findNumericValue(record: Record<string, string | number | null> | undefined, ...keywords: string[]): number | null {
  if (!record) return null;
  // Try keyword match first
  for (const [label, val] of Object.entries(record)) {
    const lower = label.toLowerCase();
    if (keywords.some(k => lower.includes(k.toLowerCase()))) {
      if (val === null || val === '') return null;
      const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/[,$]/g, ''));
      return isNaN(n) ? null : n;
    }
  }
  // Fallback: return the first parseable numeric value
  for (const [, val] of Object.entries(record)) {
    if (val === null || val === '') continue;
    const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/[,$]/g, ''));
    if (!isNaN(n)) return n;
  }
  return null;
}

/** Get the Nth numeric value from a record (0-indexed) */
function findNthNumericValue(record: Record<string, string | number | null> | undefined, index: number): number | null {
  if (!record) return null;
  let count = 0;
  for (const [, val] of Object.entries(record)) {
    if (val === null || val === '') continue;
    const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/[,$]/g, ''));
    if (!isNaN(n)) {
      if (count === index) return n;
      count++;
    }
  }
  return null;
}

function findStringValue(record: Record<string, string | number | null> | undefined, ...keywords: string[]): string {
  if (!record) return '';
  // Try keyword match first
  for (const [label, val] of Object.entries(record)) {
    const lower = label.toLowerCase();
    if (keywords.some(k => lower.includes(k.toLowerCase()))) {
      return val !== null && val !== undefined ? String(val).trim() : '';
    }
  }
  // Fallback: return the first non-empty string value
  for (const [, val] of Object.entries(record)) {
    if (val !== null && val !== undefined && String(val).trim()) return String(val).trim();
  }
  return '';
}

/** Get the Nth non-empty string value from a record (0-indexed) */
function findNthStringValue(record: Record<string, string | number | null> | undefined, index: number): string {
  if (!record) return '';
  let count = 0;
  for (const [, val] of Object.entries(record)) {
    const s = val !== null && val !== undefined ? String(val).trim() : '';
    if (s) {
      if (count === index) return s;
      count++;
    }
  }
  return '';
}

export function exportTemplatizedRentRoll(tenants: TenantObject[], fileName: string): void {
  // ─── 1. Gather unique charge codes from recurring charges ───
  const chargeCodeSet = new Set<string>();
  for (const t of tenants) {
    const entries = t.collections['charges'];
    if (!entries) continue;
    for (const entry of entries) {
      const code = findChargeCode(entry);
      if (code) chargeCodeSet.add(code);
    }
  }
  const chargeCodes = Array.from(chargeCodeSet).sort();

  // ─── 2. Gather future rent steps per charge code and find max counts ───
  const futureRentCodeMap = new Map<string, number>(); // code → max step count across tenants
  for (const t of tenants) {
    const entries = t.collections['future-rent'];
    if (!entries) continue;
    // Group by charge code within this tenant's future rent entries
    const codeCount = new Map<string, number>();
    for (const entry of entries) {
      const code = findChargeCode(entry) || 'RNT'; // default to RNT if no code
      codeCount.set(code, (codeCount.get(code) || 0) + 1);
    }
    for (const [code, count] of codeCount) {
      futureRentCodeMap.set(code, Math.max(futureRentCodeMap.get(code) || 0, count));
    }
  }
  const futureRentCodes = Array.from(futureRentCodeMap.keys()).sort();

  // ─── 3. Build column layout ───
  // Fixed columns (indices 0-9)
  const fixedHeaders = [
    'Suite', 'Tenant', 'Lease Start', 'Lease End', 'Sqft',
    'Monthly Base Rent', 'Annual Base Rent', 'Rent PSF',
    'Total Other Charges', 'Total Annual Other Charges',
  ];
  const FIXED_COUNT = fixedHeaders.length;

  // Current charges columns
  const chargeStartCol = FIXED_COUNT;
  const chargeColCount = chargeCodes.length;

  // Future rent columns
  let futureStartCol = chargeStartCol + chargeColCount;
  if (chargeColCount > 0) futureStartCol += 1; // blank separator

  // Build future rent code blocks: for each code, maxSteps * 2 columns + 1 blank separator
  const futureCodeBlocks: { code: string; startCol: number; steps: number }[] = [];
  let col = futureStartCol;
  for (const code of futureRentCodes) {
    const maxSteps = futureRentCodeMap.get(code) || 1;
    futureCodeBlocks.push({ code, startCol: col, steps: maxSteps });
    col += maxSteps * 2 + 1; // +1 for blank separator
  }

  const totalCols = col > futureStartCol ? col - 1 : futureStartCol + (futureRentCodes.length > 0 ? 0 : -1);

  // ─── 4. Build worksheet data ───
  // Row 0: Category banner (merged cells)
  // Row 1: Charge code banner for future rent
  // Row 2: Sub-headers (Step Date 1, Step Rate 1, ...)
  // Row 3+: Data rows
  const HEADER_ROWS = 3;
  const ws: XLSX.WorkSheet = {};
  const merges: XLSX.Range[] = [];

  // Set column widths
  const colWidths: { wch: number }[] = [];
  for (let c = 0; c < Math.max(totalCols + 1, FIXED_COUNT + chargeColCount + 1); c++) {
    if (c <= 1) colWidths.push({ wch: 20 }); // Suite, Tenant
    else if (c <= 4) colWidths.push({ wch: 14 }); // dates, sqft
    else if (c <= 9) colWidths.push({ wch: 18 }); // rent columns
    else colWidths.push({ wch: 14 });
  }
  ws['!cols'] = colWidths;

  // ── Row 0: Category banner ──
  // Fixed section header
  for (let c = 0; c < FIXED_COUNT; c++) {
    ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: makeHeaderStyle() };
  }

  // "Current Charges" banner
  if (chargeColCount > 0) {
    ws[XLSX.utils.encode_cell({ r: 0, c: chargeStartCol })] = {
      v: 'Current Charges', s: makeCategoryStyle(COLORS.currentChargesBg),
    };
    if (chargeColCount > 1) {
      merges.push({ s: { r: 0, c: chargeStartCol }, e: { r: 0, c: chargeStartCol + chargeColCount - 1 } });
    }
    for (let c = chargeStartCol + 1; c < chargeStartCol + chargeColCount; c++) {
      ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: makeCategoryStyle(COLORS.currentChargesBg) };
    }
  }

  // Future rent code banners on row 0 + row 1
  for (const block of futureCodeBlocks) {
    const blockWidth = block.steps * 2;
    // Row 0: "Dynamic" or just the code category (CAM/TAX/etc)
    const categoryLabel = getCategoryForCode(block.code);
    ws[XLSX.utils.encode_cell({ r: 0, c: block.startCol })] = {
      v: categoryLabel, s: makeCategoryStyle(COLORS.futureRentBg),
    };
    if (blockWidth > 1) {
      merges.push({ s: { r: 0, c: block.startCol }, e: { r: 0, c: block.startCol + blockWidth - 1 } });
    }
    for (let c = block.startCol + 1; c < block.startCol + blockWidth; c++) {
      ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: makeCategoryStyle(COLORS.futureRentBg) };
    }

    // Row 1: Charge code repeated
    ws[XLSX.utils.encode_cell({ r: 1, c: block.startCol })] = {
      v: block.code, s: makeCategoryStyle(COLORS.chargeCategoryBg),
    };
    if (blockWidth > 1) {
      merges.push({ s: { r: 1, c: block.startCol }, e: { r: 1, c: block.startCol + blockWidth - 1 } });
    }
    for (let c = block.startCol + 1; c < block.startCol + blockWidth; c++) {
      ws[XLSX.utils.encode_cell({ r: 1, c })] = { v: '', s: makeCategoryStyle(COLORS.chargeCategoryBg) };
    }
  }

  // ── Row 2: Column headers ──
  // Fixed headers
  for (let c = 0; c < FIXED_COUNT; c++) {
    ws[XLSX.utils.encode_cell({ r: 2, c })] = { v: fixedHeaders[c], s: makeHeaderStyle() };
  }
  // Merge fixed headers across rows 0-1 (they span the category rows)
  for (let c = 0; c < FIXED_COUNT; c++) {
    merges.push({ s: { r: 0, c }, e: { r: 1, c } });
  }

  // Charge code headers
  for (let i = 0; i < chargeCodes.length; i++) {
    const c = chargeStartCol + i;
    ws[XLSX.utils.encode_cell({ r: 2, c })] = { v: chargeCodes[i], s: makeChargeHeaderStyle() };
    // Merge row 1 for charge codes
    ws[XLSX.utils.encode_cell({ r: 1, c })] = { v: '', s: makeChargeHeaderStyle() };
  }

  // Future rent step headers
  for (const block of futureCodeBlocks) {
    for (let s = 0; s < block.steps; s++) {
      const dateCol = block.startCol + s * 2;
      const rateCol = block.startCol + s * 2 + 1;
      ws[XLSX.utils.encode_cell({ r: 2, c: dateCol })] = {
        v: `Step Date ${s + 1}`, s: makeFutureHeaderStyle(),
      };
      ws[XLSX.utils.encode_cell({ r: 2, c: rateCol })] = {
        v: `Step Rate ${s + 1}`, s: makeFutureHeaderStyle(),
      };
    }
  }

  // ── Data rows ──
  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = HEADER_ROWS + ti;

    // Fixed columns
    ws[XLSX.utils.encode_cell({ r, c: 0 })] = { v: t.suite_id };
    ws[XLSX.utils.encode_cell({ r, c: 1 })] = { v: t.tenant_name };

    // Lease dates from scalar group
    const leaseData = t.scalars['lease'];
    ws[XLSX.utils.encode_cell({ r, c: 2 })] = { v: findStringValue(leaseData, 'start', 'begin', 'commence') };
    ws[XLSX.utils.encode_cell({ r, c: 3 })] = { v: findStringValue(leaseData, 'end', 'expire', 'expir', 'terminat') };

    // Space
    const spaceData = t.scalars['space'];
    const sqft = findNumericValue(spaceData, 'gla', 'sqft', 'sf', 'area', 'size', 'nra');
    ws[XLSX.utils.encode_cell({ r, c: 4 })] = { v: sqft ?? '' };

    // Base rent
    const rentData = t.scalars['base-rent'];
    const monthlyRent = findNumericValue(rentData, 'monthly', 'month', 'rent');
    ws[XLSX.utils.encode_cell({ r, c: 5 })] = { v: monthlyRent ?? '' };

    // Annual Base Rent = Monthly * 12 (formula)
    const monthlyRef = colToRef(5, r + 1);
    ws[XLSX.utils.encode_cell({ r, c: 6 })] = { f: `${monthlyRef}*12` };

    // Rent PSF
    const rentPsf = findNumericValue(rentData, 'psf', 'per sq', 'per foot');
    ws[XLSX.utils.encode_cell({ r, c: 7 })] = { v: rentPsf ?? '' };

    // Current charges — fill per code
    const chargeEntries = t.collections['charges'] || [];
    const chargeByCode: Record<string, number> = {};
    for (const entry of chargeEntries) {
      const code = findChargeCode(entry);
      if (!code) continue;
      const amt = findChargeAmount(entry);
      if (amt !== null) {
        chargeByCode[code] = (chargeByCode[code] || 0) + amt;
      }
    }

    for (let i = 0; i < chargeCodes.length; i++) {
      const c = chargeStartCol + i;
      ws[XLSX.utils.encode_cell({ r, c })] = { v: chargeByCode[chargeCodes[i]] ?? 0 };
    }

    // Total Other Charges = SUM of charge columns (formula)
    if (chargeColCount > 0) {
      const firstChargeRef = colToRef(chargeStartCol, r + 1);
      const lastChargeRef = colToRef(chargeStartCol + chargeColCount - 1, r + 1);
      ws[XLSX.utils.encode_cell({ r, c: 8 })] = { f: `SUM(${firstChargeRef}:${lastChargeRef})` };
    } else {
      ws[XLSX.utils.encode_cell({ r, c: 8 })] = { v: 0 };
    }

    // Total Annual Other Charges = Total Other Charges * 12 (formula)
    const totalChargesRef = colToRef(8, r + 1);
    ws[XLSX.utils.encode_cell({ r, c: 9 })] = { f: `${totalChargesRef}*12` };

    // Future rent steps
    const futureEntries = t.collections['future-rent'] || [];
    // Group future entries by charge code
    const futureByCode: Record<string, { date: string; rate: number | null }[]> = {};
    for (const entry of futureEntries) {
      const code = findChargeCode(entry) || 'RNT';
      if (!futureByCode[code]) futureByCode[code] = [];
      const date = findStringValue(entry, 'date', 'effective', 'start');
      const rate = findNumericValue(entry, 'psf', 'rate', 'amount', 'rent');
      futureByCode[code].push({ date, rate });
    }

    for (const block of futureCodeBlocks) {
      const steps = futureByCode[block.code] || [];
      for (let s = 0; s < block.steps; s++) {
        const dateCol = block.startCol + s * 2;
        const rateCol = block.startCol + s * 2 + 1;
        if (s < steps.length) {
          ws[XLSX.utils.encode_cell({ r, c: dateCol })] = { v: steps[s].date };
          ws[XLSX.utils.encode_cell({ r, c: rateCol })] = { v: steps[s].rate ?? '' };
        }
      }
    }
  }

  // Set merges
  ws['!merges'] = merges;

  // Set range
  const maxRow = HEADER_ROWS + tenants.length - 1;
  const maxCol = Math.max(
    FIXED_COUNT - 1,
    chargeStartCol + chargeColCount - 1,
    ...futureCodeBlocks.map(b => b.startCol + b.steps * 2 - 1),
    0
  );
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(maxRow, 2), c: Math.max(maxCol, FIXED_COUNT - 1) } });

  // Create workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Templatized Rent Roll');

  const outputName = fileName.replace(/\.(xlsx|xls)$/i, '') + '_template.xlsx';
  XLSX.writeFile(wb, outputName);
}

// ─── Helpers to extract values from label→value records ───

function findChargeCode(entry: Record<string, string | number | null>): string | null {
  for (const [label, val] of Object.entries(entry)) {
    const lower = label.toLowerCase();
    if (lower.includes('code') || lower.includes('type') || lower.includes('charge code') || lower.includes('description')) {
      if (val !== null && val !== '') return String(val).trim();
    }
  }
  // Fallback: if there's a short uppercase string, it might be the code
  for (const [, val] of Object.entries(entry)) {
    if (val !== null && typeof val === 'string' && val.length <= 6 && val === val.toUpperCase() && /^[A-Z]+$/.test(val)) {
      return val;
    }
  }
  return null;
}

function findChargeAmount(entry: Record<string, string | number | null>): number | null {
  for (const [label, val] of Object.entries(entry)) {
    const lower = label.toLowerCase();
    if (lower.includes('amount') || lower.includes('charge') || lower.includes('monthly') || lower.includes('amt')) {
      if (val === null || val === '') return null;
      const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/[,$]/g, ''));
      return isNaN(n) ? null : n;
    }
  }
  // Fallback: first numeric value that isn't a PSF
  for (const [label, val] of Object.entries(entry)) {
    if (label.toLowerCase().includes('psf') || label.toLowerCase().includes('per')) continue;
    if (val !== null && val !== '') {
      const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/[,$]/g, ''));
      if (!isNaN(n)) return n;
    }
  }
  return null;
}

function getCategoryForCode(code: string): string {
  const upper = code.toUpperCase();
  if (['TAX', 'PROPTAX', 'RETAX'].includes(upper)) return 'TAX';
  if (['CAM', 'CTOC', 'MKT', 'STO', 'HVAC', 'INS', 'RNT', 'BOFC'].includes(upper)) return upper;
  return code;
}
