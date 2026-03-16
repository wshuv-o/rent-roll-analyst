import * as XLSX from 'xlsx';
import type { TenantObject, ParsingInstruction, GroupSpan, CustomGroup } from './types';
import { colLetterToIndex, getCellValue, getRawCellValue } from './col-utils';

/**
 * Templatized rent roll export — reads directly from raw rows using column mapping.
 * No heuristic string parsing. Values come straight from Excel cells.
 */

const COLORS = {
  headerBg: 'FF1B2A4A',
  headerFont: 'FFFFFFFF',
  currentChargesBg: 'FF2D5F2D',
  futureRentBg: 'FF5C3D1A',
  chargeCategoryBg: 'FF3A3A6A',
  borderColor: 'FF30363D',
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
  return { ...makeHeaderStyle(), fill: { fgColor: { rgb: COLORS.currentChargesBg }, patternType: 'solid' } };
}

function makeFutureHeaderStyle(): Record<string, unknown> {
  return { ...makeHeaderStyle(), fill: { fgColor: { rgb: COLORS.futureRentBg }, patternType: 'solid' } };
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

/** Parse a raw cell value to number, stripping currency symbols */
function toNumber(val: string | number | null): number {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  const n = parseFloat(String(val).replace(/[,$%]/g, ''));
  return isNaN(n) ? 0 : n;
}

/** Get string from raw cell */
function toStr(val: string | number | null): string {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

export function exportTemplatizedRentRoll(
  tenants: TenantObject[],
  fileName: string,
  instruction: ParsingInstruction,
  groupSpans: GroupSpan[],
  columnLabels: Record<number, string>,
  customGroups: CustomGroup[] = []
): void {
  const cm = instruction.column_map;

  // Column indices from mapping
  const colIdx = {
    leaseStart: colLetterToIndex(cm.lease_start),
    leaseEnd: colLetterToIndex(cm.lease_end),
    sqft: colLetterToIndex(cm.gla_sqft),
    monthlyRent: colLetterToIndex(cm.monthly_base_rent),
    rentPsf: colLetterToIndex(cm.base_rent_psf),
    chargeCode: colLetterToIndex(cm.recurring_charge_code),
    chargeAmount: colLetterToIndex(cm.recurring_charge_amount),
    chargePsf: colLetterToIndex(cm.recurring_charge_psf),
    futureDate: colLetterToIndex(cm.future_rent_date),
    futureAmount: colLetterToIndex(cm.future_rent_amount),
    futurePsf: colLetterToIndex(cm.future_rent_psf),
  };

  // ─── 1. Gather unique charge codes ───
  const chargeCodeSet = new Set<string>();
  for (const t of tenants) {
    for (const row of t.rawRows) {
      const code = getCellValue(row, colIdx.chargeCode);
      if (code) chargeCodeSet.add(code);
    }
  }
  const chargeCodes = Array.from(chargeCodeSet).sort();

  // ─── 2. Gather future rent steps per charge code ───
  const futureRentCodeMap = new Map<string, number>();
  for (const t of tenants) {
    const codeCount = new Map<string, number>();
    for (const row of t.rawRows) {
      const date = getCellValue(row, colIdx.futureDate);
      if (!date) continue;
      const code = getCellValue(row, colIdx.chargeCode) || 'RNT';
      codeCount.set(code, (codeCount.get(code) || 0) + 1);
    }
    for (const [code, count] of codeCount) {
      futureRentCodeMap.set(code, Math.max(futureRentCodeMap.get(code) || 0, count));
    }
  }
  const futureRentCodes = Array.from(futureRentCodeMap.keys()).sort();

  // ─── 3. Build column layout ───
  const fixedHeaders = [
    'Suite', 'Tenant', 'Lease Start', 'Lease End', 'Sqft',
    'Monthly Base Rent', 'Annual Base Rent', 'Rent PSF',
    'Total Other Charges', 'Total Annual Other Charges',
  ];
  const FIXED_COUNT = fixedHeaders.length;

  const chargeStartCol = FIXED_COUNT;
  const chargeColCount = chargeCodes.length;

  let futureStartCol = chargeStartCol + chargeColCount;
  if (chargeColCount > 0) futureStartCol += 1;

  const futureCodeBlocks: { code: string; startCol: number; steps: number }[] = [];
  let col = futureStartCol;
  for (const code of futureRentCodes) {
    const maxSteps = futureRentCodeMap.get(code) || 1;
    futureCodeBlocks.push({ code, startCol: col, steps: maxSteps });
    col += maxSteps * 2 + 1;
  }

  const totalCols = col > futureStartCol ? col - 1 : futureStartCol + (futureRentCodes.length > 0 ? 0 : -1);

  // ─── 4. Build worksheet ───
  const HEADER_ROWS = 3;
  const ws: XLSX.WorkSheet = {};
  const merges: XLSX.Range[] = [];

  const colWidths: { wch: number }[] = [];
  for (let c = 0; c < Math.max(totalCols + 1, FIXED_COUNT + chargeColCount + 1); c++) {
    if (c <= 1) colWidths.push({ wch: 20 });
    else if (c <= 4) colWidths.push({ wch: 14 });
    else if (c <= 9) colWidths.push({ wch: 18 });
    else colWidths.push({ wch: 14 });
  }
  ws['!cols'] = colWidths;

  // Row 0: Category banners
  for (let c = 0; c < FIXED_COUNT; c++) {
    ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: makeHeaderStyle() };
  }

  if (chargeColCount > 0) {
    ws[XLSX.utils.encode_cell({ r: 0, c: chargeStartCol })] = { v: 'Current Charges', s: makeCategoryStyle(COLORS.currentChargesBg) };
    if (chargeColCount > 1) {
      merges.push({ s: { r: 0, c: chargeStartCol }, e: { r: 0, c: chargeStartCol + chargeColCount - 1 } });
    }
    for (let c = chargeStartCol + 1; c < chargeStartCol + chargeColCount; c++) {
      ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: makeCategoryStyle(COLORS.currentChargesBg) };
    }
  }

  for (const block of futureCodeBlocks) {
    const blockWidth = block.steps * 2;
    ws[XLSX.utils.encode_cell({ r: 0, c: block.startCol })] = { v: block.code, s: makeCategoryStyle(COLORS.futureRentBg) };
    if (blockWidth > 1) {
      merges.push({ s: { r: 0, c: block.startCol }, e: { r: 0, c: block.startCol + blockWidth - 1 } });
    }
    for (let c = block.startCol + 1; c < block.startCol + blockWidth; c++) {
      ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: makeCategoryStyle(COLORS.futureRentBg) };
    }
    ws[XLSX.utils.encode_cell({ r: 1, c: block.startCol })] = { v: block.code, s: makeCategoryStyle(COLORS.chargeCategoryBg) };
    if (blockWidth > 1) {
      merges.push({ s: { r: 1, c: block.startCol }, e: { r: 1, c: block.startCol + blockWidth - 1 } });
    }
    for (let c = block.startCol + 1; c < block.startCol + blockWidth; c++) {
      ws[XLSX.utils.encode_cell({ r: 1, c })] = { v: '', s: makeCategoryStyle(COLORS.chargeCategoryBg) };
    }
  }

  // Row 2: Column headers
  for (let c = 0; c < FIXED_COUNT; c++) {
    ws[XLSX.utils.encode_cell({ r: 2, c })] = { v: fixedHeaders[c], s: makeHeaderStyle() };
  }
  for (let c = 0; c < FIXED_COUNT; c++) {
    merges.push({ s: { r: 0, c }, e: { r: 1, c } });
  }
  for (let i = 0; i < chargeCodes.length; i++) {
    const c = chargeStartCol + i;
    ws[XLSX.utils.encode_cell({ r: 2, c })] = { v: chargeCodes[i], s: makeChargeHeaderStyle() };
    ws[XLSX.utils.encode_cell({ r: 1, c })] = { v: '', s: makeChargeHeaderStyle() };
  }
  for (const block of futureCodeBlocks) {
    for (let s = 0; s < block.steps; s++) {
      const dateCol = block.startCol + s * 2;
      const rateCol = block.startCol + s * 2 + 1;
      ws[XLSX.utils.encode_cell({ r: 2, c: dateCol })] = { v: `Step Date ${s + 1}`, s: makeFutureHeaderStyle() };
      ws[XLSX.utils.encode_cell({ r: 2, c: rateCol })] = { v: `Step Rate ${s + 1}`, s: makeFutureHeaderStyle() };
    }
  }

  // ── Data rows ──
  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = HEADER_ROWS + ti;
    const primaryRow = t.rawRows[0] || [];

    // Fixed columns — read directly from raw cells
    ws[XLSX.utils.encode_cell({ r, c: 0 })] = { v: t.suite_id };
    ws[XLSX.utils.encode_cell({ r, c: 1 })] = { v: t.tenant_name };
    ws[XLSX.utils.encode_cell({ r, c: 2 })] = { v: toStr(getRawCellValue(primaryRow, colIdx.leaseStart)) };
    ws[XLSX.utils.encode_cell({ r, c: 3 })] = { v: toStr(getRawCellValue(primaryRow, colIdx.leaseEnd)) };
    ws[XLSX.utils.encode_cell({ r, c: 4 })] = { v: toNumber(getRawCellValue(primaryRow, colIdx.sqft)), t: 'n' };

    const monthlyRent = toNumber(getRawCellValue(primaryRow, colIdx.monthlyRent));
    ws[XLSX.utils.encode_cell({ r, c: 5 })] = { v: monthlyRent, t: 'n' };

    // Annual Base Rent = Monthly * 12 (formula)
    ws[XLSX.utils.encode_cell({ r, c: 6 })] = { f: `${colToRef(5, r + 1)}*12`, t: 'n' };

    ws[XLSX.utils.encode_cell({ r, c: 7 })] = { v: toNumber(getRawCellValue(primaryRow, colIdx.rentPsf)), t: 'n' };

    // Current charges — read code and amount directly from raw rows
    const chargeByCode: Record<string, number> = {};
    for (const row of t.rawRows) {
      const code = getCellValue(row, colIdx.chargeCode);
      if (!code) continue;
      const amt = toNumber(getRawCellValue(row, colIdx.chargeAmount));
      chargeByCode[code] = (chargeByCode[code] || 0) + amt;
    }
    for (let i = 0; i < chargeCodes.length; i++) {
      const c = chargeStartCol + i;
      ws[XLSX.utils.encode_cell({ r, c })] = { v: chargeByCode[chargeCodes[i]] ?? 0, t: 'n' };
    }

    // Total Other Charges (formula)
    if (chargeColCount > 0) {
      ws[XLSX.utils.encode_cell({ r, c: 8 })] = { f: `SUM(${colToRef(chargeStartCol, r + 1)}:${colToRef(chargeStartCol + chargeColCount - 1, r + 1)})`, t: 'n' };
    } else {
      ws[XLSX.utils.encode_cell({ r, c: 8 })] = { v: 0, t: 'n' };
    }
    ws[XLSX.utils.encode_cell({ r, c: 9 })] = { f: `${colToRef(8, r + 1)}*12`, t: 'n' };

    // Future rent steps — group by charge code
    const futureByCode: Record<string, { date: string; amount: number }[]> = {};
    for (const row of t.rawRows) {
      const date = getCellValue(row, colIdx.futureDate);
      if (!date) continue;
      const code = getCellValue(row, colIdx.chargeCode) || 'RNT';
      if (!futureByCode[code]) futureByCode[code] = [];
      const amount = toNumber(getRawCellValue(row, colIdx.futureAmount));
      futureByCode[code].push({ date, amount });
    }
    for (const block of futureCodeBlocks) {
      const steps = futureByCode[block.code] || [];
      for (let s = 0; s < block.steps; s++) {
        const dateCol = block.startCol + s * 2;
        const rateCol = block.startCol + s * 2 + 1;
        if (s < steps.length) {
          ws[XLSX.utils.encode_cell({ r, c: dateCol })] = { v: steps[s].date };
          ws[XLSX.utils.encode_cell({ r, c: rateCol })] = { v: steps[s].amount, t: 'n' };
        }
      }
    }
  }

  // ── Custom group columns ──
  let customStartCol = Math.max(
    FIXED_COUNT,
    chargeStartCol + chargeColCount + (chargeColCount > 0 ? 1 : 0),
    ...futureCodeBlocks.map(b => b.startCol + b.steps * 2 + 1),
  );

  for (const cg of customGroups) {
    const span = groupSpans.find(s => s.groupId === cg.id);
    if (!span) continue;

    // Collect column labels within the span
    const labels: string[] = [];
    const colIndices: number[] = [];
    for (let c = span.startCol; c <= span.endCol; c++) {
      labels.push(columnLabels[c] || String.fromCharCode(65 + c));
      colIndices.push(c);
    }
    if (labels.length === 0) continue;

    // Row 0: Category banner
    ws[XLSX.utils.encode_cell({ r: 0, c: customStartCol })] = { v: cg.label, s: makeCategoryStyle(COLORS.chargeCategoryBg) };
    if (labels.length > 1) {
      merges.push({ s: { r: 0, c: customStartCol }, e: { r: 0, c: customStartCol + labels.length - 1 } });
    }
    for (let i = 1; i < labels.length; i++) {
      ws[XLSX.utils.encode_cell({ r: 0, c: customStartCol + i })] = { v: '', s: makeCategoryStyle(COLORS.chargeCategoryBg) };
    }
    for (let i = 0; i < labels.length; i++) {
      ws[XLSX.utils.encode_cell({ r: 1, c: customStartCol + i })] = { v: '', s: makeHeaderStyle() };
    }
    for (let i = 0; i < labels.length; i++) {
      ws[XLSX.utils.encode_cell({ r: 2, c: customStartCol + i })] = { v: labels[i], s: makeHeaderStyle() };
    }

    // Data rows — read raw values directly
    for (let ti = 0; ti < tenants.length; ti++) {
      const t = tenants[ti];
      const r = 3 + ti;

      if (cg.collection) {
        // Concatenate values from all rows
        for (let li = 0; li < colIndices.length; li++) {
          const values = t.rawRows
            .map(row => getRawCellValue(row, colIndices[li]))
            .filter(v => v !== null && v !== undefined && String(v).trim() !== '');
          if (values.length === 1) {
            const v = values[0]!;
            const num = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,$]/g, ''));
            if (!isNaN(num)) {
              ws[XLSX.utils.encode_cell({ r, c: customStartCol + li })] = { v: num, t: 'n' };
            } else {
              ws[XLSX.utils.encode_cell({ r, c: customStartCol + li })] = { v: String(v) };
            }
          } else if (values.length > 1) {
            ws[XLSX.utils.encode_cell({ r, c: customStartCol + li })] = { v: values.map(v => String(v).trim()).join('; ') };
          }
        }
      } else {
        // Scalar: first row only
        const row = t.rawRows[0] || [];
        for (let li = 0; li < colIndices.length; li++) {
          const val = getRawCellValue(row, colIndices[li]);
          if (val !== null && val !== undefined && String(val).trim()) {
            const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/[,$]/g, ''));
            if (!isNaN(num)) {
              ws[XLSX.utils.encode_cell({ r, c: customStartCol + li })] = { v: num, t: 'n' };
            } else {
              ws[XLSX.utils.encode_cell({ r, c: customStartCol + li })] = { v: String(val).trim() };
            }
          }
        }
      }
    }

    customStartCol += labels.length + 1;
  }

  ws['!merges'] = merges;

  const maxRow = 3 + tenants.length - 1;
  const maxCol = Math.max(
    FIXED_COUNT - 1,
    chargeStartCol + chargeColCount - 1,
    ...futureCodeBlocks.map(b => b.startCol + b.steps * 2 - 1),
    customStartCol - 2,
    0
  );
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(maxRow, 2), c: Math.max(maxCol, FIXED_COUNT - 1) } });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Templatized Rent Roll');

  const outputName = fileName.replace(/\.(xlsx|xls)$/i, '') + '_template.xlsx';
  XLSX.writeFile(wb, outputName);
}
