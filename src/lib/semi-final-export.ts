// Semi Final Download — single sheet, values only, one line per tenant
// Uses ExcelJS for proper header styling with group colors
import ExcelJS from 'exceljs';
import type { MallRentRollTenant } from './rent-roll-types/mall-rent-roll-parser';

type Cell = string | number | Date | null;

const GROUP_COLORS: Record<string, string> = {
  identity: '1B2A4A',
  charges: '2D5F2D',
  total: '4A4A4A',
  future: '8B4513',
  overage: '4A2D6A',
};

const FONT_HDR: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 8, name: 'Arial' };
const FONT_BANNER: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9, name: 'Arial' };
const FONT_DATA: Partial<ExcelJS.Font> = { size: 8, name: 'Arial' };
const BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF888888' } },
  bottom: { style: 'thin', color: { argb: 'FF888888' } },
  left: { style: 'thin', color: { argb: 'FF888888' } },
  right: { style: 'thin', color: { argb: 'FF888888' } },
};

function fillBg(hex: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + hex } };
}

function fmtCell(v: Cell): string | number {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-US');
  if (typeof v === 'number') {
    if (v > 20000 && v < 60000) {
      const d = new Date(1899, 11, 30);
      d.setDate(d.getDate() + v);
      if (!isNaN(d.getTime())) return d.toLocaleDateString('en-US');
    }
    return v;
  }
  return String(v).trim();
}

export async function downloadSemiFinalRR(tenants: MallRentRollTenant[], fileName: string): Promise<void> {
  let maxCharges = 0, maxFuture = 0, maxOverage = 0;
  for (const t of tenants) {
    maxCharges = Math.max(maxCharges, t.charges.length);
    maxFuture = Math.max(maxFuture, t.futureEscalations.length);
    maxOverage = Math.max(maxOverage, t.overageEntries.length);
  }
  if (maxCharges === 0) maxCharges = 1;
  if (maxFuture === 0) maxFuture = 1;
  if (maxOverage === 0) maxOverage = 1;

  const identityHeaders = [
    'Unit', 'DBA', 'Lease ID', 'Square Footage', 'Lease Type',
    'Unit Type', 'Lease Status', '% In Lieu', 'Space Type',
    'Commencement Date', 'Open Date', 'Original End Date', 'Expire/Close Date',
  ];
  const IDENTITY_COUNT = identityHeaders.length;
  const chargeFields = ['Bill Code', 'Expense Description', 'Begin Date', 'End Date', 'Monthly Amount', 'Annual Rate/SF'];
  const CHARGE_SET = chargeFields.length;
  const chargeColCount = maxCharges * CHARGE_SET;
  const TOTAL_COUNT = 1;
  const futureFields = ['Bill Code', 'Expense Description', 'Begin Date', 'End Date', 'Monthly Amount', 'Annual Rate/SF', '% Inc.'];
  const FUTURE_SET = futureFields.length;
  const futureColCount = maxFuture * FUTURE_SET;
  const overageFields = ['Bill Code', 'Begin Date', 'End Date', 'Breakpoint', 'Overage %'];
  const OVERAGE_SET = overageFields.length;
  const overageColCount = maxOverage * OVERAGE_SET;

  // Section boundaries (1-indexed for ExcelJS)
  const chargeStart = IDENTITY_COUNT + 1;
  const totalStart = chargeStart + chargeColCount;
  const futureStart = totalStart + TOTAL_COUNT;
  const overageStart = futureStart + futureColCount;
  const totalCols = IDENTITY_COUNT + chargeColCount + TOTAL_COUNT + futureColCount + overageColCount;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Semi Final');

  // ── Row 1: Group banners ──
  const writeBanner = (startCol: number, count: number, label: string, color: string) => {
    if (count > 1) ws.mergeCells(1, startCol, 1, startCol + count - 1);
    const cell = ws.getRow(1).getCell(startCol);
    cell.value = label;
    cell.font = FONT_BANNER;
    cell.fill = fillBg(color);
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER;
    for (let cc = startCol + 1; cc < startCol + count; cc++) {
      const c = ws.getRow(1).getCell(cc);
      c.fill = fillBg(color);
      c.border = BORDER;
    }
  };

  writeBanner(1, IDENTITY_COUNT, 'Identity', GROUP_COLORS.identity);
  writeBanner(chargeStart, chargeColCount, 'Current Charges', GROUP_COLORS.charges);
  writeBanner(totalStart, TOTAL_COUNT, '', GROUP_COLORS.total);
  writeBanner(futureStart, futureColCount, 'Future Rent & Expense Escalations', GROUP_COLORS.future);
  writeBanner(overageStart, overageColCount, 'Overage/% In Lieu Rent Terms', GROUP_COLORS.overage);
  ws.getRow(1).height = 20;

  // ── Row 2: Column headers ──
  const setH = (col: number, label: string, color: string) => {
    const cell = ws.getRow(2).getCell(col);
    cell.value = label;
    cell.font = FONT_HDR;
    cell.fill = fillBg(color);
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = BORDER;
  };

  for (let i = 0; i < IDENTITY_COUNT; i++) setH(i + 1, identityHeaders[i], GROUP_COLORS.identity);
  for (let s = 0; s < maxCharges; s++)
    for (let f = 0; f < CHARGE_SET; f++)
      setH(chargeStart + s * CHARGE_SET + f, chargeFields[f], GROUP_COLORS.charges);
  setH(totalStart, 'Total', GROUP_COLORS.total);
  for (let s = 0; s < maxFuture; s++)
    for (let f = 0; f < FUTURE_SET; f++)
      setH(futureStart + s * FUTURE_SET + f, futureFields[f], GROUP_COLORS.future);
  for (let s = 0; s < maxOverage; s++)
    for (let f = 0; f < OVERAGE_SET; f++)
      setH(overageStart + s * OVERAGE_SET + f, overageFields[f], GROUP_COLORS.overage);
  ws.getRow(2).height = 28;

  // ── Data rows ──
  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = ti + 3;
    const row = ws.getRow(r);

    const write = (col: number, v: Cell) => {
      const fv = fmtCell(v);
      if (fv === '' || fv === null || fv === undefined) return;
      const cell = row.getCell(col);
      cell.value = typeof fv === 'number' ? fv : fv;
      cell.font = FONT_DATA;
      cell.border = BORDER;
    };

    // Identity (1-indexed)
    write(1, t.unit); write(2, t.dba); write(3, t.leaseId);
    write(4, t.squareFootage); write(5, t.leaseType);
    write(6, t.unitType); write(7, t.leaseStatus);
    write(8, t.percentInLieu); write(9, t.category);
    write(10, t.commencementDate); write(11, t.openDate);
    write(12, t.originalEndDate); write(13, t.expireCloseDate);

    // Current charges
    for (let ci = 0; ci < t.charges.length; ci++) {
      const ch = t.charges[ci];
      const base = chargeStart + ci * CHARGE_SET;
      write(base, ch.billCode); write(base + 1, ch.expenseDescription);
      write(base + 2, ch.beginDate); write(base + 3, ch.endDate);
      write(base + 4, ch.monthlyAmount); write(base + 5, ch.annualRateSF);
    }

    write(totalStart, t.totalMonthlyAmount);

    // Future escalations
    for (let fi = 0; fi < t.futureEscalations.length; fi++) {
      const fe = t.futureEscalations[fi];
      const base = futureStart + fi * FUTURE_SET;
      write(base, fe.billCode); write(base + 1, fe.expenseDescription);
      write(base + 2, fe.beginDate); write(base + 3, fe.endDate);
      write(base + 4, fe.monthlyAmount); write(base + 5, fe.annualRateSF);
      write(base + 6, fe.percentInc);
    }

    // Overage entries
    for (let oi = 0; oi < t.overageEntries.length; oi++) {
      const oe = t.overageEntries[oi];
      const base = overageStart + oi * OVERAGE_SET;
      write(base, oe.billCode); write(base + 1, oe.beginDate);
      write(base + 2, oe.endDate); write(base + 3, oe.breakpoint);
      write(base + 4, oe.percent);
    }
  }

  // Column widths
  for (let cc = 1; cc <= totalCols; cc++) {
    const col = ws.getColumn(cc);
    if (cc === 2) col.width = 24; // DBA
    else col.width = 13;
  }

  // Freeze panes
  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 2 }];

  // Write and download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.replace(/\.[^.]+$/, '') + '_SemiFinal.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}
