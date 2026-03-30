// Semi Final Download — single sheet, values only, one line per tenant
// Groups: Identity, Current Charges, Future Rent & Expense Escalations, Overage/% In Lieu
import * as XLSX from 'xlsx';
import type { MallRentRollTenant } from './rent-roll-types/mall-rent-roll-parser';

type Cell = string | number | Date | null;

const GROUP_COLORS: Record<string, string> = {
  identity: 'FF1B2A4A',
  charges: 'FF2D5F2D',
  future: 'FF5C3D1A',
  overage: 'FF4A2D6A',
  legal: 'FF1A4A4A',
};

function headerStyle(bg: string): Record<string, unknown> {
  return {
    font: { bold: true, color: { rgb: 'FFFFFFFF' }, sz: 10 },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top: { style: 'thin', color: { rgb: 'FF30363D' } },
      bottom: { style: 'thin', color: { rgb: 'FF30363D' } },
      left: { style: 'thin', color: { rgb: 'FF30363D' } },
      right: { style: 'thin', color: { rgb: 'FF30363D' } },
    },
  };
}

function groupBannerStyle(bg: string): Record<string, unknown> {
  return {
    font: { bold: true, color: { rgb: 'FFFFFFFF' }, sz: 11 },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: 'FF30363D' } },
      bottom: { style: 'thin', color: { rgb: 'FF30363D' } },
      left: { style: 'thin', color: { rgb: 'FF30363D' } },
      right: { style: 'thin', color: { rgb: 'FF30363D' } },
    },
  };
}

function fmtCell(v: Cell): string | number {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-US');
  if (typeof v === 'number') {
    // Excel serial date
    if (v > 20000 && v < 60000) {
      const d = new Date(1899, 11, 30);
      d.setDate(d.getDate() + v);
      if (!isNaN(d.getTime())) return d.toLocaleDateString('en-US');
    }
    return v;
  }
  return String(v).trim();
}

export function downloadSemiFinalRR(tenants: MallRentRollTenant[], fileName: string): void {
  // Determine max repeats for each section
  let maxCharges = 0, maxFuture = 0, maxOverage = 0;
  for (const t of tenants) {
    maxCharges = Math.max(maxCharges, t.charges.length);
    maxFuture = Math.max(maxFuture, t.futureEscalations.length);
    maxOverage = Math.max(maxOverage, t.overageEntries.length);
  }
  if (maxCharges === 0) maxCharges = 1;
  if (maxFuture === 0) maxFuture = 1;
  if (maxOverage === 0) maxOverage = 1;

  // Identity columns
  const identityHeaders = [
    'Unit', 'DBA', 'Lease ID', 'Square Footage', 'Lease Type',
    'Unit Type', 'Lease Status', '% In Lieu', 'Space Type',
    'Commencement Date', 'Open Date', 'Original End Date', 'Expire/Close Date',
  ];
  const IDENTITY_COUNT = identityHeaders.length;

  // Per-charge columns
  const chargeFields = ['Bill Code', 'Expense Description', 'Begin Date', 'End Date', 'Monthly Amount', 'Annual Rate/SF'];
  const CHARGE_SET = chargeFields.length;
  const chargeColCount = maxCharges * CHARGE_SET;

  // Total column between charges and future
  const TOTAL_COUNT = 1; // "Total"

  // Per-future columns
  const futureFields = ['Bill Code', 'Expense Description', 'Begin Date', 'End Date', 'Monthly Amount', 'Annual Rate/SF', '% Inc.'];
  const FUTURE_SET = futureFields.length;
  const futureColCount = maxFuture * FUTURE_SET;

  // Per-overage columns
  const overageFields = ['Bill Code', 'Begin Date', 'End Date', 'Breakpoint', 'Overage %'];
  const OVERAGE_SET = overageFields.length;
  const overageColCount = maxOverage * OVERAGE_SET;

  const totalCols = IDENTITY_COUNT + chargeColCount + TOTAL_COUNT + futureColCount + overageColCount;

  // Section boundaries
  const chargeStart = IDENTITY_COUNT;
  const totalStart = chargeStart + chargeColCount;
  const futureStart = totalStart + TOTAL_COUNT;
  const overageStart = futureStart + futureColCount;

  const ws: XLSX.WorkSheet = {};
  const merges: XLSX.Range[] = [];

  // ── Row 0: Group banners ──
  const writeBanner = (startCol: number, count: number, label: string, color: string) => {
    ws[XLSX.utils.encode_cell({ r: 0, c: startCol })] = { v: label, s: groupBannerStyle(color) };
    for (let c = startCol + 1; c < startCol + count; c++) {
      ws[XLSX.utils.encode_cell({ r: 0, c })] = { v: '', s: groupBannerStyle(color) };
    }
    if (count > 1) merges.push({ s: { r: 0, c: startCol }, e: { r: 0, c: startCol + count - 1 } });
  };

  writeBanner(0, IDENTITY_COUNT, '', GROUP_COLORS.identity); // Identity has no banner text, merged with header
  writeBanner(chargeStart, chargeColCount + TOTAL_COUNT, 'Current Charges', GROUP_COLORS.charges);
  writeBanner(futureStart, futureColCount, 'Future Rent & Expense Escalations', GROUP_COLORS.future);
  writeBanner(overageStart, overageColCount, 'Overage/% In Lieu Rent Terms', GROUP_COLORS.overage);

  // ── Row 1: Column headers ──
  // Identity
  for (let i = 0; i < IDENTITY_COUNT; i++) {
    ws[XLSX.utils.encode_cell({ r: 1, c: i })] = { v: identityHeaders[i], s: headerStyle(GROUP_COLORS.identity) };
  }
  // Current charges (repeated per charge set)
  for (let s = 0; s < maxCharges; s++) {
    for (let f = 0; f < CHARGE_SET; f++) {
      ws[XLSX.utils.encode_cell({ r: 1, c: chargeStart + s * CHARGE_SET + f })] = {
        v: chargeFields[f], s: headerStyle(GROUP_COLORS.charges),
      };
    }
  }
  // Total
  ws[XLSX.utils.encode_cell({ r: 1, c: totalStart })] = { v: 'Total', s: headerStyle(GROUP_COLORS.charges) };
  // Future (repeated)
  for (let s = 0; s < maxFuture; s++) {
    for (let f = 0; f < FUTURE_SET; f++) {
      ws[XLSX.utils.encode_cell({ r: 1, c: futureStart + s * FUTURE_SET + f })] = {
        v: futureFields[f], s: headerStyle(GROUP_COLORS.future),
      };
    }
  }
  // Overage (repeated)
  for (let s = 0; s < maxOverage; s++) {
    for (let f = 0; f < OVERAGE_SET; f++) {
      ws[XLSX.utils.encode_cell({ r: 1, c: overageStart + s * OVERAGE_SET + f })] = {
        v: overageFields[f], s: headerStyle(GROUP_COLORS.overage),
      };
    }
  }

  // ── Data rows ──
  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = ti + 2;

    const write = (c: number, v: Cell) => {
      const fv = fmtCell(v);
      if (fv === '' || fv === null || fv === undefined) return;
      if (typeof fv === 'number') {
        ws[XLSX.utils.encode_cell({ r, c })] = { v: fv, t: 'n' };
      } else {
        ws[XLSX.utils.encode_cell({ r, c })] = { v: fv, t: 's' };
      }
    };

    // Identity
    write(0, t.unit);
    write(1, t.dba);
    write(2, t.leaseId);
    write(3, t.squareFootage);
    write(4, t.leaseType);
    write(5, t.unitType);
    write(6, t.leaseStatus);
    write(7, t.percentInLieu);
    write(8, t.category);
    write(9, t.commencementDate);
    write(10, t.openDate);
    write(11, t.originalEndDate);
    write(12, t.expireCloseDate);

    // Current charges
    for (let ci = 0; ci < t.charges.length; ci++) {
      const ch = t.charges[ci];
      const base = chargeStart + ci * CHARGE_SET;
      write(base, ch.billCode);
      write(base + 1, ch.expenseDescription);
      write(base + 2, ch.beginDate);
      write(base + 3, ch.endDate);
      write(base + 4, ch.monthlyAmount);
      write(base + 5, ch.annualRateSF);
    }

    // Total
    write(totalStart, t.totalMonthlyAmount);

    // Future escalations
    for (let fi = 0; fi < t.futureEscalations.length; fi++) {
      const fe = t.futureEscalations[fi];
      const base = futureStart + fi * FUTURE_SET;
      write(base, fe.billCode);
      write(base + 1, fe.expenseDescription);
      write(base + 2, fe.beginDate);
      write(base + 3, fe.endDate);
      write(base + 4, fe.monthlyAmount);
      write(base + 5, fe.annualRateSF);
      write(base + 6, fe.percentInc);
    }

    // Overage entries
    for (let oi = 0; oi < t.overageEntries.length; oi++) {
      const oe = t.overageEntries[oi];
      const base = overageStart + oi * OVERAGE_SET;
      write(base, oe.billCode);
      write(base + 1, oe.beginDate);
      write(base + 2, oe.endDate);
      write(base + 3, oe.breakpoint);
      write(base + 4, oe.percent);
    }
  }

  // Set range & merges
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: tenants.length + 1, c: totalCols - 1 } });
  ws['!merges'] = merges;

  // Column widths
  const colWidths: XLSX.ColInfo[] = [];
  for (let c = 0; c < totalCols; c++) {
    if (c === 1) colWidths.push({ wch: 28 }); // DBA
    else if (c <= 2) colWidths.push({ wch: 14 });
    else colWidths.push({ wch: 14 });
  }
  ws['!cols'] = colWidths;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Semi Final');

  const outName = fileName.replace(/\.[^.]+$/, '') + '_SemiFinal.xlsx';
  XLSX.writeFile(wb, outName);
}
