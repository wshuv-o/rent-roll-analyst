// Final RR Export — ExcelJS with dual headers, group colors, formulas
import ExcelJS from 'exceljs';
import type { MallRentRollTenant } from './rent-roll-types/mall-rent-roll-parser';
import { DEFAULT_CHARGE_CODE_MAPPING } from './rent-roll-types/mall-rent-roll-parser';

type Cell = string | number | Date | null;

const MAPPING_BY_CODE = new Map(DEFAULT_CHARGE_CODE_MAPPING.map(m => [m.code, m]));

function computeCategoryTotal(annualByCode: Record<string, number>, category: string): number {
  let total = 0;
  for (const [code, amt] of Object.entries(annualByCode)) {
    const m = MAPPING_BY_CODE.get(code);
    if (m && m.category === category) total += amt;
  }
  return total;
}

function buildAnnualByCode(t: MallRentRollTenant): Record<string, number> {
  const byCode: Record<string, number> = {};
  for (const ch of t.charges) {
    if (!ch.billCode) continue;
    byCode[ch.billCode] = (byCode[ch.billCode] || 0) + (ch.monthlyAmount ?? 0) * 12;
  }
  return byCode;
}

function gatherAllChargeCodes(tenants: MallRentRollTenant[]): string[] {
  const codeSet = new Set<string>();
  for (const t of tenants) for (const ch of t.charges) if (ch.billCode) codeSet.add(ch.billCode);
  const knownOrder = DEFAULT_CHARGE_CODE_MAPPING.map(m => m.code);
  return [...knownOrder.filter(c => codeSet.has(c)), ...Array.from(codeSet).filter(c => !knownOrder.includes(c)).sort()];
}

function buildMappingData(codes: string[]) {
  const knownMap = new Map(DEFAULT_CHARGE_CODE_MAPPING.map(m => [m.code, m]));
  return codes.map(code => {
    const known = knownMap.get(code);
    return { code, description: known?.description ?? code, category: known?.category ?? '', reliefSubType: known?.reliefSubType ?? '' };
  });
}

function colToRef(col: number, row: number): string {
  let letter = '';
  let n = col - 1; // ExcelJS is 1-indexed
  while (n >= 0) { letter = String.fromCharCode(65 + (n % 26)) + letter; n = Math.floor(n / 26) - 1; }
  return `${letter}${row}`;
}

// Group color definitions
const GRP = {
  identity:   { bg: '1B2A4A', label: 'Identity' },
  charge:     { bg: '2D5F2D', label: 'Current Charge' },
  annual:     { bg: '4A2D6A', label: 'Annual Totals' },
  codes:      { bg: '3D4F5F', label: 'Charge Codes' },
  totals:     { bg: '4A4A4A', label: 'Totals' },
  rentBumps:  { bg: '8B4513', label: 'Rent Bumps' },
  breakpoints:{ bg: '6B2D3D', label: 'Breakpoints' },
  camBumps:   { bg: '1A4A4A', label: 'CAM Bumps' },
  utlBumps:   { bg: '1A3A5A', label: 'UTL Bumps' },
  retBumps:   { bg: '5A2D4A', label: 'RET Bumps' },
  category:   { bg: '3D4F5F', label: 'Category' },
};

const FONT_HDR: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 8, name: 'Arial' };
const FONT_DATA: Partial<ExcelJS.Font> = { size: 8, name: 'Arial' };
const BORDER_THIN: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF888888' } },
  bottom: { style: 'thin', color: { argb: 'FF888888' } },
  left: { style: 'thin', color: { argb: 'FF888888' } },
  right: { style: 'thin', color: { argb: 'FF888888' } },
};

function fillBg(hex: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + hex } };
}

function styleBanner(cell: ExcelJS.Cell, bg: string) {
  cell.font = { ...FONT_HDR, size: 9 };
  cell.fill = fillBg(bg);
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = BORDER_THIN;
}

function styleHeader(cell: ExcelJS.Cell, bg: string) {
  cell.font = FONT_HDR;
  cell.fill = fillBg(bg);
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cell.border = BORDER_THIN;
}

export async function downloadFinalRR(tenants: MallRentRollTenant[], fileName: string) {
  const wb = new ExcelJS.Workbook();
  const allCodes = gatherAllChargeCodes(tenants);
  const codeCount = allCodes.length;
  const mappingRows = buildMappingData(allCodes);

  // ── Sheet 1: DRAFT ──
  const wsDraft = wb.addWorksheet('DRAFT');
  wsDraft.addRow(['Code', 'Description', 'Mapping', 'Relief']);
  for (const m of mappingRows) wsDraft.addRow([m.code, m.description, m.category || '', m.reliefSubType || '']);
  wsDraft.getColumn(1).width = 10; wsDraft.getColumn(2).width = 28;
  wsDraft.getColumn(3).width = 14; wsDraft.getColumn(4).width = 14;
  // Style draft header
  for (let c = 1; c <= 4; c++) {
    const cell = wsDraft.getRow(1).getCell(c);
    cell.font = { bold: true, size: 8, name: 'Arial' };
    cell.fill = fillBg('333333');
    cell.font = { ...cell.font, color: { argb: 'FFFFFFFF' } };
  }

  // ── Sheet 2: Final RR ──
  const ws = wb.addWorksheet('Final RR');

  // Column layout (1-indexed for ExcelJS)
  let c = 1;
  const COL = {
    unit: c++, dba: c++, leaseId: c++, sf: c++, modifiedSf: c++,
    leaseType: c++, spaceType: c++, spaceTypeInput: c++, unitType: c++,
    leaseStatus: c++, pil: c++,
    chargeCode: c++, chargeDesc: c++,
    commence: c++, origEnd: c++, expire: c++,
    annRent: c++, rentPSF: c++,
    annCAM: c++, annRET: c++, annUTL: c++, annRelief: c++, annExcluded: c++,
  };
  const codeStart = c; c += codeCount;
  const COL2 = { total: c++, variance: c++ };

  const rentSumStart = c; c += 3; // Bump Date, Amount, UW Rent
  const rentBumpStart = c; c += 36; // 18 pairs
  const bpStart = c; c += 18; // 6 × 3
  const camSumStart = c; c += 3;
  const camBumpStart = c; c += 24; // 12 pairs
  const utlSumStart = c; c += 4; // +%
  const utlBumpStart = c; c += 24;
  const retSumStart = c; c += 3;
  const retBumpStart = c; c += 24;
  const catCol = c++;
  const totalCols = c - 1;

  // Group spans: [startCol, endCol, groupKey]
  type GS = [number, number, keyof typeof GRP];
  const groupSpans: GS[] = [
    [COL.unit, COL.pil, 'identity'],
    [COL.chargeCode, COL.annExcluded, 'charge'],
    [codeStart, codeStart + codeCount - 1, 'codes'],
    [COL2.total, COL2.variance, 'totals'],
    [rentSumStart, rentBumpStart + 35, 'rentBumps'],
    [bpStart, bpStart + 17, 'breakpoints'],
    [camSumStart, camBumpStart + 23, 'camBumps'],
    [utlSumStart, utlBumpStart + 23, 'utlBumps'],
    [retSumStart, retBumpStart + 23, 'retBumps'],
    [catCol, catCol, 'category'],
  ];

  // ── Row 1: Group banners ──
  for (const [s, e, gk] of groupSpans) {
    const g = GRP[gk];
    if (e > s) ws.mergeCells(1, s, 1, e);
    const cell = ws.getRow(1).getCell(s);
    cell.value = g.label;
    styleBanner(cell, g.bg);
    // Fill remaining cells in merged range with same style
    for (let cc = s + 1; cc <= e; cc++) styleBanner(ws.getRow(1).getCell(cc), g.bg);
  }
  ws.getRow(1).height = 20;

  // ── Row 2: Column headers ──
  const setH = (col: number, label: string, gk: keyof typeof GRP) => {
    const cell = ws.getRow(2).getCell(col);
    cell.value = label;
    styleHeader(cell, GRP[gk].bg);
  };

  setH(COL.unit, 'Units', 'identity'); setH(COL.dba, 'Tenant Name', 'identity');
  setH(COL.leaseId, 'Lease ID', 'identity'); setH(COL.sf, 'SF', 'identity');
  setH(COL.modifiedSf, 'Modified SF', 'identity'); setH(COL.leaseType, 'Lease Type', 'identity');
  setH(COL.spaceType, 'Space Type', 'identity'); setH(COL.spaceTypeInput, 'Space Type - Input', 'identity');
  setH(COL.unitType, 'Unit Type', 'identity'); setH(COL.leaseStatus, 'Lease Status', 'identity');
  setH(COL.pil, 'PIL', 'identity');

  setH(COL.chargeCode, 'Code', 'charge'); setH(COL.chargeDesc, 'Expense Description', 'charge');
  setH(COL.commence, 'Commencement Date', 'charge'); setH(COL.origEnd, 'Original End Date', 'charge');
  setH(COL.expire, 'Expire/Close Date', 'charge');
  setH(COL.annRent, 'Rent', 'charge'); setH(COL.rentPSF, 'Rent PSF', 'charge');
  setH(COL.annCAM, 'CAM', 'charge'); setH(COL.annRET, 'RET', 'charge');
  setH(COL.annUTL, 'UTL', 'charge'); setH(COL.annRelief, 'Relief', 'charge');
  setH(COL.annExcluded, 'Excluded', 'charge');

  for (let i = 0; i < codeCount; i++) setH(codeStart + i, allCodes[i], 'codes');
  setH(COL2.total, 'Total', 'totals'); setH(COL2.variance, 'Variance', 'totals');

  setH(rentSumStart, 'Bump Date', 'rentBumps'); setH(rentSumStart + 1, 'Amount', 'rentBumps');
  setH(rentSumStart + 2, 'UW Rent', 'rentBumps');
  for (let i = 0; i < 18; i++) {
    setH(rentBumpStart + i * 2, `Bump Date ${i + 1}`, 'rentBumps');
    setH(rentBumpStart + i * 2 + 1, `Bump Rent ${i + 1}`, 'rentBumps');
  }

  const bpLabels = ['Current', 'BP 1', 'BP 2', 'BP 3', 'BP 4', 'BP 5'];
  for (let i = 0; i < 6; i++) {
    setH(bpStart + i * 3, `${bpLabels[i]} BP Date`, 'breakpoints');
    setH(bpStart + i * 3 + 1, `${bpLabels[i]} Breakpoint`, 'breakpoints');
    setH(bpStart + i * 3 + 2, `${bpLabels[i]} %`, 'breakpoints');
  }

  setH(camSumStart, 'Bump Date', 'camBumps'); setH(camSumStart + 1, 'Amount', 'camBumps');
  setH(camSumStart + 2, 'Changes on CAM', 'camBumps');
  for (let i = 0; i < 12; i++) {
    setH(camBumpStart + i * 2, `CAM Bump Date ${i + 1}`, 'camBumps');
    setH(camBumpStart + i * 2 + 1, `CAM Bump Amt ${i + 1}`, 'camBumps');
  }

  setH(utlSumStart, 'Bump Date', 'utlBumps'); setH(utlSumStart + 1, 'Amount', 'utlBumps');
  setH(utlSumStart + 2, 'Changes on UTL', 'utlBumps'); setH(utlSumStart + 3, '%', 'utlBumps');
  for (let i = 0; i < 12; i++) {
    setH(utlBumpStart + i * 2, `UTL Bump Date ${i + 1}`, 'utlBumps');
    setH(utlBumpStart + i * 2 + 1, `UTL Bump Amt ${i + 1}`, 'utlBumps');
  }

  setH(retSumStart, 'Bump Date', 'retBumps'); setH(retSumStart + 1, 'Amount', 'retBumps');
  setH(retSumStart + 2, 'Changes on RET', 'retBumps');
  for (let i = 0; i < 12; i++) {
    setH(retBumpStart + i * 2, `RET Bump Date ${i + 1}`, 'retBumps');
    setH(retBumpStart + i * 2 + 1, `RET Bump Amt ${i + 1}`, 'retBumps');
  }

  setH(catCol, 'Category', 'category');
  ws.getRow(2).height = 28;

  // Mapping range for SUMPRODUCT
  const mappingCatRange = `DRAFT!$C$2:$C$${1 + codeCount}`;

  // ── Data rows (starting at row 3) ──
  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = ti + 3; // ExcelJS row (1-indexed), data starts at 3
    const row = ws.getRow(r);

    const writeCell = (col: number, v: Cell) => {
      if (v === null || v === undefined) return;
      const cell = row.getCell(col);
      if (v instanceof Date) cell.value = v;
      else if (typeof v === 'number') cell.value = v;
      else cell.value = String(v);
      cell.font = FONT_DATA;
      cell.border = BORDER_THIN;
    };

    // Identity
    writeCell(COL.unit, t.unit); writeCell(COL.dba, t.dba);
    writeCell(COL.leaseId, t.leaseId); writeCell(COL.sf, t.squareFootage);
    writeCell(COL.modifiedSf, t.squareFootage); writeCell(COL.leaseType, t.leaseType);
    writeCell(COL.spaceType, t.category); writeCell(COL.spaceTypeInput, t.category);
    writeCell(COL.unitType, t.unitType); writeCell(COL.leaseStatus, t.leaseStatus);
    writeCell(COL.pil, t.percentInLieu);

    if (t.charges.length > 0) {
      writeCell(COL.chargeCode, t.charges[0].billCode);
      writeCell(COL.chargeDesc, t.charges[0].expenseDescription);
    }
    writeCell(COL.commence, t.commencementDate);
    writeCell(COL.origEnd, t.originalEndDate);
    writeCell(COL.expire, t.expireCloseDate);

    // Individual charge codes
    const annualByCode = buildAnnualByCode(t);
    for (let i = 0; i < allCodes.length; i++) {
      writeCell(codeStart + i, annualByCode[allCodes[i]] ?? 0);
    }

    // Category totals — SUMPRODUCT formulas
    if (codeCount > 0) {
      const cRange = `${colToRef(codeStart, r)}:${colToRef(codeStart + codeCount - 1, r)}`;
      const setFormula = (col: number, formula: string) => {
        const cell = row.getCell(col);
        cell.value = { formula } as ExcelJS.CellFormulaValue;
        cell.font = FONT_DATA;
        cell.border = BORDER_THIN;
      };
      setFormula(COL.annRent, `SUMPRODUCT((${mappingCatRange}="Rent")*(${cRange}))`);
      setFormula(COL.rentPSF, `IF(${colToRef(COL.sf, r)}=0,"",${colToRef(COL.annRent, r)}/${colToRef(COL.sf, r)})`);
      setFormula(COL.annCAM, `SUMPRODUCT((${mappingCatRange}="CAM")*(${cRange}))`);
      setFormula(COL.annRET, `SUMPRODUCT((${mappingCatRange}="RET")*(${cRange}))`);
      setFormula(COL.annUTL, `SUMPRODUCT((${mappingCatRange}="UTL")*(${cRange}))`);
      setFormula(COL.annRelief, `SUMPRODUCT((${mappingCatRange}="Relief")*(${cRange}))`);
      setFormula(COL.annExcluded, `SUMPRODUCT((${mappingCatRange}="Excluded")*(${cRange}))`);
      setFormula(COL2.total, `SUM(${colToRef(codeStart, r)}:${colToRef(codeStart + codeCount - 1, r)})`);
      setFormula(COL2.variance, `${colToRef(COL2.total, r)}-${colToRef(COL.annRent, r)}-${colToRef(COL.annCAM, r)}-${colToRef(COL.annRET, r)}-${colToRef(COL.annUTL, r)}-${colToRef(COL.annRelief, r)}-${colToRef(COL.annExcluded, r)}`);
    }

    // Rent bump summary
    if (t.rentBumps.length > 0 && t.rentBumps[0]?.date) {
      writeCell(rentSumStart, t.rentBumps[0].date);
      writeCell(rentSumStart + 1, t.rentBumps[0].amount);
      const rate = typeof t.rentBumps[0].amount === 'number' ? t.rentBumps[0].amount : null;
      if (rate !== null && t.squareFootage) writeCell(rentSumStart + 2, rate * t.squareFootage);
    }
    for (let i = 0; i < 18 && i < t.rentBumps.length; i++) {
      writeCell(rentBumpStart + i * 2, t.rentBumps[i].date);
      writeCell(rentBumpStart + i * 2 + 1, t.rentBumps[i].amount);
    }

    // Breakpoints
    for (let i = 0; i < 6 && i < t.breakpoints.length; i++) {
      writeCell(bpStart + i * 3, t.breakpoints[i].date);
      writeCell(bpStart + i * 3 + 1, t.breakpoints[i].amount);
      writeCell(bpStart + i * 3 + 2, t.breakpoints[i].percent);
    }

    // CAM
    if (t.camBumps.length > 0 && t.camBumps[0]?.date) {
      writeCell(camSumStart, t.camBumps[0].date);
      const a = typeof t.camBumps[0].amount === 'number' ? t.camBumps[0].amount : null;
      writeCell(camSumStart + 1, a);
      if (a !== null) writeCell(camSumStart + 2, a - computeCategoryTotal(annualByCode, 'CAM'));
    }
    for (let i = 0; i < 12 && i < t.camBumps.length; i++) {
      writeCell(camBumpStart + i * 2, t.camBumps[i].date);
      writeCell(camBumpStart + i * 2 + 1, t.camBumps[i].amount);
    }

    // UTL
    if (t.utlBumps.length > 0 && t.utlBumps[0]?.date) {
      writeCell(utlSumStart, t.utlBumps[0].date);
      const a = typeof t.utlBumps[0].amount === 'number' ? t.utlBumps[0].amount : null;
      writeCell(utlSumStart + 1, a);
      if (a !== null) writeCell(utlSumStart + 2, a - computeCategoryTotal(annualByCode, 'UTL'));
      writeCell(utlSumStart + 3, t.utlBumps[0].percent ?? null);
    }
    for (let i = 0; i < 12 && i < t.utlBumps.length; i++) {
      writeCell(utlBumpStart + i * 2, t.utlBumps[i].date);
      writeCell(utlBumpStart + i * 2 + 1, t.utlBumps[i].amount);
    }

    // RET
    if (t.retBumps.length > 0 && t.retBumps[0]?.date) {
      writeCell(retSumStart, t.retBumps[0].date);
      const a = typeof t.retBumps[0].amount === 'number' ? t.retBumps[0].amount : null;
      writeCell(retSumStart + 1, a);
      if (a !== null) writeCell(retSumStart + 2, a - computeCategoryTotal(annualByCode, 'RET'));
    }
    for (let i = 0; i < 12 && i < t.retBumps.length; i++) {
      writeCell(retBumpStart + i * 2, t.retBumps[i].date);
      writeCell(retBumpStart + i * 2 + 1, t.retBumps[i].amount);
    }

    writeCell(catCol, t.category);
  }

  // Column widths
  for (let cc = 1; cc <= totalCols; cc++) {
    const col = ws.getColumn(cc);
    if (cc === COL.dba || cc === COL.chargeDesc) col.width = 24;
    else if (cc === COL.unit || cc === COL.leaseId) col.width = 11;
    else col.width = 12;
  }

  // Freeze panes: freeze first 2 rows + first 3 columns
  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 2 }];

  // Write and download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.replace(/\.[^.]+$/, '') + '_Final_RR.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}
