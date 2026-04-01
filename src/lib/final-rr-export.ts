// Final RR Export — ExcelJS with dual headers, group colors, formulas
import ExcelJS from 'exceljs';
import type { MallRentRollTenant } from './rent-roll-types/mall-rent-roll-parser';
import { DEFAULT_CHARGE_CODE_MAPPING } from './rent-roll-types/mall-rent-roll-parser';

type Cell = string | number | Date | null;

const NUM_FMT = '#,##0.00';
const NUM_FMT_INT = '#,##0';
const DATE_FMT = 'mm/dd/yyyy';

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
  let n = col - 1;
  while (n >= 0) { letter = String.fromCharCode(65 + (n % 26)) + letter; n = Math.floor(n / 26) - 1; }
  return `${letter}${row}`;
}

function colToLetter(col: number): string {
  let letter = '';
  let n = col - 1;
  while (n >= 0) { letter = String.fromCharCode(65 + (n % 26)) + letter; n = Math.floor(n / 26) - 1; }
  return letter;
}

function buildAnnualByCode(t: MallRentRollTenant): Record<string, number> {
  const byCode: Record<string, number> = {};
  for (const ch of t.charges) {
    if (!ch.billCode) continue;
    byCode[ch.billCode] = (byCode[ch.billCode] || 0) + (ch.monthlyAmount ?? 0) * 12;
  }
  return byCode;
}

// Group color definitions
const GRP = {
  identity:    { bg: '1f4e78', label: 'Identity' },
  charge:      { bg: '2D5F2D', label: 'Current Charge' },
  annual:      { bg: '4A2D6A', label: 'Annual Totals' },
  codes:       { bg: '3D4F5F', label: 'Charge Codes' },
  totals:      { bg: '4A4A4A', label: 'Totals' },
  rentBumps:   { bg: '8B4513', label: 'Rent Bumps' },
  breakpoints: { bg: '6B2D3D', label: 'Breakpoints' },
  camBumps:    { bg: '1A4A4A', label: 'CAM Bumps' },
  utlBumps:    { bg: '1A3A5A', label: 'UTL Bumps' },
  retBumps:    { bg: '5A2D4A', label: 'RET Bumps' },
  category:    { bg: '3D4F5F', label: 'Category' },
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

/** Write a value to a cell with explicit number formatting to prevent date auto-conversion */
function writeCell(row: ExcelJS.Row, col: number, v: Cell, fmt?: string) {
  if (v === null || v === undefined) return;
  const cell = row.getCell(col);
  if (v instanceof Date) {
    cell.value = v;
    cell.numFmt = DATE_FMT;
  } else if (typeof v === 'number') {
    cell.value = v;
    // Always set number format on numeric cells to prevent Excel from guessing "date"
    cell.numFmt = fmt || NUM_FMT;
  } else {
    cell.value = String(v);
  }
  cell.font = FONT_DATA;
  cell.border = BORDER_THIN;
}

/** Write a formula to a cell */
function writeFormula(row: ExcelJS.Row, col: number, formula: string, fmt?: string) {
  const cell = row.getCell(col);
  cell.value = { formula } as ExcelJS.CellFormulaValue;
  cell.numFmt = fmt || NUM_FMT;
  cell.font = FONT_DATA;
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
  for (let c = 1; c <= 4; c++) {
    const cell = wsDraft.getRow(1).getCell(c);
    cell.font = { bold: true, size: 8, name: 'Arial', color: { argb: 'FFFFFFFF' } };
    cell.fill = fillBg('333333');
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
  const camSumStart = c; c += 4; // Bump Date, Amount, Changes on CAM, %
  const camBumpStart = c; c += 24; // 12 pairs
  const utlSumStart = c; c += 4; // Bump Date, Amount, Changes on UTL, %
  const utlBumpStart = c; c += 24;
  const retSumStart = c; c += 3; // Bump Date, Amount, Changes on RET
  const retBumpStart = c; c += 24;
  const catCol = c++;
  const totalCols = c - 1;

  // DATA_START is the first data row (row 3, after 2 header rows)
  const DATA_START = 3;
  const DATA_END = DATA_START + tenants.length - 1;

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
  setH(camSumStart + 2, 'Changes on CAM', 'camBumps'); setH(camSumStart + 3, '%', 'camBumps');
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

  // DRAFT mapping range for SUMPRODUCT
  const mappingCatRange = `DRAFT!$C$2:$C$${1 + codeCount}`;

  // ── Data rows (starting at row 3) ──
  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = ti + DATA_START;
    const row = ws.getRow(r);

    // Identity
    writeCell(row, COL.unit, t.unit);
    writeCell(row, COL.dba, t.dba);
    writeCell(row, COL.leaseId, t.leaseId);
    writeCell(row, COL.sf, t.squareFootage, NUM_FMT_INT);
    writeCell(row, COL.modifiedSf, t.squareFootage, NUM_FMT_INT);
    writeCell(row, COL.leaseType, t.leaseType);
    writeCell(row, COL.spaceType, t.category);
    writeCell(row, COL.spaceTypeInput, t.category);
    writeCell(row, COL.unitType, t.unitType);
    writeCell(row, COL.leaseStatus, t.leaseStatus);
    writeCell(row, COL.pil, t.percentInLieu);

    if (t.charges.length > 0) {
      writeCell(row, COL.chargeCode, t.charges[0].billCode);
      writeCell(row, COL.chargeDesc, t.charges[0].expenseDescription);
    }
    writeCell(row, COL.commence, t.commencementDate);
    writeCell(row, COL.origEnd, t.originalEndDate);
    writeCell(row, COL.expire, t.expireCloseDate);

    // Individual charge codes — always write as number with explicit format
    const annualByCode = buildAnnualByCode(t);
    for (let i = 0; i < allCodes.length; i++) {
      const val = annualByCode[allCodes[i]] ?? 0;
      writeCell(row, codeStart + i, val, NUM_FMT);
    }

    // ── Formulas for category totals (linked to DRAFT mapping) ──
    if (codeCount > 0) {
      const cRange = `${colToRef(codeStart, r)}:${colToRef(codeStart + codeCount - 1, r)}`;

      // Rent = SUMPRODUCT of charge codes where DRAFT mapping = "Rent"
      writeFormula(row, COL.annRent, `SUMPRODUCT((${mappingCatRange}="Rent")*(${cRange}))`, NUM_FMT);

      // Rent PSF = IF(ModifiedSF=0,0,Rent/ModifiedSF)
      writeFormula(row, COL.rentPSF,
        `IF(${colToRef(COL.modifiedSf, r)}=0,0,IFERROR(${colToRef(COL.annRent, r)}/${colToRef(COL.modifiedSf, r)},0))`,
        NUM_FMT);

      // CAM, RET, UTL, Relief, Excluded
      writeFormula(row, COL.annCAM, `SUMPRODUCT((${mappingCatRange}="CAM")*(${cRange}))`, NUM_FMT);
      writeFormula(row, COL.annRET, `SUMPRODUCT((${mappingCatRange}="RET")*(${cRange}))`, NUM_FMT);
      writeFormula(row, COL.annUTL, `SUMPRODUCT((${mappingCatRange}="UTL")*(${cRange}))`, NUM_FMT);
      writeFormula(row, COL.annRelief, `SUMPRODUCT((${mappingCatRange}="Relief")*(${cRange}))`, NUM_FMT);
      writeFormula(row, COL.annExcluded, `SUMPRODUCT((${mappingCatRange}="Excluded")*(${cRange}))`, NUM_FMT);

      // Total = SUM of all charge codes
      writeFormula(row, COL2.total, `SUM(${cRange})`, NUM_FMT);

      // Variance = Total - Rent - CAM - RET - UTL - Relief - Excluded
      const rentRef = colToRef(COL.annRent, r);
      const camRef = colToRef(COL.annCAM, r);
      const retRef = colToRef(COL.annRET, r);
      const utlRef = colToRef(COL.annUTL, r);
      const reliefRef = colToRef(COL.annRelief, r);
      const exclRef = colToRef(COL.annExcluded, r);
      writeFormula(row, COL2.variance,
        `${colToRef(COL2.total, r)}-${rentRef}-${camRef}-${retRef}-${utlRef}-${reliefRef}-${exclRef}`,
        NUM_FMT);
    }

    // ── Rent bumps ──
    for (let i = 0; i < 18 && i < t.rentBumps.length; i++) {
      writeCell(row, rentBumpStart + i * 2, t.rentBumps[i].date);
      writeCell(row, rentBumpStart + i * 2 + 1, t.rentBumps[i].amount, NUM_FMT);
    }

    // Rent bump summary: Bump Date, Amount, UW Rent (formulas)
    if (t.rentBumps.length > 0) {
      // Bump Date summary = first bump date
      writeCell(row, rentSumStart, t.rentBumps[0]?.date);

      // Amount summary = first bump amount
      writeCell(row, rentSumStart + 1, t.rentBumps[0]?.amount, NUM_FMT);

      // UW Rent = IF(bumpAmount="", AnnualRent, bumpAmount * ModifiedSF)
      const bumpAmtRef = colToRef(rentSumStart + 1, r);
      const sfRef = colToRef(COL.modifiedSf, r);
      const annRentRef = colToRef(COL.annRent, r);
      writeFormula(row, rentSumStart + 2,
        `IF(${bumpAmtRef}="",${annRentRef},${bumpAmtRef}*${sfRef})`,
        NUM_FMT);
    }

    // ── Breakpoints ──
    for (let i = 0; i < 6 && i < t.breakpoints.length; i++) {
      writeCell(row, bpStart + i * 3, t.breakpoints[i].date);
      writeCell(row, bpStart + i * 3 + 1, t.breakpoints[i].amount, NUM_FMT);
      writeCell(row, bpStart + i * 3 + 2, t.breakpoints[i].percent, '0.00%');
    }

    // ── CAM bumps ──
    for (let i = 0; i < 12 && i < t.camBumps.length; i++) {
      writeCell(row, camBumpStart + i * 2, t.camBumps[i].date);
      writeCell(row, camBumpStart + i * 2 + 1, t.camBumps[i].amount, NUM_FMT);
    }
    // CAM summary with formulas
    if (t.camBumps.length > 0 && t.camBumps[0]?.date) {
      writeCell(row, camSumStart, t.camBumps[0].date);
      writeCell(row, camSumStart + 1, t.camBumps[0].amount, NUM_FMT);
      // Changes on CAM = bumpAmount - current CAM
      const camBumpAmtRef = colToRef(camSumStart + 1, r);
      const curCamRef = colToRef(COL.annCAM, r);
      writeFormula(row, camSumStart + 2,
        `IF(${camBumpAmtRef}="","",${camBumpAmtRef}-${curCamRef})`,
        NUM_FMT);
      // CAM % = (bumpAmount - currentCAM) / bumpAmount
      const firstCamBumpAmt = colToRef(camBumpStart + 1, r);
      writeFormula(row, camSumStart + 3,
        `IF(${firstCamBumpAmt}="","",(${firstCamBumpAmt}-${curCamRef})/${firstCamBumpAmt})`,
        '0.00%');
    }

    // ── UTL bumps ──
    for (let i = 0; i < 12 && i < t.utlBumps.length; i++) {
      writeCell(row, utlBumpStart + i * 2, t.utlBumps[i].date);
      writeCell(row, utlBumpStart + i * 2 + 1, t.utlBumps[i].amount, NUM_FMT);
    }
    // UTL summary with formulas
    if (t.utlBumps.length > 0 && t.utlBumps[0]?.date) {
      writeCell(row, utlSumStart, t.utlBumps[0].date);
      writeCell(row, utlSumStart + 1, t.utlBumps[0].amount, NUM_FMT);
      const utlBumpAmtRef = colToRef(utlSumStart + 1, r);
      const curUtlRef = colToRef(COL.annUTL, r);
      writeFormula(row, utlSumStart + 2,
        `IF(${utlBumpAmtRef}="","",${utlBumpAmtRef}-${curUtlRef})`,
        NUM_FMT);
      const firstUtlBumpAmt = colToRef(utlBumpStart + 1, r);
      writeFormula(row, utlSumStart + 3,
        `IF(${firstUtlBumpAmt}="","",(${firstUtlBumpAmt}-${curUtlRef})/${firstUtlBumpAmt})`,
        '0.00%');
    }

    // ── RET bumps ──
    for (let i = 0; i < 12 && i < t.retBumps.length; i++) {
      writeCell(row, retBumpStart + i * 2, t.retBumps[i].date);
      writeCell(row, retBumpStart + i * 2 + 1, t.retBumps[i].amount, NUM_FMT);
    }
    // RET summary with formulas
    if (t.retBumps.length > 0 && t.retBumps[0]?.date) {
      writeCell(row, retSumStart, t.retBumps[0].date);
      writeCell(row, retSumStart + 1, t.retBumps[0].amount, NUM_FMT);
      const retBumpAmtRef = colToRef(retSumStart + 1, r);
      const curRetRef = colToRef(COL.annRET, r);
      writeFormula(row, retSumStart + 2,
        `IF(${retBumpAmtRef}="","",${retBumpAmtRef}-${curRetRef})`,
        NUM_FMT);
    }

    writeCell(row, catCol, t.category);
  }

  // ── Sum totals row (row just before data, or we can add at bottom) ──
  // Add SUM formulas for key columns in a totals area
  // Using row 2 sub-header area to add sum validation (like the source file)
  // We'll skip this to keep it clean - the source puts sums in row 3 which is our data start

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
