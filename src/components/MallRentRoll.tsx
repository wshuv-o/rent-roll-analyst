// src/components/MallRentRoll.tsx — Final RR display + 2-sheet Excel export
import { useMemo } from 'react';
import type { MallRentRollTenant } from '@/lib/rent-roll-types/mall-rent-roll-parser';
import { DEFAULT_CHARGE_CODE_MAPPING, CHARGE_CODES } from '@/lib/rent-roll-types/mall-rent-roll-parser';
import * as XLSX from 'xlsx';

type Cell = string | number | Date | null;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function fmt(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
  if (typeof v === 'number') {
    if (v > 20000 && v < 60000) {
      const d = excelDateToJS(v);
      if (d) return d.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    }
    if (Math.abs(v) >= 1) return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (v !== 0) return v.toLocaleString('en-US', { maximumFractionDigits: 4 });
    return '0.00';
  }
  return String(v).trim();
}

function excelDateToJS(serial: number): Date | null {
  if (serial < 1) return null;
  const d = new Date(1899, 11, 30);
  d.setDate(d.getDate() + serial);
  return isNaN(d.getTime()) ? null : d;
}

function colToRef(col: number, row: number): string {
  let letter = '';
  let n = col;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return `${letter}${row}`;
}

function isEmpty(v: Cell): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  if (typeof v === 'number' && v === 0) return true;
  return false;
}

// ─── Build flat row data ─────────────────────────────────────────────────────

// Compute category totals from charge-by-code using the mapping
const MAPPING_BY_CODE = new Map(DEFAULT_CHARGE_CODE_MAPPING.map(m => [m.code, m]));

function computeCategoryTotal(byCode: Record<string, number>, category: string): number {
  let total = 0;
  for (const [code, amt] of Object.entries(byCode)) {
    const m = MAPPING_BY_CODE.get(code);
    if (m && m.category === category) total += amt;
  }
  return total;
}

// ─── Column definitions for display ──────────────────────────────────────────

interface ColDef {
  key: string;
  label: string;
  group: string;
  right?: boolean;
  getter: (t: MallRentRollTenant) => Cell;
}

function buildColumns(): { cols: ColDef[]; groups: { name: string; count: number; color: string }[] } {
  const cols: ColDef[] = [];
  const groups: { name: string; count: number; color: string }[] = [];

  // 1. Identity
  const identityCols: ColDef[] = [
    { key: 'unit', label: 'Units', group: 'Identity', getter: t => t.unit },
    { key: 'dba', label: 'Tenant Name', group: 'Identity', getter: t => t.dba },
    { key: 'leaseId', label: 'Lease ID', group: 'Identity', getter: t => t.leaseId },
    { key: 'sf', label: 'SF', group: 'Identity', right: true, getter: t => t.squareFootage },
    { key: 'leaseType', label: 'Lease Type', group: 'Identity', getter: t => t.leaseType },
    { key: 'spaceType', label: 'Space Type', group: 'Identity', getter: t => t.category },
    { key: 'unitType', label: 'Unit Type', group: 'Identity', getter: t => t.unitType },
    { key: 'leaseStatus', label: 'Lease Status', group: 'Identity', getter: t => t.leaseStatus },
    { key: 'pil', label: 'PIL', group: 'Identity', getter: t => t.percentInLieu },
  ];
  cols.push(...identityCols);
  groups.push({ name: 'Identity', count: identityCols.length, color: 'bg-primary/10 text-primary' });

  // 2. Current charge
  const chargeCols: ColDef[] = [
    { key: 'chargeCode', label: 'Code', group: 'Charge', getter: t => t.charges[0]?.billCode ?? null },
    { key: 'chargeDesc', label: 'Expense Description', group: 'Charge', getter: t => t.charges[0]?.expenseDescription ?? null },
    { key: 'commence', label: 'Commencement Date', group: 'Charge', getter: t => t.commencementDate },
    { key: 'origEnd', label: 'Original End Date', group: 'Charge', getter: t => t.originalEndDate },
    { key: 'expire', label: 'Expire/Close Date', group: 'Charge', getter: t => t.expireCloseDate },
  ];
  cols.push(...chargeCols);
  groups.push({ name: 'Current Charge', count: chargeCols.length, color: 'bg-emerald-500/10 text-emerald-400' });

  // 3. Annual category totals
  const catCols: ColDef[] = [
    { key: 'annRent', label: 'Rent', group: 'Annual', right: true, getter: t => computeCategoryTotal(t.annualChargesByCode, 'Rent') },
    { key: 'rentPSF', label: 'Rent PSF', group: 'Annual', right: true, getter: t => { const r = computeCategoryTotal(t.annualChargesByCode, 'Rent'); const s = t.squareFootage; return s ? r / s : null; } },
    { key: 'annCAM', label: 'CAM', group: 'Annual', right: true, getter: t => computeCategoryTotal(t.annualChargesByCode, 'CAM') },
    { key: 'annRET', label: 'RET', group: 'Annual', right: true, getter: t => computeCategoryTotal(t.annualChargesByCode, 'RET') },
    { key: 'annUTL', label: 'UTL', group: 'Annual', right: true, getter: t => computeCategoryTotal(t.annualChargesByCode, 'UTL') },
    { key: 'annRelief', label: 'Relief', group: 'Annual', right: true, getter: t => computeCategoryTotal(t.annualChargesByCode, 'Relief') },
    { key: 'annExcluded', label: 'Excluded', group: 'Annual', right: true, getter: t => computeCategoryTotal(t.annualChargesByCode, 'Excluded') },
  ];
  cols.push(...catCols);
  groups.push({ name: 'Annual Totals', count: catCols.length, color: 'bg-violet-500/10 text-violet-400' });

  // 4. Individual charge codes
  const codeCols: ColDef[] = CHARGE_CODES.map(code => ({
    key: `code_${code}`, label: code, group: 'Codes', right: true,
    getter: (t: MallRentRollTenant) => t.annualChargesByCode[code] ?? 0,
  }));
  cols.push(...codeCols);
  groups.push({ name: 'Charge Codes', count: codeCols.length, color: 'bg-slate-500/10 text-slate-400' });

  // 5. Total & Variance
  const tvCols: ColDef[] = [
    { key: 'total', label: 'Total', group: 'Totals', right: true, getter: t => t.annualTotal },
    { key: 'variance', label: 'Variance', group: 'Totals', right: true, getter: t => t.variance },
  ];
  cols.push(...tvCols);
  groups.push({ name: 'Totals', count: tvCols.length, color: 'bg-slate-500/10 text-slate-400' });

  // 6. Rent bumps (18 pairs)
  const rbCols: ColDef[] = [];
  for (let i = 0; i < 18; i++) {
    rbCols.push(
      { key: `rb_d${i}`, label: `Bump Date ${i + 1}`, group: 'RentBumps', getter: t => t.rentBumps[i]?.date ?? null },
      { key: `rb_a${i}`, label: `Bump Rent ${i + 1}`, group: 'RentBumps', right: true, getter: t => t.rentBumps[i]?.amount ?? null },
    );
  }
  cols.push(...rbCols);
  groups.push({ name: 'Rent Bumps', count: rbCols.length, color: 'bg-orange-500/10 text-orange-400' });

  // 7. Breakpoints (current + 5 future)
  const bpLabels = ['Current', 'BP 1', 'BP 2', 'BP 3', 'BP 4', 'BP 5'];
  const bpCols: ColDef[] = [];
  for (let i = 0; i < 6; i++) {
    bpCols.push(
      { key: `bp_d${i}`, label: `${bpLabels[i]} Date`, group: 'Breakpoints', getter: t => t.breakpoints[i]?.date ?? null },
      { key: `bp_a${i}`, label: `${bpLabels[i]} Amount`, group: 'Breakpoints', right: true, getter: t => t.breakpoints[i]?.amount ?? null },
      { key: `bp_p${i}`, label: `${bpLabels[i]} %`, group: 'Breakpoints', right: true, getter: t => t.breakpoints[i]?.percent ?? null },
    );
  }
  cols.push(...bpCols);
  groups.push({ name: 'Breakpoints', count: bpCols.length, color: 'bg-rose-500/10 text-rose-400' });

  // 8-10. CAM/UTL/RET bumps (12 pairs each)
  for (const [prefix, field, label, clr] of [
    ['cam', 'camBumps', 'CAM Bumps', 'bg-teal-500/10 text-teal-400'],
    ['utl', 'utlBumps', 'UTL Bumps', 'bg-cyan-500/10 text-cyan-400'],
    ['ret', 'retBumps', 'RET Bumps', 'bg-pink-500/10 text-pink-400'],
  ] as const) {
    const bCols: ColDef[] = [];
    for (let i = 0; i < 12; i++) {
      bCols.push(
        { key: `${prefix}_d${i}`, label: `${prefix.toUpperCase()} Bump Date ${i + 1}`, group: label, getter: t => (t[field] as { date: Cell; amount: Cell }[])[i]?.date ?? null },
        { key: `${prefix}_a${i}`, label: `${prefix.toUpperCase()} Bump Amt ${i + 1}`, group: label, right: true, getter: t => (t[field] as { date: Cell; amount: Cell }[])[i]?.amount ?? null },
      );
    }
    cols.push(...bCols);
    groups.push({ name: label, count: bCols.length, color: clr });
  }

  // 11. Category
  cols.push({ key: 'catLabel', label: 'Category', group: 'Category', getter: t => t.category });
  groups.push({ name: 'Category', count: 1, color: 'bg-slate-500/10 text-slate-400' });

  return { cols, groups };
}

// ─── Excel export: 2-sheet workbook ──────────────────────────────────────────

function downloadFinalRR(tenants: MallRentRollTenant[], fileName: string) {
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: Mapping ──
  const mappingData: (string | null)[][] = [
    ['Code', 'Description', 'Category', 'Relief Sub-Type'],
    ...DEFAULT_CHARGE_CODE_MAPPING.map(m => [m.code, m.description, m.category, m.reliefSubType]),
  ];
  const wsMapping = XLSX.utils.aoa_to_sheet(mappingData);
  wsMapping['!cols'] = [{ wch: 8 }, { wch: 25 }, { wch: 12 }, { wch: 15 }];
  XLSX.utils.book_append_sheet(wb, wsMapping, 'Mapping');

  // ── Sheet 2: Final RR ──
  const ws: XLSX.WorkSheet = {};
  const codeCount = CHARGE_CODES.length; // 28

  // Column layout indices for Final RR
  let c = 0;
  const COL = {
    unit: c++, dba: c++, leaseId: c++, sf: c++, leaseType: c++,
    spaceType: c++, unitType: c++, leaseStatus: c++, pil: c++,
    chargeCode: c++, chargeDesc: c++,
    commence: c++, origEnd: c++, expire: c++,
    // Category totals (formulas)
    annRent: c++, rentPSF: c++, annCAM: c++, annRET: c++, annUTL: c++, annRelief: c++, annExcluded: c++,
  };
  // Individual charge codes start here
  const codeStart = c;
  c += codeCount; // 28 columns for individual codes
  const COL2 = {
    total: c++, variance: c++,
  };
  // Rent bumps (18 pairs)
  const rentBumpStart = c;
  c += 36;
  // Breakpoints (current + 5 future, each date/amount/%)
  const bpStart = c;
  c += 18;
  // CAM bumps (12 pairs)
  const camBumpStart = c;
  c += 24;
  // UTL bumps (12 pairs)
  const utlBumpStart = c;
  c += 24;
  // RET bumps (12 pairs)
  const retBumpStart = c;
  c += 24;
  // Category label
  const catLabelCol = c++;
  const totalCols = c;

  // Row 1: Headers
  const headers: string[] = new Array(totalCols).fill('');
  headers[COL.unit] = 'Units'; headers[COL.dba] = 'Tenant Name'; headers[COL.leaseId] = 'Lease ID';
  headers[COL.sf] = 'SF'; headers[COL.leaseType] = 'Lease Type';
  headers[COL.spaceType] = 'Space Type'; headers[COL.unitType] = 'Unit Type';
  headers[COL.leaseStatus] = 'Lease Status'; headers[COL.pil] = 'PIL';
  headers[COL.chargeCode] = 'Code'; headers[COL.chargeDesc] = 'Expense Description';
  headers[COL.commence] = 'Commencement Date'; headers[COL.origEnd] = 'Original End Date';
  headers[COL.expire] = 'Expire/Close Date';
  headers[COL.annRent] = 'Rent'; headers[COL.rentPSF] = 'Rent PSF';
  headers[COL.annCAM] = 'CAM'; headers[COL.annRET] = 'RET'; headers[COL.annUTL] = 'UTL';
  headers[COL.annRelief] = 'Relief'; headers[COL.annExcluded] = 'Excluded';
  for (let i = 0; i < codeCount; i++) headers[codeStart + i] = CHARGE_CODES[i];
  headers[COL2.total] = 'Total'; headers[COL2.variance] = 'Variance';
  for (let i = 0; i < 18; i++) { headers[rentBumpStart + i * 2] = `Bump Date ${i + 1}`; headers[rentBumpStart + i * 2 + 1] = `Bump Rent ${i + 1}`; }
  const bpLabels = ['Current', 'BP 1', 'BP 2', 'BP 3', 'BP 4', 'BP 5'];
  for (let i = 0; i < 6; i++) { headers[bpStart + i * 3] = `${bpLabels[i]} BP Date`; headers[bpStart + i * 3 + 1] = `${bpLabels[i]} Breakpoint`; headers[bpStart + i * 3 + 2] = `${bpLabels[i]} %`; }
  for (let i = 0; i < 12; i++) { headers[camBumpStart + i * 2] = `CAM Bump Date ${i + 1}`; headers[camBumpStart + i * 2 + 1] = `CAM Bump Amount ${i + 1}`; }
  for (let i = 0; i < 12; i++) { headers[utlBumpStart + i * 2] = `UTL Bump Date ${i + 1}`; headers[utlBumpStart + i * 2 + 1] = `UTL Bump Amount ${i + 1}`; }
  for (let i = 0; i < 12; i++) { headers[retBumpStart + i * 2] = `RET Bump Date ${i + 1}`; headers[retBumpStart + i * 2 + 1] = `RET Bump Amount ${i + 1}`; }
  headers[catLabelCol] = 'Category';

  // Write header row (row 1, 0-indexed r=0)
  for (let ci = 0; ci < headers.length; ci++) {
    if (headers[ci]) ws[XLSX.utils.encode_cell({ r: 0, c: ci })] = { v: headers[ci], t: 's' };
  }

  // Write data rows (starting at r=1)
  // Mapping range for SUMPRODUCT: Mapping!$C$2:$C$29
  const mappingCatRange = `Mapping!$C$2:$C$${1 + codeCount}`;
  const codeStartRef = (row: number) => colToRef(codeStart, row);
  const codeEndRef = (row: number) => colToRef(codeStart + codeCount - 1, row);

  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const r = ti + 1; // 0-indexed row (row 0 = header)
    const excelRow = r + 1; // 1-indexed for formulas

    const writeCell = (col: number, v: Cell, type?: string) => {
      if (v === null || v === undefined) return;
      if (v instanceof Date) { ws[XLSX.utils.encode_cell({ r, c: col })] = { v, t: 'd' }; return; }
      if (typeof v === 'number') {
        // Convert Excel serial dates for date columns
        ws[XLSX.utils.encode_cell({ r, c: col })] = { v, t: type || 'n' };
        return;
      }
      ws[XLSX.utils.encode_cell({ r, c: col })] = { v: String(v), t: 's' };
    };

    // Identity
    writeCell(COL.unit, t.unit); writeCell(COL.dba, t.dba); writeCell(COL.leaseId, t.leaseId);
    writeCell(COL.sf, t.squareFootage); writeCell(COL.leaseType, t.leaseType);
    writeCell(COL.spaceType, t.category); writeCell(COL.unitType, t.unitType);
    writeCell(COL.leaseStatus, t.leaseStatus); writeCell(COL.pil, t.percentInLieu);

    // Current charge
    if (t.charges.length > 0) {
      writeCell(COL.chargeCode, t.charges[0].billCode);
      writeCell(COL.chargeDesc, t.charges[0].expenseDescription);
    }
    writeCell(COL.commence, t.commencementDate);
    writeCell(COL.origEnd, t.originalEndDate);
    writeCell(COL.expire, t.expireCloseDate);

    // Category totals — SUMPRODUCT formulas referencing Mapping sheet
    const cRange = `${codeStartRef(excelRow)}:${codeEndRef(excelRow)}`;
    ws[XLSX.utils.encode_cell({ r, c: COL.annRent })] =     { f: `SUMPRODUCT((${mappingCatRange}="Rent")*(${cRange}))`, t: 'n' };
    ws[XLSX.utils.encode_cell({ r, c: COL.rentPSF })] =     { f: `IF(${colToRef(COL.sf, excelRow)}=0,"",${colToRef(COL.annRent, excelRow)}/${colToRef(COL.sf, excelRow)})`, t: 'n' };
    ws[XLSX.utils.encode_cell({ r, c: COL.annCAM })] =      { f: `SUMPRODUCT((${mappingCatRange}="CAM")*(${cRange}))`, t: 'n' };
    ws[XLSX.utils.encode_cell({ r, c: COL.annRET })] =      { f: `SUMPRODUCT((${mappingCatRange}="RET")*(${cRange}))`, t: 'n' };
    ws[XLSX.utils.encode_cell({ r, c: COL.annUTL })] =      { f: `SUMPRODUCT((${mappingCatRange}="UTL")*(${cRange}))`, t: 'n' };
    ws[XLSX.utils.encode_cell({ r, c: COL.annRelief })] =   { f: `SUMPRODUCT((${mappingCatRange}="Relief")*(${cRange}))`, t: 'n' };
    ws[XLSX.utils.encode_cell({ r, c: COL.annExcluded })] = { f: `SUMPRODUCT((${mappingCatRange}="Excluded")*(${cRange}))`, t: 'n' };

    // Individual charge codes
    for (let i = 0; i < CHARGE_CODES.length; i++) {
      writeCell(codeStart + i, t.annualChargesByCode[CHARGE_CODES[i]] ?? 0);
    }

    // Total & Variance
    writeCell(COL2.total, t.annualTotal);
    writeCell(COL2.variance, t.variance);

    // Rent bumps (18 pairs)
    for (let i = 0; i < 18; i++) {
      if (t.rentBumps[i]) {
        writeCell(rentBumpStart + i * 2, t.rentBumps[i].date);
        writeCell(rentBumpStart + i * 2 + 1, t.rentBumps[i].amount);
      }
    }

    // Breakpoints (6 entries: current + 5 future)
    for (let i = 0; i < 6 && i < t.breakpoints.length; i++) {
      writeCell(bpStart + i * 3, t.breakpoints[i].date);
      writeCell(bpStart + i * 3 + 1, t.breakpoints[i].amount);
      writeCell(bpStart + i * 3 + 2, t.breakpoints[i].percent);
    }

    // CAM/UTL/RET bumps (12 pairs each)
    for (let i = 0; i < 12; i++) {
      if (t.camBumps[i]) { writeCell(camBumpStart + i * 2, t.camBumps[i].date); writeCell(camBumpStart + i * 2 + 1, t.camBumps[i].amount); }
      if (t.utlBumps[i]) { writeCell(utlBumpStart + i * 2, t.utlBumps[i].date); writeCell(utlBumpStart + i * 2 + 1, t.utlBumps[i].amount); }
      if (t.retBumps[i]) { writeCell(retBumpStart + i * 2, t.retBumps[i].date); writeCell(retBumpStart + i * 2 + 1, t.retBumps[i].amount); }
    }

    // Category label
    writeCell(catLabelCol, t.category);
  }

  // Set range
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: tenants.length, c: totalCols - 1 } });

  // Column widths
  const colWidths: XLSX.ColInfo[] = [];
  for (let ci = 0; ci < totalCols; ci++) {
    if (ci === COL.dba || ci === COL.chargeDesc) colWidths.push({ wch: 28 });
    else if (ci === COL.unit || ci === COL.leaseId) colWidths.push({ wch: 12 });
    else colWidths.push({ wch: 14 });
  }
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, 'Final RR');

  const outName = fileName.replace(/\.[^.]+$/, '') + '_Final_RR.xlsx';
  XLSX.writeFile(wb, outName);
}

// ─── Component ───────────────────────────────────────────────────────────────

interface Props {
  tenants: MallRentRollTenant[];
  fileName: string;
  onBack: () => void;
}

export function MallRentRollTable({ tenants, fileName, onBack }: Props) {
  const { cols, groups } = useMemo(() => buildColumns(), []);

  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="shrink-0 flex items-center justify-between px-4 py-2 border-b border-panel-border bg-background">
        <div className="flex items-center gap-3">
          <button onClick={onBack} className="text-[11px] font-mono text-muted-foreground hover:text-foreground transition-colors">&larr; Back</button>
          <span className="text-[11px] font-mono text-foreground">{tenants.length} tenant{tenants.length !== 1 ? 's' : ''}</span>
        </div>
        <button
          onClick={() => downloadFinalRR(tenants, fileName)}
          className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors"
        >
          &darr; Download Final RR
        </button>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="text-[11px] font-mono border-collapse w-full">
          <thead className="sticky top-0 z-10">
            {/* Group header */}
            <tr>
              {groups.map(g => (
                <th key={g.name} colSpan={g.count} className={`px-2 py-1 text-left border border-panel-border font-medium tracking-wide ${g.color}`}>
                  {g.name}
                </th>
              ))}
            </tr>
            {/* Column labels */}
            <tr>
              {cols.map(col => (
                <th key={col.key} className={`px-2 py-1 border border-panel-border whitespace-nowrap font-medium text-muted-foreground bg-background ${col.right ? 'text-right' : 'text-left'}`}>
                  {col.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tenants.map((t, ri) => (
              <tr key={ri} className={[
                'hover:bg-muted/30 transition-colors',
                ri > 0 && t.category !== tenants[ri - 1]?.category ? 'border-t-2 border-t-primary/20' : '',
              ].join(' ')}>
                {cols.map(col => {
                  const raw = col.getter(t);
                  const display = col.right && typeof raw === 'number' ? fmt(raw) : (isEmpty(raw) ? '' : fmt(raw));
                  return (
                    <td key={col.key} className={[
                      'px-2 py-1 border border-panel-border whitespace-nowrap',
                      col.right ? 'text-right tabular-nums' : '',
                      !display ? 'text-muted-foreground/30' : 'text-foreground',
                    ].join(' ')}>
                      {display || '\u2014'}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
