// src/components/MallRentRoll.tsx — Mall Rent Roll display + Excel exports
import { useMemo } from 'react';
import type { MallRentRollTenant } from '@/lib/rent-roll-types/mall-rent-roll-parser';
import { DEFAULT_CHARGE_CODE_MAPPING } from '@/lib/rent-roll-types/mall-rent-roll-parser';
import { downloadSemiFinalRR } from '@/lib/semi-final-export';
import { downloadFinalRR } from '@/lib/final-rr-export';

type Cell = string | number | Date | null;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function fmt(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    const m = v.getMonth() + 1;
    const d = v.getDate();
    const y = v.getFullYear();
    return `${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}/${y}`;
  }
  if (typeof v === 'number') {
    if (Math.abs(v) >= 1) return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (v !== 0) return v.toLocaleString('en-US', { maximumFractionDigits: 4 });
    return '0.00';
  }
  return String(v).trim();
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

// ─── Build annual charges from tenant.charges[] ─────────────────────────────

function buildAnnualByCode(t: MallRentRollTenant): Record<string, number> {
  const byCode: Record<string, number> = {};
  for (const ch of t.charges) {
    if (!ch.billCode) continue;
    const monthly = ch.monthlyAmount ?? 0;
    byCode[ch.billCode] = (byCode[ch.billCode] || 0) + monthly * 12;
  }
  return byCode;
}

// ─── Gather all unique charge codes from all tenants ────────────────────────

function gatherAllChargeCodes(tenants: MallRentRollTenant[]): string[] {
  const codeSet = new Set<string>();
  for (const t of tenants) {
    for (const ch of t.charges) {
      if (ch.billCode) codeSet.add(ch.billCode);
    }
  }
  const knownOrder = DEFAULT_CHARGE_CODE_MAPPING.map(m => m.code);
  const known = knownOrder.filter(c => codeSet.has(c));
  const unknown = Array.from(codeSet).filter(c => !knownOrder.includes(c)).sort();
  return [...known, ...unknown];
}


// ─── Compute category totals from charges ───────────────────────────────────

const MAPPING_BY_CODE = new Map(DEFAULT_CHARGE_CODE_MAPPING.map(m => [m.code, m]));

function computeCategoryTotal(annualByCode: Record<string, number>, category: string): number {
  let total = 0;
  for (const [code, amt] of Object.entries(annualByCode)) {
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
  getter: (t: MallRentRollTenant, annualByCode: Record<string, number>) => Cell;
}

function buildColumns(allCodes: string[]): { cols: ColDef[]; groups: { name: string; count: number; color: string }[] } {
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
    { key: 'annRent', label: 'Rent', group: 'Annual', right: true, getter: (t, abc) => computeCategoryTotal(abc, 'Rent') },
    { key: 'rentPSF', label: 'Rent PSF', group: 'Annual', right: true, getter: (t, abc) => { const r = computeCategoryTotal(abc, 'Rent'); const s = t.squareFootage; return s ? r / s : null; } },
    { key: 'annCAM', label: 'CAM', group: 'Annual', right: true, getter: (t, abc) => computeCategoryTotal(abc, 'CAM') },
    { key: 'annRET', label: 'RET', group: 'Annual', right: true, getter: (t, abc) => computeCategoryTotal(abc, 'RET') },
    { key: 'annUTL', label: 'UTL', group: 'Annual', right: true, getter: (t, abc) => computeCategoryTotal(abc, 'UTL') },
    { key: 'annRelief', label: 'Relief', group: 'Annual', right: true, getter: (t, abc) => computeCategoryTotal(abc, 'Relief') },
    { key: 'annExcluded', label: 'Excluded', group: 'Annual', right: true, getter: (t, abc) => computeCategoryTotal(abc, 'Excluded') },
  ];
  cols.push(...catCols);
  groups.push({ name: 'Annual Totals', count: catCols.length, color: 'bg-violet-500/10 text-violet-400' });

  // 4. Individual charge codes (dynamic)
  const codeCols: ColDef[] = allCodes.map(code => ({
    key: `code_${code}`, label: code, group: 'Codes', right: true,
    getter: (_t: MallRentRollTenant, abc: Record<string, number>) => abc[code] ?? 0,
  }));
  cols.push(...codeCols);
  groups.push({ name: 'Charge Codes', count: codeCols.length, color: 'bg-slate-500/10 text-slate-400' });

  // 5. Total & Variance
  const tvCols: ColDef[] = [
    { key: 'total', label: 'Total', group: 'Totals', right: true, getter: (t, abc) => Object.values(abc).reduce((s, v) => s + v, 0) },
    { key: 'variance', label: 'Variance', group: 'Totals', right: true, getter: t => t.variance },
  ];
  cols.push(...tvCols);
  groups.push({ name: 'Totals', count: tvCols.length, color: 'bg-slate-500/10 text-slate-400' });

  // 6. Rent bump summary + 18 pairs
  const rbSumCols: ColDef[] = [
    { key: 'rb_sum_date', label: 'Bump Date', group: 'RentBumps', getter: t => t.rentBumps[0]?.date ?? null },
    { key: 'rb_sum_amt', label: 'Amount', group: 'RentBumps', right: true, getter: t => t.rentBumps[0]?.amount ?? null },
    { key: 'rb_sum_uw', label: 'UW Rent', group: 'RentBumps', right: true, getter: t => {
      const bump = t.rentBumps[0];
      if (!bump?.amount || typeof bump.amount !== 'number') return null;
      return bump.amount * (t.squareFootage ?? 0);
    }},
  ];
  const rbCols: ColDef[] = [];
  for (let i = 0; i < 18; i++) {
    rbCols.push(
      { key: `rb_d${i}`, label: `Bump Date ${i + 1}`, group: 'RentBumps', getter: t => t.rentBumps[i]?.date ?? null },
      { key: `rb_a${i}`, label: `Bump Rent ${i + 1}`, group: 'RentBumps', right: true, getter: t => t.rentBumps[i]?.amount ?? null },
    );
  }
  cols.push(...rbSumCols, ...rbCols);
  groups.push({ name: 'Rent Bumps', count: rbSumCols.length + rbCols.length, color: 'bg-orange-500/10 text-orange-400' });

  // 7. Breakpoints (current + 5 future)
  const bpLabels = ['Current', 'BP 1', 'BP 2', 'BP 3', 'BP 4', 'BP 5'];
  const bpCols: ColDef[] = [];
  for (let i = 0; i < 6; i++) {
    bpCols.push(
      { key: `bp_d${i}`, label: `${bpLabels[i]} BP Date`, group: 'Breakpoints', getter: t => t.breakpoints[i]?.date ?? null },
      { key: `bp_a${i}`, label: `${bpLabels[i]} Breakpoint`, group: 'Breakpoints', right: true, getter: t => t.breakpoints[i]?.amount ?? null },
      { key: `bp_p${i}`, label: `${bpLabels[i]} %`, group: 'Breakpoints', right: true, getter: t => t.breakpoints[i]?.percent ?? null },
    );
  }
  cols.push(...bpCols);
  groups.push({ name: 'Breakpoints', count: bpCols.length, color: 'bg-rose-500/10 text-rose-400' });

  // 8-10. CAM/UTL/RET bumps with summaries (12 pairs each)
  for (const [prefix, field, label, clr] of [
    ['cam', 'camBumps', 'CAM Bumps', 'bg-teal-500/10 text-teal-400'],
    ['utl', 'utlBumps', 'UTL Bumps', 'bg-cyan-500/10 text-cyan-400'],
    ['ret', 'retBumps', 'RET Bumps', 'bg-pink-500/10 text-pink-400'],
  ] as const) {
    const sumCols: ColDef[] = [
      { key: `${prefix}_sum_date`, label: 'Bump Date', group: label, getter: t => (t[field] as { date: Cell; amount: Cell }[])[0]?.date ?? null },
      { key: `${prefix}_sum_amt`, label: 'Amount', group: label, right: true, getter: t => (t[field] as { date: Cell; amount: Cell }[])[0]?.amount ?? null },
      { key: `${prefix}_sum_chg`, label: `Changes on ${prefix.toUpperCase()}`, group: label, right: true, getter: (t, abc) => {
        const bumps = t[field] as { date: Cell; amount: Cell }[];
        if (!bumps[0]?.amount || typeof bumps[0].amount !== 'number') return null;
        const currentAnnual = computeCategoryTotal(abc, prefix === 'cam' ? 'CAM' : prefix === 'utl' ? 'UTL' : 'RET');
        return bumps[0].amount - currentAnnual;
      }},
    ];
    const bCols: ColDef[] = [];
    for (let i = 0; i < 12; i++) {
      bCols.push(
        { key: `${prefix}_d${i}`, label: `${prefix.toUpperCase()} Bump Date ${i + 1}`, group: label, getter: t => (t[field] as { date: Cell; amount: Cell }[])[i]?.date ?? null },
        { key: `${prefix}_a${i}`, label: `${prefix.toUpperCase()} Bump Amt ${i + 1}`, group: label, right: true, getter: t => (t[field] as { date: Cell; amount: Cell }[])[i]?.amount ?? null },
      );
    }
    cols.push(...sumCols, ...bCols);
    groups.push({ name: label, count: sumCols.length + bCols.length, color: clr });
  }

  // 11. Category
  cols.push({ key: 'catLabel', label: 'Category', group: 'Category', getter: t => t.category });
  groups.push({ name: 'Category', count: 1, color: 'bg-slate-500/10 text-slate-400' });

  return { cols, groups };
}

// downloadFinalRR is now in src/lib/final-rr-export.ts

// ─── Component ───────────────────────────────────────────────────────────────

interface Props {
  tenants: MallRentRollTenant[];
  fileName: string;
  onBack: () => void;
}

export function MallRentRollTable({ tenants, fileName, onBack }: Props) {
  const allCodes = useMemo(() => gatherAllChargeCodes(tenants), [tenants]);
  const { cols, groups } = useMemo(() => buildColumns(allCodes), [allCodes]);
  const annualByCodeMap = useMemo(() => {
    const map = new Map<MallRentRollTenant, Record<string, number>>();
    for (const t of tenants) map.set(t, buildAnnualByCode(t));
    return map;
  }, [tenants]);

  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="shrink-0 flex items-center justify-between px-4 py-2 border-b border-panel-border bg-background">
        <div className="flex items-center gap-3">
          <button onClick={onBack} className="text-[11px] font-mono text-muted-foreground hover:text-foreground transition-colors">&larr; Back</button>
          <span className="text-[11px] font-mono text-foreground">{tenants.length} tenant{tenants.length !== 1 ? 's' : ''}</span>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => downloadSemiFinalRR(tenants, fileName)}
            className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors"
          >
            &darr; Semi Final Download
          </button>
          <button
            onClick={() => downloadFinalRR(tenants, fileName)}
            className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors"
          >
            &darr; Download Final RR
          </button>
        </div>
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
            {tenants.map((t, ri) => {
              const abc = annualByCodeMap.get(t) || {};
              return (
                <tr key={ri} className={[
                  'hover:bg-muted/30 transition-colors',
                  ri > 0 && t.category !== tenants[ri - 1]?.category ? 'border-t-2 border-t-primary/20' : '',
                ].join(' ')}>
                  {cols.map(col => {
                    const raw = col.getter(t, abc);
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
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
