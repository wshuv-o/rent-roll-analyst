// src/components/MallRentRoll.tsx
import { useMemo } from 'react';
import type { MallRentRollTenant } from '@/lib/rent-roll-types/mall-rent-roll-parser';
import * as XLSX from 'xlsx';

// ─── Types ───────────────────────────────────────────────────────────────────

type Cell = string | number | Date | null;

interface FlatRow {
  _tenantIdx: number;
  unit: string;
  dba: string;
  leaseId: string;
  squareFootage: number | null;
  category: string;
  leaseType: string | null;
  unitType: string | null;
  leaseStatus: string | null;
  commencementDate: Cell;
  originalEndDate: Cell;
  expireCloseDate: Cell;
  // Charge summary
  totalMonthlyAmount: number | null;
  annualTotal: number | null;
  // Per-code charges (dynamic keys set later)
  chargeBreakdown: Record<string, number>;
  // Charge detail string
  chargeDetails: string;
  // Future escalation summary
  futureDetails: string;
  futureMonthlyTotal: number | null;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function fmt(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
  if (typeof v === 'number') {
    // Check if this looks like an Excel serial date (range 1-60000)
    if (v > 20000 && v < 60000) {
      const d = excelDateToJS(v);
      if (d) return d.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    }
    if (Math.abs(v) >= 1000) return v.toLocaleString('en-US', { maximumFractionDigits: 2 });
    return String(v);
  }
  return String(v).trim();
}

function fmtMoney(v: number | null): string {
  if (v === null) return '';
  return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function excelDateToJS(serial: number): Date | null {
  if (serial < 1) return null;
  // Excel epoch: Jan 1, 1900 (with the Lotus 1-2-3 bug)
  const epoch = new Date(1899, 11, 30);
  const d = new Date(epoch.getTime() + serial * 86400000);
  return isNaN(d.getTime()) ? null : d;
}

// ─── Flatten ─────────────────────────────────────────────────────────────────

function flatten(tenants: MallRentRollTenant[]): { rows: FlatRow[]; allCodes: string[] } {
  // Collect all unique charge codes across tenants
  const codeSet = new Set<string>();
  for (const t of tenants) {
    for (const code of Object.keys(t.annualChargesByCode)) {
      codeSet.add(code);
    }
  }
  const allCodes = [...codeSet].sort();

  const rows: FlatRow[] = tenants.map((t, idx) => {
    // Build charge detail string
    const chargeDetails = t.charges
      .map(c => `${c.billCode}: ${c.expenseDescription} $${fmtMoney(c.monthlyAmount)}/mo`)
      .join('; ');

    // Build future escalation summary
    const futureDetails = t.futureEscalations
      .map(f => `${f.billCode}: ${f.expenseDescription} $${fmtMoney(f.monthlyAmount)}/mo`)
      .join('; ');

    const futureMonthlyTotal = t.futureEscalations.reduce(
      (sum, f) => sum + (f.monthlyAmount ?? 0), 0
    ) || null;

    return {
      _tenantIdx: idx,
      unit: t.unit,
      dba: t.dba,
      leaseId: t.leaseId,
      squareFootage: t.squareFootage,
      category: t.category,
      leaseType: t.leaseType,
      unitType: t.unitType,
      leaseStatus: t.leaseStatus,
      commencementDate: t.commencementDate,
      originalEndDate: t.originalEndDate,
      expireCloseDate: t.expireCloseDate,
      totalMonthlyAmount: t.totalMonthlyAmount,
      annualTotal: t.annualTotal,
      chargeBreakdown: t.annualChargesByCode,
      chargeDetails,
      futureDetails,
      futureMonthlyTotal,
    };
  });

  return { rows, allCodes };
}

// ─── Column definitions ──────────────────────────────────────────────────────

interface ColDef {
  key: string;
  label: string;
  group: 'tenant' | 'charges' | 'future';
  right?: boolean;
  getter: (row: FlatRow) => Cell;
}

function buildColumns(allCodes: string[]): ColDef[] {
  const cols: ColDef[] = [
    { key: 'unit', label: 'Unit', group: 'tenant', getter: r => r.unit },
    { key: 'dba', label: 'Tenant (DBA)', group: 'tenant', getter: r => r.dba },
    { key: 'leaseId', label: 'Lease ID', group: 'tenant', getter: r => r.leaseId },
    { key: 'sqft', label: 'Sq Ft', group: 'tenant', right: true, getter: r => r.squareFootage },
    { key: 'category', label: 'Category', group: 'tenant', getter: r => r.category },
    { key: 'leaseType', label: 'Lease Type', group: 'tenant', getter: r => r.leaseType },
    { key: 'unitType', label: 'Unit Type', group: 'tenant', getter: r => r.unitType },
    { key: 'status', label: 'Status', group: 'tenant', getter: r => r.leaseStatus },
    { key: 'commence', label: 'Commencement', group: 'tenant', getter: r => r.commencementDate },
    { key: 'origEnd', label: 'Original End', group: 'tenant', getter: r => r.originalEndDate },
    { key: 'expire', label: 'Expire/Close', group: 'tenant', getter: r => r.expireCloseDate },
    // Current charges
    { key: 'totalMonthly', label: 'Monthly Total', group: 'charges', right: true, getter: r => r.totalMonthlyAmount },
    { key: 'annualTotal', label: 'Annual Total', group: 'charges', right: true, getter: r => r.annualTotal },
    { key: 'chargeDetails', label: 'Charge Details', group: 'charges', getter: r => r.chargeDetails },
  ];

  // Add per-code annual charge columns
  for (const code of allCodes) {
    cols.push({
      key: `code_${code}`,
      label: code,
      group: 'charges',
      right: true,
      getter: r => r.chargeBreakdown[code] ?? null,
    });
  }

  // Future escalations
  cols.push(
    { key: 'futureMonthly', label: 'Future Monthly', group: 'future', right: true, getter: r => r.futureMonthlyTotal },
    { key: 'futureDetails', label: 'Future Details', group: 'future', getter: r => r.futureDetails },
  );

  return cols;
}

// ─── Excel export ────────────────────────────────────────────────────────────

function downloadXLSX(rows: FlatRow[], cols: ColDef[], fileName: string) {
  const header = cols.map(c => c.label);
  const wsData: (string | number | Date | null)[][] = [header];

  for (const row of rows) {
    wsData.push(
      cols.map(col => {
        const v = col.getter(row);
        if (v instanceof Date) return v;
        if (typeof v === 'number') {
          // Convert Excel serial dates to JS Dates for date columns
          if (['commence', 'origEnd', 'expire'].includes(col.key) && v > 20000 && v < 60000) {
            return excelDateToJS(v) ?? v;
          }
          return v;
        }
        return v as string | null;
      })
    );
  }

  const ws = XLSX.utils.aoa_to_sheet(wsData, { cellDates: true });
  ws['!cols'] = cols.map(col =>
    col.right ? { wch: 14 } : col.key === 'dba' ? { wch: 30 } : col.key.includes('Details') ? { wch: 50 } : { wch: 16 }
  );
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Mall Rent Roll');

  const outName = fileName.replace(/\.[^.]+$/, '') + '_extracted.xlsx';
  XLSX.writeFile(wb, outName);
}

// ─── Component ───────────────────────────────────────────────────────────────

interface Props {
  tenants: MallRentRollTenant[];
  fileName: string;
  onBack: () => void;
}

export function MallRentRollTable({ tenants, fileName, onBack }: Props) {
  const { rows, allCodes } = useMemo(() => flatten(tenants), [tenants]);
  const cols = useMemo(() => buildColumns(allCodes), [allCodes]);

  const tenantColCount = cols.filter(c => c.group === 'tenant').length;
  const chargesColCount = cols.filter(c => c.group === 'charges').length;
  const futureColCount = cols.filter(c => c.group === 'future').length;

  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="shrink-0 flex items-center justify-between px-4 py-2 border-b border-panel-border bg-background">
        <div className="flex items-center gap-3">
          <button
            onClick={onBack}
            className="text-[11px] font-mono text-muted-foreground hover:text-foreground transition-colors flex items-center gap-1"
          >
            &larr; Back
          </button>
          <span className="text-[11px] font-mono text-foreground">
            {tenants.length} tenant{tenants.length !== 1 ? 's' : ''}
          </span>
        </div>
        <button
          onClick={() => downloadXLSX(rows, cols, fileName)}
          className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors flex items-center gap-1.5"
        >
          &darr; Download Excel
        </button>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="text-[11px] font-mono border-collapse w-full">
          <thead className="sticky top-0 z-10">
            {/* Group header */}
            <tr>
              <th
                colSpan={tenantColCount}
                className="px-2 py-1 text-left border border-panel-border bg-primary/10 text-primary font-medium tracking-wide"
              >
                Tenant
              </th>
              <th
                colSpan={chargesColCount}
                className="px-2 py-1 text-left border border-panel-border bg-emerald-500/10 text-emerald-400 font-medium tracking-wide"
              >
                Current Charges
              </th>
              {futureColCount > 0 && (
                <th
                  colSpan={futureColCount}
                  className="px-2 py-1 text-left border border-panel-border bg-amber-500/10 text-amber-400 font-medium tracking-wide"
                >
                  Future Escalations
                </th>
              )}
            </tr>
            {/* Column labels */}
            <tr>
              {cols.map(col => (
                <th
                  key={col.key}
                  className={[
                    'px-2 py-1 border border-panel-border whitespace-nowrap font-medium',
                    col.group === 'tenant'
                      ? 'bg-primary/5 text-primary'
                      : col.group === 'charges'
                        ? 'bg-emerald-500/5 text-emerald-400'
                        : 'bg-amber-500/5 text-amber-400',
                    col.right ? 'text-right' : 'text-left',
                  ].join(' ')}
                >
                  {col.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, ri) => (
              <tr
                key={ri}
                className={[
                  'hover:bg-muted/30 transition-colors',
                  ri > 0 && row.category !== rows[ri - 1]?.category ? 'border-t-2 border-t-primary/20' : '',
                ].join(' ')}
              >
                {cols.map(col => {
                  const raw = col.getter(row);
                  const display = col.right && typeof raw === 'number'
                    ? fmtMoney(raw)
                    : fmt(raw);

                  return (
                    <td
                      key={col.key}
                      className={[
                        'px-2 py-1 border border-panel-border',
                        col.group === 'tenant' ? 'text-foreground' : 'text-muted-foreground',
                        col.right ? 'text-right tabular-nums whitespace-nowrap' : '',
                        col.key.includes('Details') ? 'max-w-[300px] truncate' : 'whitespace-nowrap',
                        !display ? 'text-muted-foreground/30' : '',
                      ].join(' ')}
                      title={col.key.includes('Details') ? display : undefined}
                    >
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
