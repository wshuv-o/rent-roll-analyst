// src/components/TenancyScheduleTable.tsx
import { useMemo } from 'react';
import type { TenancyScheduleTenant } from '@/lib/rent-roll-types/tenancy-schedule-parser';

// ─── Types ────────────────────────────────────────────────────────────────────

type Cell = string | number | Date | null;

interface FlatRow {
  // identity
  _tenantIdx: number;
  _section: string;
  // tenant main fields
  property: Cell;
  unit: Cell;
  lease: Cell;
  leaseType: Cell;
  area: Cell;
  leaseFrom: Cell;
  leaseTo: Cell;
  term: Cell;
  tenancyYears: Cell;
  monthlyRent: Cell;
  monthlyRentPerArea: Cell;
  annualRent: Cell;
  annualRentPerArea: Cell;
  annualRecPerArea: Cell;
  annualMiscPerArea: Cell;
  securityDepositReceived: Cell;
  locAmount: Cell;
  // sub-section fields (null when row is the main-only summary)
  charge: Cell;
  chargeType: Cell;
  chargeUnit: Cell;
  areaLabel: Cell;
  subArea: Cell;
  from: Cell;
  to: Cell;
  monthlyAmt: Cell;
  amtPerArea: Cell;
  annual: Cell;
  annualPerArea: Cell;
  managementFee: Cell;
  annualGrossAmount: Cell;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function fmt(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' });
  if (typeof v === 'number') {
    // If it looks like a large currency/area figure use comma formatting
    if (Math.abs(v) >= 1000) return v.toLocaleString('en-AU', { maximumFractionDigits: 2 });
    return String(v);
  }
  return String(v).trim();
}

function pick(mr: Record<string, Cell>, ...keys: string[]): Cell {
  for (const k of keys) {
    if (mr[k] !== undefined && mr[k] !== null) return mr[k];
  }
  return null;
}

// ─── Flatten ──────────────────────────────────────────────────────────────────

function flatten(tenants: TenancyScheduleTenant[]): FlatRow[] {
  const rows: FlatRow[] = [];

  for (let ti = 0; ti < tenants.length; ti++) {
    const t = tenants[ti];
    const mr = t.mainRow;

    const tenantBase = {
      _tenantIdx: ti,
      property:              pick(mr, 'Property'),
      unit:                  pick(mr, 'Unit(s)', 'Unit'),
      lease:                 pick(mr, 'Lease'),
      leaseType:             pick(mr, 'Lease Type'),
      area:                  pick(mr, 'Area'),
      leaseFrom:             pick(mr, 'Lease From'),
      leaseTo:               pick(mr, 'Lease To'),
      term:                  pick(mr, 'Term'),
      tenancyYears:          pick(mr, 'Tenancy Years'),
      monthlyRent:           pick(mr, 'Monthly Rent'),
      monthlyRentPerArea:    pick(mr, 'Monthly Rent/Area'),
      annualRent:            pick(mr, 'Annual Rent'),
      annualRentPerArea:     pick(mr, 'Annual Rent/Area'),
      annualRecPerArea:      pick(mr, 'Annual Rec./Area'),
      annualMiscPerArea:     pick(mr, 'Annual Misc/Area'),
      securityDepositReceived: pick(mr, 'Security Deposit Received'),
      locAmount:             pick(mr, 'LOC Amount/ Bank Guarantee'),
    };

    if (t.subSections.length === 0) {
      // Tenant with no sub-sections — emit one summary row
      rows.push({
        ...tenantBase,
        _section: '',
        charge: null, chargeType: null, chargeUnit: null, areaLabel: null,
        subArea: null, from: null, to: null,
        monthlyAmt: null, amtPerArea: null, annual: null, annualPerArea: null,
        managementFee: null, annualGrossAmount: null,
      });
      continue;
    }

    // Emit one row per sub-section data row
    for (const section of t.subSections) {
      for (const dataRow of section.rows) {
        const v = dataRow.values;
        rows.push({
          ...tenantBase,
          _section: section.name,
          charge:             v['Charge'] ?? null,
          chargeType:         v['Type'] ?? null,
          chargeUnit:         v['Unit'] ?? null,
          areaLabel:          v['Area Label'] ?? null,
          subArea:            v['Area'] ?? null,
          from:               v['From'] ?? null,
          to:                 v['To'] ?? null,
          monthlyAmt:         v['Monthly Amt'] ?? null,
          amtPerArea:         v['Amt/Area'] ?? null,
          annual:             v['Annual'] ?? null,
          annualPerArea:      v['Annual/Area'] ?? null,
          managementFee:      v['Management Fee'] ?? null,
          annualGrossAmount:  v['Annual Gross Amount'] ?? null,
        });
      }
    }
  }

  return rows;
}

// ─── Column definitions ───────────────────────────────────────────────────────

interface ColDef {
  key: keyof FlatRow;
  label: string;
  group: 'tenant' | 'schedule';
  right?: boolean;
}

const COLS: ColDef[] = [
  // ── Tenant identity
  { key: 'property',              label: 'Property',          group: 'tenant' },
  { key: 'unit',                  label: 'Unit',               group: 'tenant' },
  { key: 'lease',                 label: 'Tenant',             group: 'tenant' },
  { key: 'leaseType',             label: 'Lease Type',         group: 'tenant' },
  { key: 'area',                  label: 'Area',               group: 'tenant', right: true },
  { key: 'leaseFrom',             label: 'Lease From',         group: 'tenant' },
  { key: 'leaseTo',               label: 'Lease To',           group: 'tenant' },
  { key: 'monthlyRent',           label: 'Monthly Rent',       group: 'tenant', right: true },
  { key: 'annualRent',            label: 'Annual Rent',        group: 'tenant', right: true },
  { key: 'securityDepositReceived', label: 'Security Deposit', group: 'tenant', right: true },
  // ── Sub-section / schedule
  { key: '_section',              label: 'Section',            group: 'schedule' },
  { key: 'charge',                label: 'Charge',             group: 'schedule' },
  { key: 'chargeType',            label: 'Type',               group: 'schedule' },
  { key: 'chargeUnit',            label: 'Unit',               group: 'schedule' },
  { key: 'areaLabel',             label: 'Area Label',         group: 'schedule' },
  { key: 'from',                  label: 'From',               group: 'schedule' },
  { key: 'to',                    label: 'To',                 group: 'schedule' },
  { key: 'subArea',               label: 'Area',               group: 'schedule', right: true },
  { key: 'monthlyAmt',            label: 'Monthly Amt',        group: 'schedule', right: true },
  { key: 'amtPerArea',            label: 'Amt/Area',           group: 'schedule', right: true },
  { key: 'annual',                label: 'Annual',             group: 'schedule', right: true },
  { key: 'annualPerArea',         label: 'Annual/Area',        group: 'schedule', right: true },
  { key: 'managementFee',         label: 'Mgmt Fee',           group: 'schedule', right: true },
  { key: 'annualGrossAmount',     label: 'Annual Gross',       group: 'schedule', right: true },
];

// ─── CSV export ───────────────────────────────────────────────────────────────

function toCSV(rows: FlatRow[]): string {
  const header = COLS.map(c => `"${c.label}"`).join(',');
  const body = rows.map(row =>
    COLS.map(c => {
      const v = row[c.key];
      const s = fmt(v as Cell);
      return `"${s.replace(/"/g, '""')}"`;
    }).join(',')
  );
  return [header, ...body].join('\n');
}

function downloadCSV(rows: FlatRow[], fileName: string) {
  const csv = toCSV(rows);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.replace(/\.[^.]+$/, '') + '_extracted.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// ─── Component ────────────────────────────────────────────────────────────────

interface Props {
  tenants: TenancyScheduleTenant[];
  fileName: string;
  onBack: () => void;
}

export function TenancyScheduleTable({ tenants, fileName, onBack }: Props) {
  const rows = useMemo(() => flatten(tenants), [tenants]);

  // Section badge colour
  const sectionColour = (s: string) => {
    if (/rent step/i.test(s)) return 'text-blue-400 bg-blue-400/10 border-blue-400/30';
    if (/charge/i.test(s))    return 'text-amber-400 bg-amber-400/10 border-amber-400/30';
    return 'text-muted-foreground bg-muted border-panel-border';
  };

  // Group header rows by tenantIdx so we can draw a separator
  let lastTenantIdx = -1;

  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="shrink-0 flex items-center justify-between px-4 py-2 border-b border-panel-border bg-background">
        <div className="flex items-center gap-3">
          <button
            onClick={onBack}
            className="text-[11px] font-mono text-muted-foreground hover:text-foreground transition-colors flex items-center gap-1"
          >
            ← Back
          </button>
          <span className="text-[11px] font-mono text-foreground">
            {tenants.length} tenant{tenants.length !== 1 ? 's' : ''} · {rows.length} row{rows.length !== 1 ? 's' : ''}
          </span>
        </div>
        <button
          onClick={() => downloadCSV(rows, fileName)}
          className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors flex items-center gap-1.5"
        >
          ↓ Download CSV
        </button>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="text-[11px] font-mono border-collapse w-full">
          <thead className="sticky top-0 z-10">
            {/* Group header */}
            <tr>
              {/* tenant columns */}
              <th
                colSpan={COLS.filter(c => c.group === 'tenant').length}
                className="px-2 py-1 text-left border border-panel-border bg-primary/10 text-primary font-medium tracking-wide"
              >
                Tenant
              </th>
              {/* schedule columns */}
              <th
                colSpan={COLS.filter(c => c.group === 'schedule').length}
                className="px-2 py-1 text-left border border-panel-border bg-amber-500/10 text-amber-400 font-medium tracking-wide"
              >
                Rent Steps &amp; Charge Schedules
              </th>
            </tr>
            {/* Column labels */}
            <tr>
              {COLS.map(col => (
                <th
                  key={col.key}
                  className={[
                    'px-2 py-1 border border-panel-border whitespace-nowrap font-medium',
                    col.group === 'tenant'
                      ? 'bg-primary/5 text-primary'
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
            {rows.map((row, ri) => {
              const isTenantBoundary = row._tenantIdx !== lastTenantIdx;
              lastTenantIdx = row._tenantIdx;

              return (
                <tr
                  key={ri}
                  className={[
                    'hover:bg-muted/30 transition-colors',
                    isTenantBoundary && ri > 0 ? 'border-t-2 border-t-primary/20' : '',
                  ].join(' ')}
                >
                  {COLS.map(col => {
                    const raw = row[col.key];
                    const display = col.key === '_section'
                      ? (raw as string)
                      : fmt(raw as Cell);

                    if (col.key === '_section') {
                      return (
                        <td key={col.key} className="px-2 py-1 border border-panel-border whitespace-nowrap">
                          {display ? (
                            <span className={`px-1.5 py-0.5 rounded border text-[10px] font-mono ${sectionColour(display)}`}>
                              {display}
                            </span>
                          ) : null}
                        </td>
                      );
                    }

                    return (
                      <td
                        key={col.key}
                        className={[
                          'px-2 py-1 border border-panel-border whitespace-nowrap',
                          col.group === 'tenant' ? 'text-foreground' : 'text-muted-foreground',
                          col.right ? 'text-right tabular-nums' : '',
                          !display ? 'text-muted-foreground/30' : '',
                        ].join(' ')}
                      >
                        {display || '—'}
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