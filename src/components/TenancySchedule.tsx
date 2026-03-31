// src/components/TenancyScheduleTable.tsx
import { useMemo, useState } from 'react';
import type { TenancyScheduleTenant } from '@/lib/rent-roll-types/tenancy-schedule-parser';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { MappingDialog, pairKey } from './MappingDialog';
import type { UniqueChargePair } from './MappingDialog';

// ─── Types ────────────────────────────────────────────────────────────────────

type Cell = string | number | Date | null;

interface FlatRow {
  _tenantIdx: number;
  _section: string;
  _isSplit: boolean;         // true = this row came from a multi-unit split
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
  // sub-section fields
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

type SubValues = Record<string, Cell>;

const DATE_COL_KEYS = new Set<keyof FlatRow>(['leaseFrom', 'leaseTo', 'from', 'to']);

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

function toDateString(v: Cell): string | null {
  if (v === null || v === undefined) return null;

  if (v instanceof Date) {
    const y = v.getUTCFullYear();
    const m = v.getUTCMonth() + 1;
    const d = v.getUTCDate();
    return `${pad2(m)}/${pad2(d)}/${y}`;
  }

  if (typeof v === 'number' && v > 20000 && v < 90000) {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const dt = new Date(ms);
    if (!isNaN(dt.getTime())) {
      return `${pad2(dt.getUTCMonth() + 1)}/${pad2(dt.getUTCDate())}/${dt.getUTCFullYear()}`;
    }
  }

  if (typeof v !== 'string') return null;
  const s = v.trim();
  if (!s) return null;

  const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+.*)?$/i);
  if (us) {
    const mm = Number(us[1]);
    const dd = Number(us[2]);
    const yy = us[3];
    const yyyy = yy.length === 2 ? (Number(yy) >= 70 ? 1900 + Number(yy) : 2000 + Number(yy)) : Number(yy);
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31 && yyyy >= 1900 && yyyy <= 9999) {
      return `${pad2(mm)}/${pad2(dd)}/${yyyy}`;
    }
  }

  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  if (iso) {
    const yyyy = Number(iso[1]);
    const mm = Number(iso[2]);
    const dd = Number(iso[3]);
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      return `${pad2(mm)}/${pad2(dd)}/${yyyy}`;
    }
  }

  return null;
}

function dateSortValue(v: Cell): number {
  if (typeof v === 'number') return v;
  const d = toDateString(v);
  if (!d) return 0;
  const m = d.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return 0;
  return Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
}

function dateKey(v: Cell): string {
  const d = toDateString(v);
  if (d) return d;
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  return '';
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function fmt(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return toDateString(v) ?? '';
  if (typeof v === 'number') {
    if (Math.abs(v) >= 1000) return v.toLocaleString('en-US', { maximumFractionDigits: 2 });
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

/**
 * Coerce a cell value to a JS number, handling both native number cells
 * and string cells formatted with commas (e.g. "16,500.00" from Excel).
 * Returns null if the value cannot be parsed as a number.
 */
function toNumber(v: Cell): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = parseFloat(v.replace(/,/g, '').trim());
    return isNaN(n) ? null : n;
  }
  return null;
}


// ─── Unit splitting utilities ─────────────────────────────────────────────────

/**
 * Parse a unit cell value into sorted individual unit IDs.
 * "A0105,  A102" → ["A0105", "A102"]  (sorted for stable set comparison)
 */
function parseUnits(val: Cell): string[] {
  if (!val) return [];
  return String(val)
    .split(',')
    .map(u => u.trim())
    .filter(Boolean)
    .sort();
}

/** True if two sorted unit arrays represent the same set */
function sameUnitSet(a: string[], b: string[]): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

/** area map: unit ID → area (sqft), scraped from exact-match sub-section rows */
type AreaMap = Record<string, number>;

function buildAreaMap(
  subSections: TenancyScheduleTenant['subSections'],
  mainUnits: string[],
): AreaMap {
  const map: AreaMap = {};
  for (const section of subSections) {
    for (const row of section.rows) {
      const rowUnits = parseUnits(row.values['Unit'] ?? null);
      if (rowUnits.length === 1 && mainUnits.includes(rowUnits[0])) {
        const area = toNumber(row.values['Area'] ?? null);
        if (area !== null && map[rowUnits[0]] === undefined) {
          map[rowUnits[0]] = area;
        }
      }
    }
  }
  return map;
}

/**
 * Decide whether this tenant needs to be split into per-unit rows.
 *
 * Conditions for split:
 *   1. Main row Unit(s) contains commas  → multiple units
 *   2. At least one sub-section data row uses an individual unit ID  → split required
 *   If all sub-section rows use only the combined string → no split, keep as-is
 */
type SplitAnalysis =
  | { needsSplit: false }
  | { needsSplit: true; mainUnits: string[]; areaMap: AreaMap; totalArea: number };

function analyseMultiUnit(t: TenancyScheduleTenant): SplitAnalysis {
  const rawUnit = pick(t.mainRow, 'Unit(s)', 'Unit');
  const mainUnits = parseUnits(rawUnit);

  if (mainUnits.length <= 1) return { needsSplit: false };

  // Look for any sub-section row that names an individual unit
  let hasIndividualRow = false;
  outer: for (const section of t.subSections) {
    for (const row of section.rows) {
      const rowUnits = parseUnits(row.values['Unit'] ?? null);
      if (rowUnits.length === 1 && mainUnits.includes(rowUnits[0])) {
        hasIndividualRow = true;
        break outer;
      }
    }
  }

  if (!hasIndividualRow) return { needsSplit: false };

  const areaMap = buildAreaMap(t.subSections, mainUnits);
  const totalArea = Object.values(areaMap).reduce((s, a) => s + a, 0);

  return { needsSplit: true, mainUnits, areaMap, totalArea };
}

/**
 * Fields that are absolute amounts → scale by area weight.
 * Per-area rate fields (Amt/Area, Annual/Area) → unchanged (rate stays constant).
 * Area → replaced with this unit's area.
 * Unit → replaced with this unit's ID.
 * Everything else → copied as-is.
 */
const WEIGHTED_FIELDS = new Set([
  'Monthly Amt',
  'Annual',
  'Annual Gross Amount',
  'Management Fee',
]);

function applyWeight(v: Cell, weight: number): Cell {
  const n = toNumber(v);
  if (n === null) return v;
  return n * weight;  // full float precision — no rounding
}

/**
 * Produce sub-section values for a single unit from a combined-unit row.
 * Weight is (thisUnitArea / combinedRowTotalArea) — NOT totalArea of all main units,
 * because the combined row may only cover a subset of the main units.
 */
function splitSubValues(
  values: SubValues,
  targetUnit: string,
  unitArea: number | null,
  combinedWeight: number,
): SubValues {
  const out: SubValues = { ...values };
  out['Unit'] = targetUnit;
  if (unitArea !== null) out['Area'] = unitArea;
  for (const field of WEIGHTED_FIELDS) {
    if (field in out) out[field] = applyWeight(out[field], combinedWeight);
  }
  return out;
}

// ─── Row builders ─────────────────────────────────────────────────────────────

interface TenantBase extends Omit<FlatRow,
  '_section' | 'charge' | 'chargeType' | 'chargeUnit' | 'areaLabel' |
  'subArea' | 'from' | 'to' | 'monthlyAmt' | 'amtPerArea' |
  'annual' | 'annualPerArea' | 'managementFee' | 'annualGrossAmount'
> {}

function buildTenantBase(
  mr: Record<string, Cell>,
  idx: number,
  overrides?: { unit: string; area: number | null; weight: number },
): TenantBase {
  const isSplit = !!overrides;
  const w = overrides?.weight ?? 1;
  const wNum = (v: Cell): Cell => {
    if (!isSplit) return v;
    const n = toNumber(v);
    return n !== null ? n * w : v;  // full float precision — no rounding
  };

  return {
    _tenantIdx:              idx,
    _isSplit:                isSplit,
    property:                pick(mr, 'Property'),
    unit:                    overrides ? overrides.unit : pick(mr, 'Unit(s)', 'Unit'),
    lease:                   pick(mr, 'Lease'),
    leaseType:               pick(mr, 'Lease Type'),
    area:                    overrides ? (overrides.area ?? pick(mr, 'Area')) : pick(mr, 'Area'),
    leaseFrom:               pick(mr, 'Lease From'),
    leaseTo:                 pick(mr, 'Lease To'),
    term:                    pick(mr, 'Term'),
    tenancyYears:            pick(mr, 'Tenancy Years'),
    monthlyRent:             wNum(pick(mr, 'Monthly Rent')),
    monthlyRentPerArea:      pick(mr, 'Monthly Rent/Area'),    // rate → unchanged
    annualRent:              wNum(pick(mr, 'Annual Rent')),
    annualRentPerArea:       pick(mr, 'Annual Rent/Area'),     // rate → unchanged
    annualRecPerArea:        pick(mr, 'Annual Rec./Area'),     // rate → unchanged
    annualMiscPerArea:       pick(mr, 'Annual Misc/Area'),     // rate → unchanged
    securityDepositReceived: wNum(pick(mr, 'Security Deposit Received')),
    locAmount:               wNum(pick(mr, 'LOC Amount/ Bank Guarantee')),
  };
}

function buildScheduleRow(base: TenantBase, sectionName: string, v: SubValues): FlatRow {
  return {
    ...base,
    _section:          sectionName,
    charge:            v['Charge'] ?? null,
    chargeType:        v['Type'] ?? null,
    chargeUnit:        v['Unit'] ?? null,
    areaLabel:         v['Area Label'] ?? null,
    subArea:           v['Area'] ?? null,
    from:              v['From'] ?? null,
    to:                v['To'] ?? null,
    monthlyAmt:        v['Monthly Amt'] ?? null,
    amtPerArea:        v['Amt/Area'] ?? null,
    annual:            v['Annual'] ?? null,
    annualPerArea:     v['Annual/Area'] ?? null,
    managementFee:     v['Management Fee'] ?? null,
    annualGrossAmount: v['Annual Gross Amount'] ?? null,
  };
}

function emptyScheduleRow(base: TenantBase): FlatRow {
  return {
    ...base,
    _section: '',
    charge: null, chargeType: null, chargeUnit: null, areaLabel: null,
    subArea: null, from: null, to: null,
    monthlyAmt: null, amtPerArea: null, annual: null, annualPerArea: null,
    managementFee: null, annualGrossAmount: null,
  };
}

// ─── Flatten ──────────────────────────────────────────────────────────────────

function flatten(tenants: TenancyScheduleTenant[]): FlatRow[] {
  const rows: FlatRow[] = [];
  let idx = 0; // _tenantIdx — increments per emitted tenant or per-unit split

  for (const t of tenants) {
    const analysis = analyseMultiUnit(t);

    // ── Single unit, or combined rows everywhere: no splitting ──
    if (!analysis.needsSplit) {
      const base = buildTenantBase(t.mainRow, idx++);
      if (t.subSections.length === 0) {
        rows.push(emptyScheduleRow(base));
        continue;
      }
      for (const section of t.subSections) {
        for (const dataRow of section.rows) {
          rows.push(buildScheduleRow(base, section.name, dataRow.values));
        }
      }
      continue;
    }

    // ── Multi-unit split: one logical tenant per unit ──
    const { mainUnits, areaMap } = analysis;

    for (const unit of mainUnits) {
      const unitArea = areaMap[unit] ?? null;
      // Fallback to equal split if we couldn't find this unit's area
      const mainWeight = (unitArea !== null && analysis.totalArea > 0)
        ? unitArea / analysis.totalArea
        : 1 / mainUnits.length;

      const base = buildTenantBase(t.mainRow, idx++, { unit, area: unitArea, weight: mainWeight });

      if (t.subSections.length === 0) {
        rows.push(emptyScheduleRow(base));
        continue;
      }

      for (const section of t.subSections) {
        for (const dataRow of section.rows) {
          const rowUnits = parseUnits(dataRow.values['Unit'] ?? null);

          if (sameUnitSet(rowUnits, [unit])) {
            // Exact match → include as-is
            rows.push(buildScheduleRow(base, section.name, dataRow.values));

          } else if (rowUnits.length > 1 && rowUnits.includes(unit)) {
            // Combined row (e.g. "A0105, A102") → split this row by area weight.
            // Use the area of the units actually referenced in THIS row, not all main units,
            // so partial-combination rows are weighted correctly.
            const combinedArea = rowUnits.reduce(
              (sum, u) => sum + (areaMap[u] ?? 0), 0,
            );
            const combinedWeight = (unitArea !== null && combinedArea > 0)
              ? unitArea / combinedArea
              : 1 / rowUnits.length;

            const weightedValues = splitSubValues(
              dataRow.values, unit, unitArea, combinedWeight,
            );
            rows.push(buildScheduleRow(base, section.name, weightedValues));

          }
          // Belongs to a different single unit → skip
        }
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
  // ── Tenant
  { key: 'property',               label: 'Property',          group: 'tenant' },
  { key: 'unit',                   label: 'Unit',               group: 'tenant' },
  { key: 'lease',                  label: 'Tenant',             group: 'tenant' },
  { key: 'leaseType',              label: 'Lease Type',         group: 'tenant' },
  { key: 'area',                   label: 'Area',               group: 'tenant', right: true },
  { key: 'leaseFrom',              label: 'Lease From',         group: 'tenant' },
  { key: 'leaseTo',                label: 'Lease To',           group: 'tenant' },
  { key: 'monthlyRent',            label: 'Monthly Rent',       group: 'tenant', right: true },
  { key: 'annualRent',             label: 'Annual Rent',        group: 'tenant', right: true },
  { key: 'securityDepositReceived', label: 'Security Deposit',  group: 'tenant', right: true },
  // ── Schedule
  { key: '_section',               label: 'Section',            group: 'schedule' },
  { key: 'charge',                 label: 'Charge',             group: 'schedule' },
  { key: 'chargeType',             label: 'Type',               group: 'schedule' },
  { key: 'chargeUnit',             label: 'Unit',               group: 'schedule' },
  { key: 'areaLabel',              label: 'Area Label',         group: 'schedule' },
  { key: 'from',                   label: 'From',               group: 'schedule' },
  { key: 'to',                     label: 'To',                 group: 'schedule' },
  { key: 'subArea',                label: 'Area',               group: 'schedule', right: true },
  { key: 'monthlyAmt',             label: 'Monthly Amt',        group: 'schedule', right: true },
  { key: 'amtPerArea',             label: 'Amt/Area',           group: 'schedule', right: true },
  { key: 'annual',                 label: 'Annual',             group: 'schedule', right: true },
  { key: 'annualPerArea',          label: 'Annual/Area',        group: 'schedule', right: true },
  { key: 'managementFee',          label: 'Mgmt Fee',           group: 'schedule', right: true },
  { key: 'annualGrossAmount',      label: 'Annual Gross',       group: 'schedule', right: true },
];

// ─── Excel export ─────────────────────────────────────────────────────────────

async function downloadXLSX(
  rows: FlatRow[],
  fileName: string,
  mappings: Record<string, string> = {},
) {
  type X = string | number | Date | null;

  // ── Main columns (Property → Security Deposit) ────────────────────────────
  const MAIN_KEYS: (keyof FlatRow)[] = [
    'property', 'unit', 'lease', 'leaseType', 'area',
    'leaseFrom', 'leaseTo', 'monthlyRent', 'annualRent', 'securityDepositReceived',
  ];
  const MAIN_HDRS = [
    'Property', 'Unit', 'Tenant', 'Lease Type', 'Area (sqft)',
    'Lease From', 'Lease To', 'Monthly Rent', 'Annual Rent', 'Security Deposit',
  ];
  const nM = MAIN_KEYS.length;

  // ── Group flat rows by tenant ─────────────────────────────────────────────
  interface TG { base: FlatRow; rs: FlatRow[]; cs: FlatRow[] }
  const tMap = new Map<number, TG>();
  for (const row of rows) {
    const i = row._tenantIdx;
    if (!tMap.has(i)) tMap.set(i, { base: row, rs: [], cs: [] });
    const g = tMap.get(i)!;
    const sec = String(row._section ?? '').toLowerCase();
    if (sec.includes('rent') && sec.includes('step')) g.rs.push(row);
    else if (row._section) g.cs.push(row);
  }
  const tenants = [...tMap.values()];

  // ── Per-tenant RS key sets: compound (charge + fromDate) ─────────────────
  // Used to exclude from CS only those (charge, date) pairs already in RS for that tenant.
  const rsKeySetByTenant = new Map<number, Set<string>>();
  for (const [idx, { rs }] of tMap.entries()) {
    const ks = new Set<string>();
    for (const r of rs) {
      const c = String(r.charge ?? '').trim();
      const f = dateKey(r.from);
      if (c) ks.add(`${c}\x00${f}`);
    }
    rsKeySetByTenant.set(idx, ks);
  }

  // Filter a tenant's CS rows: remove any (charge, from) already present in their RS
  const filteredCS = (tenantIdx: number, cs: FlatRow[]): FlatRow[] => {
    const ks = rsKeySetByTenant.get(tenantIdx) ?? new Set<string>();
    if (ks.size === 0) return cs;
    return cs.filter(r => {
      const c = String(r.charge ?? '').trim();
      const f = dateKey(r.from);
      return !ks.has(`${c}\x00${f}`);
    });
  };

  // ── Helper: date value → sort key ────────────────────────────────────────
  const dNum = (v: Cell): number => dateSortValue(v);

  // ── Unique charge codes for Rent Steps ───────────────────────────────────
  const rsCodes: string[] = [];
  const rsSet = new Set<string>();
  for (const { rs } of tenants)
    for (const r of rs) { const c = String(r.charge ?? '').trim(); if (c && !rsSet.has(c)) { rsSet.add(c); rsCodes.push(c); } }

  const rsMax: Record<string, number> = Object.fromEntries(rsCodes.map(c => [c, 0]));
  for (const { rs } of tenants) {
    const cnt: Record<string, number> = {};
    for (const r of rs) { const c = String(r.charge ?? '').trim(); if (c) cnt[c] = (cnt[c] ?? 0) + 1; }
    for (const c of rsCodes) rsMax[c] = Math.max(rsMax[c], cnt[c] ?? 0);
  }

  // ── Unique charge codes for Charge Schedules ─────────────────────────────
  // A code gets a CS column if any tenant has CS rows for it after per-tenant
  // (charge, from) exclusion. Codes can overlap with rsCodes — that's intentional.
  const csCodes: string[] = [];
  const csSet = new Set<string>();
  const codeType: Record<string, string> = {};
  for (const { base, cs } of tenants)
    for (const r of filteredCS(base._tenantIdx, cs)) {
      const c = String(r.charge ?? '').trim();
      if (!c) continue;
      if (!codeType[c]) codeType[c] = String(r.chargeType ?? '').trim();
      if (!csSet.has(c)) { csSet.add(c); csCodes.push(c); }
    }

  // Sort CS codes by mapping category
  const MAP_ORD = ['Rent', 'Opex', 'Utility', 'Management', 'Insurance', 'Tax', 'Excluded'];
  csCodes.sort((a, b) => {
    const ia = MAP_ORD.indexOf(mappings[pairKey(a, codeType[a] ?? '')] ?? '');
    const ib = MAP_ORD.indexOf(mappings[pairKey(b, codeType[b] ?? '')] ?? '');
    return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
  });

  const csMax: Record<string, number> = Object.fromEntries(csCodes.map(c => [c, 0]));
  for (const { base, cs } of tenants) {
    const cnt: Record<string, number> = {};
    for (const r of filteredCS(base._tenantIdx, cs)) { const c = String(r.charge ?? '').trim(); if (c) cnt[c] = (cnt[c] ?? 0) + 1; }
    for (const c of csCodes) csMax[c] = Math.max(csMax[c], cnt[c] ?? 0);
  }

  // ── Column layout ─────────────────────────────────────────────────────────
  const rsTotal = rsCodes.reduce((s, c) => s + 2 * rsMax[c], 0);
  const csTotal = csCodes.reduce((s, c) => s + 2 * csMax[c], 0);

  const COL_RS    = nM;                         // rent steps start
  const COL_BLANK = COL_RS + rsTotal;            // separator (only when both sections present)
  const COL_CS    = rsTotal > 0 ? COL_BLANK + 1 : nM;
  const TOTAL     = (csTotal > 0 ? COL_CS + csTotal : COL_RS + rsTotal);

  const rsStart = (code: string): number => { let c = COL_RS;  for (const x of rsCodes) { if (x === code) return c; c += 2 * rsMax[x]; } return c; };
  const csStart = (code: string): number => { let c = COL_CS;  for (const x of csCodes) { if (x === code) return c; c += 2 * csMax[x]; } return c; };

  // ── Build 4 header rows ───────────────────────────────────────────────────
  const mk = (): X[] => Array<X>(TOTAL).fill(null);
  const h1 = mk(); const h2 = mk(); const h3 = mk(); const h4 = mk();

  // Fill every cell — no merges
  for (let i = 0; i < nM; i++) h1[i] = 'Tenant Info';
  for (let i = 0; i < nM; i++) h4[i] = MAIN_HDRS[i];

  if (rsTotal > 0) for (let i = COL_RS; i < COL_RS + rsTotal; i++) h1[i] = 'Rent Steps';
  if (csTotal > 0) for (let i = COL_CS; i < COL_CS + csTotal; i++) h1[i] = 'Charge Schedules';

  for (const code of rsCodes) {
    const s = rsStart(code);
    for (let p = 0; p < rsMax[code]; p++) {
      h2[s + p * 2]     = code;
      h2[s + p * 2 + 1] = code;
      h3[s + p * 2]     = `Rent Step ${p + 1}`;
      h3[s + p * 2 + 1] = `Rent Step ${p + 1}`;
      h4[s + p * 2]     = `Rent Date ${p + 1}`;
      h4[s + p * 2 + 1] = `Rent Rate ${p + 1}`;
    }
  }

  for (const code of csCodes) {
    const s = csStart(code);
    const label = mappings[pairKey(code, codeType[code] ?? '')] || code;
    for (let p = 0; p < csMax[code]; p++) {
      h2[s + p * 2]     = code;
      h2[s + p * 2 + 1] = code;
      h3[s + p * 2]     = label;
      h3[s + p * 2 + 1] = label;
      h4[s + p * 2]     = `Date ${p + 1}`;
      h4[s + p * 2 + 1] = `Rate ${p + 1}`;
    }
  }

  // ── Build data rows (one per tenant, starting at row 5) ──────────────────
  const rate = (r: FlatRow, base: FlatRow): number | null => {
    const apa = toNumber(r.annualPerArea);
    if (apa !== null) return apa;
    const ann = toNumber(r.annual); const area = toNumber(base.area);
    return ann !== null && area ? ann / area : null;
  };

  // leaseFrom / leaseTo column indices (needed before building data rows)
  const dateMainCols = new Set(
    (['leaseFrom', 'leaseTo'] as (keyof FlatRow)[]).map(k => MAIN_KEYS.indexOf(k)).filter(i => i >= 0)
  );

  const dataRows: X[][] = tenants.map(({ base, rs, cs }) => {
    const row = mk();
    for (let i = 0; i < nM; i++) {
      const v = base[MAIN_KEYS[i]] as Cell;
      if (dateMainCols.has(i)) {
        row[i] = toDateString(v);
      } else {
        row[i] = v instanceof Date ? toDateString(v) : typeof v === 'number' ? v : (v as string | null) ?? null;
      }
    }
    for (const code of rsCodes) {
      const steps = rs.filter(r => String(r.charge ?? '').trim() === code).sort((a, b) => dNum(a.from) - dNum(b.from));
      const s = rsStart(code);
      steps.forEach((st, p) => { row[s + p * 2] = toDateString(st.from); row[s + p * 2 + 1] = rate(st, base); });
    }
    for (const code of csCodes) {
      const charges = filteredCS(base._tenantIdx, cs).filter(r => String(r.charge ?? '').trim() === code).sort((a, b) => dNum(a.from) - dNum(b.from));
      const s = csStart(code);
      charges.forEach((ch, p) => { row[s + p * 2] = toDateString(ch.from); row[s + p * 2 + 1] = rate(ch, base); });
    }
    return row;
  });

  // ── Build styled workbook with ExcelJS ───────────────────────────────────
  const wb2 = new ExcelJS.Workbook();
  wb2.creator = 'Rent Roll Analyst';

  const ws2 = wb2.addWorksheet('Rent Roll', {
    views: [{ state: 'frozen', xSplit: nM, ySplit: 4, showGridLines: false }],
  });

  // Column widths
  ws2.columns = Array.from({ length: TOTAL }, (_, i) => ({
    width: i === 0 ? 34
         : (i === COL_BLANK && rsTotal > 0 && csTotal > 0) ? 2
         : i < nM ? 16
         : 14,
  }));

  // ── Color palette (ARGB) ─────────────────────────────────────────────────
  // Tenant Info: blue family  |  Rent Steps: green family  |  CS: orange family
  const PAL = {
    ti1: 'FF1F3864', ti2: 'FF2E75B6', ti3: 'FF9DC3E6', ti4: 'FFDEEAF1',
    rs1: 'FF1E4620', rs2: 'FF548235', rs3: 'FFA9D18E', rs4: 'FFE2EFDA',
    cs1: 'FF833C00', cs2: 'FFC55A11', cs3: 'FFF4B183', cs4: 'FFFCE4D6',
    blank: 'FF303030',
    white: 'FFFFFFFF', dark: 'FF1A1A1A',
    rowOdd: 'FFFFFFFF', rowEven: 'FFF5F7FA',
    borderColor: 'FFB8C4CE',
  };

  const DARK_FILLS = new Set([PAL.ti1, PAL.ti2, PAL.rs1, PAL.rs2, PAL.cs1, PAL.cs2, PAL.blank]);

  const mkFill = (argb: string): ExcelJS.Fill =>
    ({ type: 'pattern', pattern: 'solid', fgColor: { argb } } as ExcelJS.Fill);

  const mkBorder = (weight: 'hair' | 'thin' | 'medium' = 'thin'): Partial<ExcelJS.Borders> => {
    const side = { style: weight as ExcelJS.BorderStyle, color: { argb: PAL.borderColor } };
    return { top: side, left: side, bottom: side, right: side };
  };

  const colSection = (ci: number): 'tenant' | 'rs' | 'cs' | 'blank' => {
    if (ci < nM) return 'tenant';
    if (rsTotal > 0 && ci >= COL_RS && ci < COL_RS + rsTotal) return 'rs';
    if (ci === COL_BLANK && rsTotal > 0 && csTotal > 0) return 'blank';
    if (csTotal > 0 && ci >= COL_CS && ci < COL_CS + csTotal) return 'cs';
    return 'blank';
  };

  const hdrFill = (ci: number, level: 1 | 2 | 3 | 4): string => {
    const s = colSection(ci);
    if (s === 'blank') return PAL.blank;
    const map = {
      tenant: [PAL.ti1, PAL.ti2, PAL.ti3, PAL.ti4],
      rs:     [PAL.rs1, PAL.rs2, PAL.rs3, PAL.rs4],
      cs:     [PAL.cs1, PAL.cs2, PAL.cs3, PAL.cs4],
    } as const;
    return map[s][level - 1];
  };

  // ── Add 4 header rows ─────────────────────────────────────────────────────
  const HDR_HEIGHTS = [22, 18, 18, 16];
  const hdrs = [h1, h2, h3, h4];

  hdrs.forEach((hdr, hi) => {
    const level = (hi + 1) as 1 | 2 | 3 | 4;
    const exRow = ws2.addRow(hdr as (string | number | Date | null)[]);
    exRow.height = HDR_HEIGHTS[hi];
    exRow.eachCell({ includeEmpty: true }, (cell, colIdx) => {
      const ci = colIdx - 1;
      const bg = hdrFill(ci, level);
      cell.fill = mkFill(bg);
      cell.font = {
        bold: true,
        color: { argb: DARK_FILLS.has(bg) ? PAL.white : PAL.dark },
        size: level === 1 ? 11 : 10,
        name: 'Calibri',
      };
      cell.border = mkBorder(level === 4 ? 'medium' : 'thin');
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    });
  });

  dataRows.forEach((dataRow, ri) => {
    const exRow = ws2.addRow(dataRow as (string | number | Date | null)[]);
    exRow.height = 15;
    const rowBg = ri % 2 === 0 ? PAL.rowOdd : PAL.rowEven;
    exRow.eachCell({ includeEmpty: true }, (cell, colIdx) => {
      const ci = colIdx - 1;
      cell.fill = mkFill(rowBg);
      cell.font = { size: 10, name: 'Calibri', color: { argb: PAL.dark } };
      cell.border = mkBorder('hair');

      if (dateMainCols.has(ci) || (ci >= nM && h4[ci] != null && String(h4[ci]).toLowerCase().includes('date'))) {
        if (typeof cell.value === 'number') cell.numFmt = 'mm/dd/yyyy';
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (typeof cell.value === 'number') {
        cell.numFmt = '#,##0.00';
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: ci < nM ? 'left' : 'center' };
      }
    });
  });

  // ── Download ──────────────────────────────────────────────────────────────
  const buf = await wb2.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.replace(/\.[^.]+$/, '') + '_extracted.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── Flat (raw) export ────────────────────────────────────────────────────────

function downloadFlatXLSX(rows: FlatRow[], fileName: string) {
  const header = COLS.map(c => c.label);
  const wsData: (string | number | Date | null)[][] = [header];
  for (const row of rows) {
    wsData.push(
      COLS.map(col => {
        const v = row[col.key];
        if (col.key === '_isSplit') return null;
        if (v instanceof Date) return v;
        if (typeof v === 'number') return v;
        return fmt(v as Cell) || null;
      })
    );
  }
  const ws = XLSX.utils.aoa_to_sheet(wsData, { cellDates: true });
  ws['!cols'] = COLS.map(col =>
    col.right ? { wch: 14 } : col.key === 'property' ? { wch: 38 } : { wch: 18 }
  );
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Rent Roll');
  XLSX.writeFile(wb, fileName.replace(/\.[^.]+$/, '') + '_flat.xlsx');
}

// ─── Component ────────────────────────────────────────────────────────────────

interface Props {
  tenants: TenancyScheduleTenant[];
  fileName: string;
  onBack: () => void;
}

export function TenancyScheduleTable({ tenants, fileName, onBack }: Props) {
  const rows = useMemo(() => flatten(tenants), [tenants]);

  // Unique (charge, chargeType) pairs for the mapping dialog, sorted by Lease Type then Code
  const uniquePairs = useMemo<UniqueChargePair[]>(() => {
    const seen = new Set<string>();
    const pairs: UniqueChargePair[] = [];
    for (const row of rows) {
      const charge     = String(row.charge     ?? '').trim();
      const chargeType = String(row.chargeType ?? '').trim();
      if (!charge) continue;
      const k = pairKey(charge, chargeType);
      if (!seen.has(k)) { seen.add(k); pairs.push({ charge, chargeType }); }
    }
    return pairs.sort((a, b) => {
      const ct = a.chargeType.localeCompare(b.chargeType);
      return ct !== 0 ? ct : a.charge.localeCompare(b.charge);
    });
  }, [rows]);

  const [showMapping, setShowMapping] = useState(false);

  const sectionColour = (s: string) => {
    if (/rent step/i.test(s)) return 'text-blue-400 bg-blue-400/10 border-blue-400/30';
    if (/charge/i.test(s))    return 'text-amber-400 bg-amber-400/10 border-amber-400/30';
    return 'text-muted-foreground bg-muted border-panel-border';
  };

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
        <div className="flex items-center gap-2">
          <button
            onClick={() => downloadFlatXLSX(rows, fileName)}
            className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors flex items-center gap-1.5"
          >
            ↓ Raw Export
          </button>
          <button
            onClick={() => setShowMapping(true)}
            className="px-3 py-1.5 text-[11px] font-mono rounded border border-panel-border bg-background hover:border-muted-foreground text-foreground transition-colors flex items-center gap-1.5"
          >
            ↓ Structured Export
          </button>
        </div>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="text-[11px] font-mono border-collapse w-full">
          <thead className="sticky top-0 z-10">
            <tr>
              <th
                colSpan={COLS.filter(c => c.group === 'tenant').length}
                className="px-2 py-1 text-left border border-panel-border bg-primary/10 text-primary font-medium tracking-wide"
              >
                Tenant
              </th>
              <th
                colSpan={COLS.filter(c => c.group === 'schedule').length}
                className="px-2 py-1 text-left border border-panel-border bg-amber-500/10 text-amber-400 font-medium tracking-wide"
              >
                Rent Steps &amp; Charge Schedules
              </th>
            </tr>
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
                    if (col.key === '_isSplit') return null; // internal only

                    const raw = row[col.key];

                    if (col.key === '_section') {
                      const display = raw as string;
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

                    const display = DATE_COL_KEYS.has(col.key)
                      ? (toDateString(raw as Cell) ?? (typeof raw === 'string' ? raw.trim() : ''))
                      : fmt(raw as Cell);
                    return (
                      <td
                        key={col.key}
                        className={[
                          'px-2 py-1 border border-panel-border whitespace-nowrap',
                          col.group === 'tenant' ? 'text-foreground' : 'text-muted-foreground',
                          col.right ? 'text-right tabular-nums' : '',
                          !display ? 'text-muted-foreground/30' : '',
                          // Italic on weighted tenant-level amounts so it's clear they were split
                          row._isSplit && col.right && col.group === 'tenant' ? 'italic' : '',
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

      {showMapping && (
        <MappingDialog
          uniquePairs={uniquePairs}
          onClose={() => setShowMapping(false)}
          onExport={(mappings, _categories) => {
            downloadXLSX(rows, fileName, mappings).then(() => setShowMapping(false));
          }}
        />
      )}
    </div>
  );
}