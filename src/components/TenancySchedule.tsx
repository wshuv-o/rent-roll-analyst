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
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

/** Convert a Date or date string (mm/dd/yyyy) to an Excel serial number, timezone-safe */
function toExcelSerial(v: Cell): number | null {
  if (v instanceof Date) {
    const serial = Math.round((v.getTime() - EXCEL_EPOCH_UTC) / 86400000);
    return serial > 0 ? serial : null;
  }
  if (typeof v === 'number' && Number.isFinite(v) && v > 0) return Math.round(v);
  const s = toDateString(v);
  if (!s) return null;
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  const ms = Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
  return Math.round((ms - EXCEL_EPOCH_UTC) / 86400000);
}

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

function serialToDateString(serial: number): string | null {
  if (!Number.isFinite(serial)) return null;
  const p = XLSX.SSF.parse_date_code(Math.round(serial));
  if (!p || !p.y || !p.m || !p.d) return null;
  return `${pad2(p.m)}/${pad2(p.d)}/${p.y}`;
}

function toDateString(v: Cell): string | null {
  if (v === null || v === undefined) return null;

  if (v instanceof Date) {
    // Convert Date -> Excel serial day and round to nearest day.
    // This strips timezone/time artifacts and keeps the original worksheet day.
    const serial = (v.getTime() - EXCEL_EPOCH_UTC) / 86400000;
    const fromSerial = serialToDateString(serial);
    if (fromSerial) return fromSerial;
  }

  if (typeof v === 'number' && Number.isFinite(v) && v > 0) {
    return serialToDateString(v);
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
  if (typeof v === 'number') return Math.round(v);
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
  categories: string[] = [],
  rentRollDate: string = '',
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

  const filteredCS = (tenantIdx: number, cs: FlatRow[]): FlatRow[] => {
    const ks = rsKeySetByTenant.get(tenantIdx) ?? new Set<string>();
    if (ks.size === 0) return cs;
    return cs.filter(r => {
      const c = String(r.charge ?? '').trim();
      const f = dateKey(r.from);
      return !ks.has(`${c}\x00${f}`);
    });
  };

  // ── Helpers ───────────────────────────────────────────────────────────────
  const dNum = (v: Cell): number => dateSortValue(v);

  // 0-based index → Excel column letter (0→A, 25→Z, 26→AA)
  const CL = (i: number): string => {
    let n = i + 1, s = '';
    while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
    return s;
  };

  // Build a mapping from charge code → category using the (code, type) pair key
  const codeType: Record<string, string> = {};
  for (const { rs, cs } of tenants) {
    for (const r of [...rs, ...cs]) {
      const c = String(r.charge ?? '').trim();
      if (c && !codeType[c]) codeType[c] = String(r.chargeType ?? '').trim();
    }
  }
  const codeCategory = (code: string): string =>
    mappings[pairKey(code, codeType[code] ?? '')] || '';

  // Merge RS + filtered CS into one list per tenant (CS is superset, RS fills gaps)
  const allRows = (t: TG): FlatRow[] => [...t.rs, ...filteredCS(t.base._tenantIdx, t.cs)];

  // Parse rent roll date to sort value for range checks
  const rrDateNum = rentRollDate ? Date.parse(rentRollDate) : 0;

  // Check if a row is active on a given date (from <= date <= to)
  const isActiveOn = (r: FlatRow, dateNum: number): boolean => {
    if (!dateNum) return false;
    const f = dNum(r.from);
    const t = dNum(r.to);
    if (!f) return false;
    return f <= dateNum && (!t || t >= dateNum);
  };

  // ── Section A: Bumps per mapping category ──────────────────────────────────
  // For each non-Excluded category, compute bumps (date + PSF) like rent bumps.
  // "Rent" bumps come first, then other categories in order.
  const bumpCategories = categories //.filter(c => c !== 'Excluded');

  type BumpEntry = { dateStr: string; psf: number | null };
  // bumpsByTenant[catIdx][tenantIdx] = BumpEntry[]
  const bumpsByTenant: BumpEntry[][][] = [];
  const maxBumpsPerCat: number[] = [];

  for (let ci = 0; ci < bumpCategories.length; ci++) {
    const cat = bumpCategories[ci];
    const perTenant: BumpEntry[][] = [];
    let maxB = 0;
    for (const t of tenants) {
      const all = allRows(t);
      const catRows = all.filter(r => codeCategory(String(r.charge ?? '').trim()) === cat);
      // Collect unique from-dates
      const fromDates = new Map<number, string>();
      for (const r of catRows) {
        const sv = dNum(r.from);
        if (sv && !fromDates.has(sv)) fromDates.set(sv, toDateString(r.from) ?? '');
      }
      const sortedDates = [...fromDates.entries()].sort((a, b) => a[0] - b[0]);
      const area = toNumber(t.base.area);
      const bumps: BumpEntry[] = sortedDates.map(([sv, ds]) => {
        let totalAnnual = 0;
        for (const r of catRows) {
          const f = dNum(r.from);
          const to = dNum(r.to);
          if (f && f <= sv && (!to || to >= sv)) {
            totalAnnual += toNumber(r.annual) ?? 0;
          }
        }
        return { dateStr: ds, psf: area ? totalAnnual / area : null };
      });
      perTenant.push(bumps);
      maxB = Math.max(maxB, bumps.length);
    }
    bumpsByTenant.push(perTenant);
    maxBumpsPerCat.push(maxB);
  }

  // ── Section B: Current Charges (1 col per charge code, PSF on rent roll date) ─
  const allCodes: string[] = [];
  const allCodeSet = new Set<string>();
  for (const t of tenants) {
    for (const r of allRows(t)) {
      const c = String(r.charge ?? '').trim();
      if (c && !allCodeSet.has(c)) { allCodeSet.add(c); allCodes.push(c); }
    }
  }
  const MAP_ORD = ['Rent', 'Opex', 'Utility', 'Management', 'Insurance', 'Tax', 'Excluded'];
  const starCount = (c: string) => { const m = c.match(/^\*+/); return m ? m[0].length : 0; };
  allCodes.sort((a, b) => {
    // 1. Star count ascending (no stars first, then *, then **)
    const sa = starCount(a), sb = starCount(b);
    if (sa !== sb) return sa - sb;
    // 2. Mapping category order
    const ia = MAP_ORD.indexOf(codeCategory(a));
    const ib = MAP_ORD.indexOf(codeCategory(b));
    const oa = ia < 0 ? 999 : ia;
    const ob = ib < 0 ? 999 : ib;
    if (oa !== ob) return oa - ob;
    // 3. Alphabetical
    return a.localeCompare(b);
  });

  // ── Column layout ─────────────────────────────────────────────────────────
  // [Tenant Info] [TotalRentsCC] [VarInputRents]
  // [Annual Totals] [AnnualRec] [TotalRecCC] [VarInputRec]
  // [blank] [StepDate] [StepRent] [Var%] [Bumps per cat] [blank] [Current Charges]
  const atTotal = bumpCategories.length;
  const bumpTotalCols = bumpCategories.reduce((s, _, ci) => s + maxBumpsPerCat[ci] * 2, 0);
  const ccTotal = allCodes.length;
  const rentCatIdx = bumpCategories.indexOf('Rent');

  let col = nM;

  // After Security Deposit: 2 formula cols
  const COL_TRCC  = col++;  // Total Rents from Charge codes
  const COL_VRENT = col++;  // Var with Input Rents

  // Annual Totals (1 col per mapping category)
  const COL_AT = col; col += atTotal;

  // After Annual Totals: 3 formula cols
  const COL_AREC = col++;   // Annual Recovery
  const COL_TREC = col++;   // Total Rec from Charge codes
  const COL_VREC = col++;   // Var with Input Rec

  // Blank separator
  const COL_B1 = col++;

  // Step Date / Step Rent / Var %
  const COL_SD = col++;
  const COL_SR = col++;
  const COL_VP = col++;

  // Bump sections (one after another)
  const bumpStarts: number[] = [];
  for (let ci = 0; ci < bumpCategories.length; ci++) {
    bumpStarts.push(col);
    col += maxBumpsPerCat[ci] * 2;
  }

  // Blank separator
  const COL_B2 = bumpTotalCols > 0 && ccTotal > 0 ? col++ : col;

  // Current Charges (1 col per charge code)
  const COL_CC = col; col += ccTotal;
  const TOTAL = col;

  // Key column indices for formulas
  const annualRentColIdx = MAIN_KEYS.indexOf('annualRent');
  const areaColIdx = MAIN_KEYS.indexOf('area');

  // Step date = rent roll date + 365 days (as Excel serial number)
  const stepDateMs = rrDateNum ? rrDateNum + 365 * 24 * 60 * 60 * 1000 : 0;
  const stepDateSerial = stepDateMs ? Math.round((stepDateMs - EXCEL_EPOCH_UTC) / 86400000) : null;

  // AT columns for recovery (not Rent, not Excluded)
  const recAtCols = bumpCategories
    .map((cat, ci) => ({ cat, col: COL_AT + ci }))
    .filter(x => x.cat !== 'Rent' && x.cat !== 'Excluded')
    .map(x => x.col);

  // ── Build 4 header rows ───────────────────────────────────────────────────
  const mk = (): X[] => Array<X>(TOTAL).fill(null);
  const h1 = mk(); const h2 = mk(); const h3 = mk(); const h4 = mk();

  // Tenant Info section
  for (let i = 0; i < nM; i++) h1[i] = 'Tenant Info';
  for (let i = 0; i < nM; i++) h4[i] = MAIN_HDRS[i];
  h1[COL_TRCC] = 'Tenant Info'; h1[COL_VRENT] = 'Tenant Info';
  h4[COL_TRCC] = 'Total Rents from CC';
  h4[COL_VRENT] = 'Var with Input Rents';

  // Annual Totals headers
  if (atTotal > 0) {
    for (let i = COL_AT; i < COL_AT + atTotal; i++) h1[i] = 'Annual Totals';
  }
  for (let i = 0; i < bumpCategories.length; i++) {
    h2[COL_AT + i] = bumpCategories[i];
    h3[COL_AT + i] = 'Annual';
    h4[COL_AT + i] = bumpCategories[i];
  }
  h1[COL_AREC] = 'Annual Totals'; h1[COL_TREC] = 'Annual Totals'; h1[COL_VREC] = 'Annual Totals';
  h4[COL_AREC] = 'Annual Recovery';
  h4[COL_TREC] = 'Total Rec from CC';
  h4[COL_VREC] = 'Var with Input Rec';

  // Step Date / Step Rent / Var % headers
  h1[COL_SD] = 'Rent Bumps'; h1[COL_SR] = 'Rent Bumps'; h1[COL_VP] = 'Rent Bumps';
  if (stepDateSerial) h2[COL_SD] = stepDateSerial;
  h4[COL_SD] = 'Step Date';
  h4[COL_SR] = 'Step Rent';
  h4[COL_VP] = 'Var %';

  // Bump section headers
  for (let ci = 0; ci < bumpCategories.length; ci++) {
    const cat = bumpCategories[ci];
    const s = bumpStarts[ci];
    const w = maxBumpsPerCat[ci] * 2;
    for (let i = s; i < s + w; i++) h1[i] = `${cat} Bumps`;
    for (let p = 0; p < maxBumpsPerCat[ci]; p++) {
      const c = s + p * 2;
      h2[c] = `${cat} ${p + 1}`;       h2[c + 1] = `${cat} ${p + 1}`;
      h3[c] = 'Date';                   h3[c + 1] = 'PSF';
      h4[c] = `${cat} Date ${p + 1}`;   h4[c + 1] = `${cat} PSF ${p + 1}`;
    }
  }

  // Current Charges headers (1 col per charge code)
  if (ccTotal > 0) {
    for (let i = COL_CC; i < COL_CC + ccTotal; i++) h1[i] = 'Current Charges';
  }
  for (let i = 0; i < allCodes.length; i++) {
    const code = allCodes[i];
    const catLabel = codeCategory(code) || code;
    h2[COL_CC + i] = catLabel;
    h3[COL_CC + i] = code;
    h4[COL_CC + i] = 'Current';
  }

  // ── Formula column definitions ────────────────────────────────────────────
  // CC section column range (Excel letters)
  const ccFirstL = CL(COL_CC);
  const ccLastL  = ccTotal > 0 ? CL(COL_CC + ccTotal - 1) : ccFirstL;
  const annRentL = CL(annualRentColIdx);
  const areaL    = CL(areaColIdx);
  const sdL      = CL(COL_SD);

  // Rent bump section range
  const rbFirstL = rentCatIdx >= 0 && maxBumpsPerCat[rentCatIdx] > 0
    ? CL(bumpStarts[rentCatIdx]) : '';
  const rbLastL = rentCatIdx >= 0 && maxBumpsPerCat[rentCatIdx] > 0
    ? CL(bumpStarts[rentCatIdx] + maxBumpsPerCat[rentCatIdx] * 2 - 1) : '';

  type FormulaCol = { col: number; formula: (row: number) => string };
  const formulaCols: FormulaCol[] = [];

  if (ccTotal > 0) {
    // Total Rents from CC: =SUMIF(mapping_row,"Rent",data_row)
    formulaCols.push({
      col: COL_TRCC,
      formula: (r) => `_xlfn.SUMIF($${ccFirstL}$2:$${ccLastL}$2,"Rent",${ccFirstL}${r}:${ccLastL}${r})`,
    });
    // Var with Input Rents: = Annual Rent - Total Rents from CC
    formulaCols.push({
      col: COL_VRENT,
      formula: (r) => `${annRentL}${r}-${CL(COL_TRCC)}${r}`,
    });
    // Total Rec from CC: =SUMIFS(data,"<>Rent","<>Excluded")
    formulaCols.push({
      col: COL_TREC,
      formula: (r) => `_xlfn.SUMIFS(${ccFirstL}${r}:${ccLastL}${r},$${ccFirstL}$2:$${ccLastL}$2,"<>Rent",$${ccFirstL}$2:$${ccLastL}$2,"<>Excluded")`,
    });
  }

  // Annual Recovery: =SUM(recovery AT columns)
  if (recAtCols.length > 0) {
    formulaCols.push({
      col: COL_AREC,
      formula: (r) => recAtCols.map(c => `${CL(c)}${r}`).join('+'),
    });
  }

  // Var with Input Rec: = Total Rec from CC - Annual Recovery
  formulaCols.push({
    col: COL_VREC,
    formula: (r) => `${CL(COL_TREC)}${r}-${CL(COL_AREC)}${r}`,
  });

  // Step Date, Step Rent, Var % (only if Rent bumps exist)
  if (rbFirstL && rbLastL) {
    formulaCols.push({
      col: COL_SD,
      formula: (r) => `_xlfn.IF(_xlfn.OR(${rbFirstL}${r}="",${rbFirstL}${r}>$${sdL}$2),"",_xlfn.MAXIFS(${rbFirstL}${r}:${rbLastL}${r},${rbFirstL}${r}:${rbLastL}${r},"<="&$${sdL}$2))`,
    });
    formulaCols.push({
      col: COL_SR,
      formula: (r) => `_xlfn.IF(${sdL}${r}="","",OFFSET(${rbFirstL}${r},0,MATCH(${sdL}${r},${rbFirstL}${r}:${rbLastL}${r},0)))`,
    });
    formulaCols.push({
      col: COL_VP,
      formula: (r) => `_xlfn.IF(${CL(COL_SR)}${r}="","",_xlfn.IFERROR((${CL(COL_SR)}${r}-${annRentL}${r}/${areaL}${r})/(${annRentL}${r}/${areaL}${r}),0))`,
    });
  }

  // ── Build data rows ───────────────────────────────────────────────────────
  const dateMainCols = new Set(
    (['leaseFrom', 'leaseTo'] as (keyof FlatRow)[]).map(k => MAIN_KEYS.indexOf(k)).filter(i => i >= 0)
  );

  const dataRows: X[][] = tenants.map((t, ti) => {
    const { base } = t;
    const row = mk();

    // Tenant info
    const all = allRows(t);
    for (let i = 0; i < nM; i++) {
      const key = MAIN_KEYS[i];
      const v = base[key] as Cell;
      if (dateMainCols.has(i)) {
        row[i] = toExcelSerial(v);
      } else {
        row[i] = v instanceof Date ? toExcelSerial(v) : typeof v === 'number' ? v : (v as string | null) ?? null;
      }
    }

    // Override Annual Rent & Monthly Rent: sum from rent steps active on rent roll date
    if (rrDateNum) {
      let activeAnnual = 0;
      for (const r of all) {
        const code = String(r.charge ?? '').trim();
        if (codeCategory(code) === 'Rent' && isActiveOn(r, rrDateNum)) {
          activeAnnual += toNumber(r.annual) ?? 0;
        }
      }
      const annualIdx = MAIN_KEYS.indexOf('annualRent');
      const monthlyIdx = MAIN_KEYS.indexOf('monthlyRent');
      if (annualIdx >= 0) row[annualIdx] = activeAnnual;
      if (monthlyIdx >= 0) row[monthlyIdx] = activeAnnual / 12;
    }

    // Annual Totals: sum annual by category for rows active on rent roll date
    for (let ci = 0; ci < bumpCategories.length; ci++) {
      const cat = bumpCategories[ci];
      let sum = 0;
      for (const r of all) {
        const code = String(r.charge ?? '').trim();
        if (codeCategory(code) === cat && isActiveOn(r, rrDateNum)) {
          sum += toNumber(r.annual) ?? 0;
        }
      }
      row[COL_AT + ci] = sum;
    }

    // Bumps per category (write Excel serial numbers to avoid timezone issues)
    for (let ci = 0; ci < bumpCategories.length; ci++) {
      const bumps = bumpsByTenant[ci][ti];
      const s = bumpStarts[ci];
      for (let p = 0; p < bumps.length; p++) {
        row[s + p * 2] = bumps[p].dateStr ? toExcelSerial(bumps[p].dateStr) : null;
        row[s + p * 2 + 1] = bumps[p].psf;
      }
    }

    // Current Charges: annual amount of charge code row active on rent roll date
    for (let i = 0; i < allCodes.length; i++) {
      const code = allCodes[i];
      const activeRow = all.find(r =>
        String(r.charge ?? '').trim() === code && isActiveOn(r, rrDateNum)
      );
      row[COL_CC + i] = activeRow ? (toNumber(activeRow.annual) ?? 0) : 0;
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
  const blankColSet = new Set([COL_B1, COL_B2]);
  const isBlankCol = (i: number) => blankColSet.has(i);

  ws2.columns = Array.from({ length: TOTAL }, (_, i) => ({
    width: i === 0 ? 34
         : isBlankCol(i) ? 2
         : i < nM ? 16
         : 14,
  }));

  // ── Color palette (ARGB) ─────────────────────────────────────────────────
  const PAL = {
    ti1: 'FF1F3864', ti2: 'FF2E75B6', ti3: 'FF9DC3E6', ti4: 'FFDEEAF1',
    bp1: 'FF1E4620', bp2: 'FF548235', bp3: 'FFA9D18E', bp4: 'FFE2EFDA',
    cc1: 'FF833C00', cc2: 'FFC55A11', cc3: 'FFF4B183', cc4: 'FFFCE4D6',
    blank: 'FF303030',
    white: 'FFFFFFFF', dark: 'FF1A1A1A',
    rowOdd: 'FFFFFFFF', rowEven: 'FFF5F7FA',
    borderColor: 'FFB8C4CE',
  };

  const DARK_FILLS = new Set([PAL.ti1, PAL.ti2, PAL.bp1, PAL.bp2, PAL.cc1, PAL.cc2, PAL.blank]);

  const mkFill = (argb: string): ExcelJS.Fill =>
    ({ type: 'pattern', pattern: 'solid', fgColor: { argb } } as ExcelJS.Fill);

  const mkBorder = (weight: 'hair' | 'thin' | 'medium' = 'thin'): Partial<ExcelJS.Borders> => {
    const side = { style: weight as ExcelJS.BorderStyle, color: { argb: PAL.borderColor } };
    return { top: side, left: side, bottom: side, right: side };
  };

  const bumpRangeStart = bumpStarts.length > 0 ? bumpStarts[0] : -1;
  const bumpRangeEnd = bumpRangeStart >= 0 ? bumpRangeStart + bumpTotalCols : -1;

  // Classify columns into colour sections
  const tiExtras = new Set([COL_TRCC, COL_VRENT]);
  const atExtras = new Set([COL_AREC, COL_TREC, COL_VREC]);
  const stepCols = new Set([COL_SD, COL_SR, COL_VP]);

  const colSection = (ci: number): 'tenant' | 'at' | 'bp' | 'cc' | 'blank' => {
    if (ci < nM || tiExtras.has(ci)) return 'tenant';
    if (isBlankCol(ci)) return 'blank';
    if ((atTotal > 0 && ci >= COL_AT && ci < COL_AT + atTotal) || atExtras.has(ci)) return 'at';
    if (stepCols.has(ci)) return 'bp';
    if (bumpTotalCols > 0 && ci >= bumpRangeStart && ci < bumpRangeEnd) return 'bp';
    if (ccTotal > 0 && ci >= COL_CC && ci < COL_CC + ccTotal) return 'cc';
    return 'blank';
  };

  const hdrFill = (ci: number, level: 1 | 2 | 3 | 4): string => {
    const s = colSection(ci);
    if (s === 'blank') return PAL.blank;
    const map = {
      tenant: [PAL.ti1, PAL.ti2, PAL.ti3, PAL.ti4],
      at:     [PAL.cc1, PAL.cc2, PAL.cc3, PAL.cc4],
      bp:     [PAL.bp1, PAL.bp2, PAL.bp3, PAL.bp4],
      cc:     [PAL.cc1, PAL.cc2, PAL.cc3, PAL.cc4],
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

  // Build lookup of formula col → formula fn for quick access
  const formulaMap = new Map(formulaCols.map(fc => [fc.col, fc.formula]));

  dataRows.forEach((dataRow, ri) => {
    const exRow = ws2.addRow(dataRow as (string | number | Date | null)[]);
    const exRowNum = ri + 5; // 4 header rows + 1-based
    exRow.height = 15;
    const rowBg = ri % 2 === 0 ? PAL.rowOdd : PAL.rowEven;

    // Inject formulas for formula columns
    for (const fc of formulaCols) {
      const cell = exRow.getCell(fc.col + 1); // 1-based
      cell.value = { formula: fc.formula(exRowNum) } as ExcelJS.CellFormulaValue;
    }

    exRow.eachCell({ includeEmpty: true }, (cell, colIdx) => {
      const ci = colIdx - 1;
      cell.fill = mkFill(rowBg);
      cell.font = { size: 10, name: 'Calibri', color: { argb: PAL.dark } };
      cell.border = mkBorder('hair');

      const isFormula = formulaMap.has(ci);
      const isDateCol = dateMainCols.has(ci) || (ci >= nM && h4[ci] != null && String(h4[ci]).toLowerCase().includes('date'));
      const isVarPct = ci === COL_VP;

      if (isDateCol) {
        cell.numFmt = 'mm/dd/yyyy';
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (isVarPct) {
        cell.numFmt = '0.00%';
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (isFormula || typeof cell.value === 'number' || (cell.value && typeof cell.value === 'object' && 'formula' in cell.value)) {
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
        if (v instanceof Date) return toExcelSerial(v);
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
  rentRollDate: string;
  onRentRollDateChange: (date: string) => void;
}

export function TenancyScheduleTable({ tenants, fileName, onBack, rentRollDate, onRentRollDateChange }: Props) {
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

  // ── RS ⊆ CS validation: every charge code in Rent Steps must exist in Charge Schedules ──
  const rsCsMismatch = useMemo<Map<number, string[]>>(() => {
    const byTenant = new Map<number, { rsCodes: Set<string>; csCodes: Set<string> }>();
    for (const row of rows) {
      const i = row._tenantIdx;
      if (!byTenant.has(i)) byTenant.set(i, { rsCodes: new Set(), csCodes: new Set() });
      const g = byTenant.get(i)!;
      const sec = String(row._section ?? '').toLowerCase();
      const c = String(row.charge ?? '').trim();
      if (!c) continue;
      if (sec.includes('rent') && sec.includes('step')) g.rsCodes.add(c);
      else if (row._section) g.csCodes.add(c);
    }
    const warnings = new Map<number, string[]>();
    for (const [idx, { rsCodes, csCodes }] of byTenant.entries()) {
      if (rsCodes.size === 0) continue;
      const missing = [...rsCodes].filter(c => !csCodes.has(c));
      if (missing.length > 0) warnings.set(idx, missing);
    }
    return warnings;
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
          {rsCsMismatch.size > 0 && (
            <span
              className="text-[11px] font-mono text-amber-400 bg-amber-400/10 border border-amber-400/30 px-2 py-0.5 rounded"
              title={[...rsCsMismatch.entries()].map(([idx, missing]) => {
                const t = rows.find(r => r._tenantIdx === idx);
                const name = t ? fmt(t.lease) || fmt(t.unit) || `#${idx}` : `#${idx}`;
                return `${name}: ${missing.join(', ')}`;
              }).join('\n')}
            >
              ⚠ {rsCsMismatch.size} tenant{rsCsMismatch.size !== 1 ? 's' : ''} with RS/CS mismatch
            </span>
          )}
        </div>
        <div className="flex items-center gap-2">
          <label className="flex items-center gap-1.5 text-[11px] font-mono text-muted-foreground">
            Rent Roll Date
            <input
              type="date"
              value={rentRollDate}
              onChange={e => onRentRollDateChange(e.target.value)}
              className="px-2 py-1 text-[11px] font-mono rounded border border-panel-border bg-background text-foreground"
            />
          </label>
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
              const mismatchLines = isTenantBoundary ? rsCsMismatch.get(row._tenantIdx) : undefined;

              return (<>
                {mismatchLines && (
                  <tr key={`warn-${ri}`} className="bg-amber-400/5">
                    <td colSpan={COLS.length} className="px-2 py-1 text-[10px] font-mono text-amber-400 border border-amber-400/20">
                      ⚠ Charge codes in Rent Steps not found in Charge Schedules: {mismatchLines.join(', ')}
                    </td>
                  </tr>
                )}
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
              </>);
            })}
          </tbody>
        </table>
      </div>

      {showMapping && (
        <MappingDialog
          uniquePairs={uniquePairs}
          onClose={() => setShowMapping(false)}
          onExport={(mappings, cats) => {
            downloadXLSX(rows, fileName, mappings, cats, rentRollDate).then(() => setShowMapping(false));
          }}
        />
      )}
    </div>
  );
}