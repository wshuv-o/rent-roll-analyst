// src/lib/rent-roll-types/mall-rent-roll-parser.ts
//
// Deterministic parser for JDE EnterpriseOne Mall Rent Roll exports.
// Supports both wide XLSX format (280+ cols) and narrow CSV format (~35 cols).

type Cell = string | number | Date | null;
type LogFn = (type: string, message: string) => void;

// ─── Exported types ──────────────────────────────────────────────────────────

export interface MallChargeLine {
  billCode: string;
  expenseDescription: string;
  beginDate: Cell;
  endDate: Cell;
  monthlyAmount: number | null;
  annualRateSF: number | null;
  chargeCategory: string | null;
}

export interface MallFutureEscalation {
  billCode: string;
  expenseDescription: string;
  beginDate: Cell;
  endDate: Cell;
  monthlyAmount: number | null;
  annualRateSF: number | null;
  percentInc: number | null;
}

export interface BumpPair { date: Cell; amount: Cell; percent?: Cell; }
export interface BreakpointEntry { date: Cell; amount: Cell; percent: Cell; }

export interface OverageEntry {
  billCode: string;
  beginDate: Cell;
  endDate: Cell;
  breakpoint: number | null;
  percent: number | null;
}

export interface EscalationSummary {
  code: Cell;
  description: Cell;
  beginDate: Cell;
  endDate: Cell;
  monthlyAmount: number | null;
  rateSF: number | null;
  count: number | null;
}

export interface MallRentRollTenant {
  unit: string;
  dba: string;
  leaseId: string;
  squareFootage: number | null;
  category: string;

  leaseType: string | null;
  unitType: string | null;
  leaseStatus: string | null;
  percentInLieu: Cell;
  commencementDate: Cell;
  openDate: Cell;
  originalEndDate: Cell;
  expireCloseDate: Cell;

  charges: MallChargeLine[];
  totalMonthlyAmount: number | null;

  futureEscalations: MallFutureEscalation[];
  overageEntries: OverageEntry[];

  annualChargesByCode: Record<string, number>;
  annualTotal: number | null;
  variance: number | null;

  rentBumps: BumpPair[];
  breakpoints: BreakpointEntry[];

  camEscalation: EscalationSummary | null;
  utlEscalation: EscalationSummary | null;
  retEscalation: EscalationSummary | null;

  camBumps: BumpPair[];
  utlBumps: BumpPair[];
  retBumps: BumpPair[];

  buildingCode: Cell;
  buildingName: Cell;

  rawRows: Cell[][];
}

// ─── Default charge code mapping ─────────────────────────────────────────────

export interface ChargeCodeMapping {
  code: string;
  description: string;
  category: string;
  reliefSubType: string;
}

export const DEFAULT_CHARGE_CODE_MAPPING: ChargeCodeMapping[] = [
  { code: 'BMRP', description: 'MINIMUM RENT',         category: 'Rent',     reliefSubType: '' },
  { code: 'BMGP', description: 'GROSS RENT',           category: 'Rent',     reliefSubType: '' },
  { code: 'BMRB', description: 'MINIMUM RENT',         category: 'Rent',     reliefSubType: '' },
  { code: 'BMRE', description: 'MINIMUM RENT',         category: 'Rent',     reliefSubType: '' },
  { code: 'CAFD', description: 'CAM FIXED',            category: 'CAM',      reliefSubType: '' },
  { code: 'INPD', description: 'INSURANCE-PRORATA',    category: 'CAM',      reliefSubType: '' },
  { code: 'CAID', description: 'CAM INDOOR PRORATA',   category: 'CAM',      reliefSubType: '' },
  { code: 'CAOD', description: 'CAM OUTDOOR PRORATA',  category: 'CAM',      reliefSubType: '' },
  { code: 'MKFP', description: 'MKT FUND',             category: 'CAM',      reliefSubType: '' },
  { code: 'CAOO', description: 'CAM OUTDOOR PRORATA',  category: 'CAM',      reliefSubType: '' },
  { code: 'CATP', description: 'CAM TOE',              category: 'CAM',      reliefSubType: '' },
  { code: 'CATB', description: 'CAM TOE',              category: 'CAM',      reliefSubType: '' },
  { code: 'FCFP', description: 'FIXED CAM FOOD CT',    category: 'CAM',      reliefSubType: '' },
  { code: 'WTSP', description: 'WATER & SEWER',        category: 'UTL',      reliefSubType: '' },
  { code: 'UTPP', description: 'UTILITIES',             category: 'UTL',      reliefSubType: '' },
  { code: 'REPP', description: 'RET PRORATA',          category: 'RET',      reliefSubType: '' },
  { code: 'REPB', description: 'RET PRORATA',          category: 'RET',      reliefSubType: '' },
  { code: 'RRRT', description: 'RELIEF-MIN RENT',      category: 'Relief',   reliefSubType: 'Rent' },
  { code: 'RRGR', description: 'RELIEF-GROSS BILLED',  category: 'Relief',   reliefSubType: 'Rent' },
  { code: 'RRTS', description: 'RELIEF-RET PRO-RATA',  category: 'Relief',   reliefSubType: 'RET' },
  { code: 'SPBM', description: 'SP RENT-12+',          category: 'Excluded', reliefSubType: '' },
  { code: 'NRCT', description: 'NOTES REC-TENANT',     category: 'Excluded', reliefSubType: '' },
  { code: 'DEFR', description: 'DEFERRED RENT',        category: 'Excluded', reliefSubType: '' },
  { code: 'SPST', description: 'SP-STORAGE',           category: 'Excluded', reliefSubType: '' },
  { code: 'BSTR', description: 'STORAGE RENT',         category: 'Excluded', reliefSubType: '' },
  { code: 'BMAN', description: 'MIN RENT ANTENNA',     category: 'Excluded', reliefSubType: '' },
  { code: 'BANT', description: 'RENT-TELECOM',         category: 'Excluded', reliefSubType: '' },
  { code: 'PADC', description: 'TRASH PAD RENTAL',     category: 'Excluded', reliefSubType: '' },
];

export const CHARGE_CODES = DEFAULT_CHARGE_CODE_MAPPING.map(m => m.code);
const MAPPING_BY_CODE = new Map(DEFAULT_CHARGE_CODE_MAPPING.map(m => [m.code, m]));

// ─── Helpers ─────────────────────────────────────────────────────────────────

function str(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-US');
  return String(v).trim();
}

function num(v: Cell): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = parseFloat(v.replace(/,/g, '').replace(/%$/, '').trim());
    return isNaN(n) ? null : n;
  }
  return null;
}

function cell(row: Cell[], idx: number): Cell {
  if (idx < 0 || idx >= row.length) return null;
  return row[idx] ?? null;
}

function extractBumps(row: Cell[], startCol: number, count: number): BumpPair[] {
  const bumps: BumpPair[] = [];
  for (let i = 0; i < count; i++) {
    bumps.push({ date: cell(row, startCol + i * 2), amount: cell(row, startCol + i * 2 + 1) });
  }
  return bumps;
}

function extractBreakpoints(row: Cell[], startCol: number): BreakpointEntry[] {
  const bps: BreakpointEntry[] = [];
  bps.push({ date: cell(row, startCol), amount: cell(row, startCol + 1), percent: cell(row, startCol + 2) });
  for (let i = 0; i < 5; i++) {
    const base = startCol + 4 + i * 3;
    bps.push({ date: cell(row, base), amount: cell(row, base + 1), percent: cell(row, base + 2) });
  }
  return bps;
}

function extractEscSummary(row: Cell[], startCol: number): EscalationSummary | null {
  const code = cell(row, startCol + 1);
  if (!code && !cell(row, startCol + 4)) return null;
  return {
    code: cell(row, startCol + 1), description: cell(row, startCol + 2),
    beginDate: cell(row, startCol + 3), endDate: cell(row, startCol + 4),
    monthlyAmount: num(cell(row, startCol + 5)), rateSF: num(cell(row, startCol + 6)),
    count: num(cell(row, startCol + 7)),
  };
}

// ─── Column index map ────────────────────────────────────────────────────────

interface ColMap {
  unit: number; dba: number; leaseId: number; squareFootage: number;
  leaseType: number; unitType: number; leaseStatus: number; percentInLieu: number;
  commencementDate: number; originalEndDate: number; expireCloseDate: number;
  billCode: number; expenseDescription: number; beginDate: number; endDate: number;
  monthlyAmount: number; rateSF: number; chargeCategory: number; total: number;
  variance: number;
  futureBillCode: number; futureExpenseDesc: number; futureBeginDate: number;
  futureEndDate: number; futureMonthlyAmt: number; futureRateSF: number; futurePercentInc: number;
  overageBillCode: number; overageBeginDate: number; overageEndDate: number;
  overageBreakpoint: number; overagePercent: number;
  rentBumpStart: number; breakpointStart: number;
  camEscStart: number; utlEscStart: number; retEscStart: number;
  camBumpStart: number; utlBumpStart: number; retBumpStart: number;
  categoryLabel: number; buildingCode: number; buildingName: number;
}

const HEADER_KEYWORDS: [string, keyof ColMap, boolean][] = [
  ['expense description', 'expenseDescription', false],
  ['commencement date',   'commencementDate',   false],
  ['original end date',   'originalEndDate',     false],
  ['expire/close date',   'expireCloseDate',     false],
  ['square footage',      'squareFootage',       false],
  ['monthly amount',      'monthlyAmount',       false],
  ['lease status',        'leaseStatus',         false],
  ['lease type',          'leaseType',           false],
  ['unit type',           'unitType',            false],
  ['begin date',          'beginDate',           false],
  ['end date',            'endDate',             false],
  ['bill code',           'billCode',            false],
  ['% in lieu',           'percentInLieu',       false],
  ['lease id',            'leaseId',             false],
  ['rate/sf',             'rateSF',              false],
  ['footage',             'squareFootage',       false],
  ['unit',                'unit',                true],
  ['dba',                 'dba',                 true],
  ['code',                'billCode',            true],
  ['amount',              'monthlyAmount',       true],
];

function buildMergedHeaders(data: Cell[][], headerRow: number): string[] {
  const row = data[headerRow];
  const prevRow = headerRow > 0 ? data[headerRow - 1] : null;
  const maxCols = Math.max(row?.length ?? 0, prevRow?.length ?? 0);
  const merged: string[] = [];
  for (let c = 0; c < maxCols; c++) {
    const top = prevRow ? str(prevRow[c] ?? null).toLowerCase().replace(/:$/, '').trim() : '';
    const bot = str(row?.[c] ?? null).toLowerCase().replace(/:$/, '').trim();
    merged[c] = top && bot ? `${top} ${bot}` : (bot || top);
  }
  return merged;
}

function findHeaderRow(data: Cell[][]): { headerRow: number; colMap: Partial<ColMap> } {
  for (let r = 0; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;

    const mergedHeaders = buildMergedHeaders(data, r);
    let matchCount = 0;
    const colMap: Partial<ColMap> = {};
    const assigned = new Set<number>();

    for (let c = 0; c < mergedHeaders.length; c++) {
      const val = mergedHeaders[c];
      if (!val) continue;
      for (const [keyword, field, exact] of HEADER_KEYWORDS) {
        if (assigned.has(c)) break;
        const matches = exact ? val === keyword : val.includes(keyword);
        if (matches && colMap[field] === undefined) {
          colMap[field] = c;
          assigned.add(c);
          matchCount++;
          break;
        }
      }
      if (val === 'total' && colMap.total === undefined) colMap.total = c;
    }

    if (matchCount >= 5) return { headerRow: r, colMap };
  }

  return {
    headerRow: 10,
    colMap: {
      unit: 2, dba: 3, leaseId: 4, squareFootage: 7,
      leaseType: 8, unitType: 9, leaseStatus: 10, percentInLieu: 11,
      commencementDate: 12, originalEndDate: 13, expireCloseDate: 14,
      billCode: 15, expenseDescription: 16, beginDate: 17, endDate: 18,
      monthlyAmount: 19, rateSF: 20, chargeCategory: 21, total: 80,
    },
  };
}

// ─── Section detection (for narrow CSV format) ──────────────────────────────

function detectSections(data: Cell[][], headerRow: number, mainBillCodeCol: number): {
  isNarrow: boolean;
  futureBillCode: number; futureExpenseDesc: number; futureBeginDate: number;
  futureEndDate: number; futureMonthlyAmt: number; futureRateSF: number; futurePercentInc: number;
  overageBillCode: number; overageBeginDate: number; overageEndDate: number;
  overageBreakpoint: number; overagePercent: number;
} {
  const defaults = {
    isNarrow: false,
    futureBillCode: 84, futureExpenseDesc: 85, futureBeginDate: 86,
    futureEndDate: 87, futureMonthlyAmt: 88, futureRateSF: 89, futurePercentInc: 90,
    overageBillCode: -1, overageBeginDate: -1, overageEndDate: -1,
    overageBreakpoint: -1, overagePercent: -1,
  };

  // Strategy 1: find all "code"/"bill code" columns in merged headers
  const mergedHeaders = buildMergedHeaders(data, headerRow);
  const codeColumns: number[] = [];
  for (let c = 0; c < mergedHeaders.length; c++) {
    const val = mergedHeaders[c];
    if (val === 'code' || val === 'bill code' || val.includes('bill code')) {
      codeColumns.push(c);
    }
  }

  // Strategy 2: find section labels in rows above header
  let futureStart = -1, overageStart = -1;
  for (let r = Math.max(0, headerRow - 5); r < headerRow; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = str(row[c]).toLowerCase().replace(/-/g, ' ').trim();
      if (!val) continue;
      if (val.includes('future rent') || val.includes('expense escalation')) {
        if (futureStart < 0) futureStart = c;
      } else if (val.includes('overage') || val.includes('in lieu rent')) {
        if (overageStart < 0) overageStart = c;
      }
    }
  }

  // If we found multiple code columns, this is narrow format
  if (codeColumns.length >= 2) {
    const fc = codeColumns[1];
    const result = {
      isNarrow: true,
      futureBillCode: fc,
      futureExpenseDesc: fc + 1,
      futureBeginDate: fc + 2,
      futureEndDate: fc + 3,
      futureMonthlyAmt: fc + 4,
      futureRateSF: fc + 5,
      futurePercentInc: fc + 6,
      overageBillCode: -1, overageBeginDate: -1, overageEndDate: -1,
      overageBreakpoint: -1, overagePercent: -1,
    };

    if (codeColumns.length >= 3) {
      const oc = codeColumns[2];
      result.overageBillCode = oc;
      result.overageBeginDate = oc + 1;
      result.overageEndDate = oc + 2;
      // Find breakpoint column near overage section
      for (let c = oc + 1; c < oc + 6 && c < mergedHeaders.length; c++) {
        if (mergedHeaders[c].includes('breakpoint')) {
          result.overageBreakpoint = c;
          result.overagePercent = c + 1;
          break;
        }
      }
      if (result.overageBreakpoint < 0) {
        result.overageBreakpoint = oc + 3;
        result.overagePercent = oc + 4;
      }
    }
    return result;
  }

  // If section labels found but code columns not duplicated, use section starts
  if (futureStart > 0) {
    const fc = futureStart;
    defaults.isNarrow = true;
    defaults.futureBillCode = fc;
    defaults.futureExpenseDesc = fc + 1;
    defaults.futureBeginDate = fc + 2;
    defaults.futureEndDate = fc + 3;
    defaults.futureMonthlyAmt = fc + 4;
    defaults.futureRateSF = fc + 5;
    defaults.futurePercentInc = fc + 6;

    if (overageStart > 0) {
      defaults.overageBillCode = overageStart;
      defaults.overageBeginDate = overageStart + 1;
      defaults.overageEndDate = overageStart + 2;
      defaults.overageBreakpoint = overageStart + 3;
      defaults.overagePercent = overageStart + 4;
    }
  }

  return defaults;
}

// ─── Charge code column detection ───────────────────────────────────────────

function findChargeCodeColumns(data: Cell[][], headerRow: number): { col: number; code: string }[] {
  const codes: { col: number; code: string }[] = [];
  const knownCodes = new Set(CHARGE_CODES);
  for (let offset = 1; offset <= 2; offset++) {
    const r = headerRow - offset;
    if (r < 0) continue;
    const row = data[r];
    if (!row) continue;
    for (let c = 28; c < 80 && c < row.length; c++) {
      const val = str(row[c]).toUpperCase();
      if (val && knownCodes.has(val)) codes.push({ col: c, code: val });
    }
    if (codes.length > 0) break;
  }
  if (codes.length > 0) return codes;

  const descToCodes = new Map<string, string[]>();
  for (const m of DEFAULT_CHARGE_CODE_MAPPING) {
    const key = m.description.toUpperCase();
    const arr = descToCodes.get(key) || [];
    arr.push(m.code);
    descToCodes.set(key, arr);
  }
  const usageIdx = new Map<string, number>();
  for (let offset = 1; offset <= 3; offset++) {
    const r = headerRow - offset;
    if (r < 0) continue;
    const row = data[r];
    if (!row) continue;
    let found = 0;
    for (let c = 28; c < 80 && c < row.length; c++) {
      const val = str(row[c]).toUpperCase();
      if (!val) continue;
      const candidates = descToCodes.get(val);
      if (candidates) {
        const idx = usageIdx.get(val) || 0;
        if (idx < candidates.length) {
          codes.push({ col: c, code: candidates[idx] });
          usageIdx.set(val, idx + 1);
          found++;
        }
      }
    }
    if (found >= 5) break;
  }
  if (codes.length > 0) return codes;

  return CHARGE_CODES.map((code, i) => ({ col: 29 + i, code }));
}

// ─── Metadata extraction ─────────────────────────────────────────────────────

const METADATA_LABELS: Record<string, keyof Pick<MallRentRollTenant,
  'leaseType' | 'unitType' | 'leaseStatus' | 'percentInLieu' |
  'commencementDate' | 'openDate' | 'originalEndDate' | 'expireCloseDate'
>> = {
  'lease type': 'leaseType',
  'unit type': 'unitType',
  'lease status': 'leaseStatus',
  '% in lieu': 'percentInLieu',
  'commencement date': 'commencementDate',
  'open date': 'openDate',
  'original end date': 'originalEndDate',
  'expire/close date': 'expireCloseDate',
};

function extractMetadata(row: Cell[], tenant: MallRentRollTenant, labelCol: number, valueCol: number): boolean {
  const rawLabel = str(cell(row, labelCol));
  if (!rawLabel.includes(':')) return false;
  const label = rawLabel.replace(':', '').trim().toLowerCase();
  const field = METADATA_LABELS[label];
  if (!field) return false;
  (tenant as unknown as Record<string, Cell>)[field] = cell(row, valueCol);
  return true;
}

// ─── Overage extraction helper ──────────────────────────────────────────────

function tryExtractOverage(row: Cell[], C: ColMap, tenant: MallRentRollTenant) {
  if (C.overageBillCode < 0) return;
  const overageBill = str(cell(row, C.overageBillCode));
  if (overageBill) {
    tenant.overageEntries.push({
      billCode: overageBill,
      beginDate: cell(row, C.overageBeginDate),
      endDate: cell(row, C.overageEndDate),
      breakpoint: num(cell(row, C.overageBreakpoint)),
      percent: num(cell(row, C.overagePercent)),
    });
  }
}

// ─── Post-processing: derive bumps/breakpoints ─────────────────────────────

function deriveFromEscalationsAndOverage(tenants: MallRentRollTenant[], isNarrow: boolean) {
  if (!isNarrow) return; // Wide format already extracted bumps from columns

  for (const t of tenants) {
    // Derive breakpoints from overage entries
    t.breakpoints = t.overageEntries.map(oe => ({
      date: oe.beginDate,
      amount: oe.breakpoint,
      percent: oe.percent,
    }));

    // Derive rent bumps from future escalations with Rent category
    const rentFutures = t.futureEscalations.filter(fe => {
      const m = MAPPING_BY_CODE.get(fe.billCode);
      return m?.category === 'Rent';
    });
    t.rentBumps = rentFutures.map(fe => ({
      date: fe.beginDate,
      amount: fe.annualRateSF, // Rate/SF for rent bumps
      percent: fe.percentInc,
    }));

    // Derive CAM bumps from future escalations with CAM category
    const camFutures = t.futureEscalations.filter(fe => {
      const m = MAPPING_BY_CODE.get(fe.billCode);
      return m?.category === 'CAM';
    });
    t.camBumps = camFutures.map(fe => ({
      date: fe.beginDate,
      amount: fe.monthlyAmount !== null ? fe.monthlyAmount * 12 : null,
      percent: fe.percentInc,
    }));

    // Derive UTL bumps
    const utlFutures = t.futureEscalations.filter(fe => {
      const m = MAPPING_BY_CODE.get(fe.billCode);
      return m?.category === 'UTL';
    });
    t.utlBumps = utlFutures.map(fe => ({
      date: fe.beginDate,
      amount: fe.monthlyAmount !== null ? fe.monthlyAmount * 12 : null,
      percent: fe.percentInc,
    }));

    // Derive RET bumps
    const retFutures = t.futureEscalations.filter(fe => {
      const m = MAPPING_BY_CODE.get(fe.billCode);
      return m?.category === 'RET';
    });
    t.retBumps = retFutures.map(fe => ({
      date: fe.beginDate,
      amount: fe.monthlyAmount !== null ? fe.monthlyAmount * 12 : null,
      percent: fe.percentInc,
    }));
  }
}

// ─── Main parser ─────────────────────────────────────────────────────────────

export function parseMallRentRoll(data: Cell[][], addLog?: LogFn): MallRentRollTenant[] {
  const log = addLog || (() => {});
  const tenants: MallRentRollTenant[] = [];

  const { headerRow, colMap } = findHeaderRow(data);
  log('system', `Header row detected at row ${headerRow + 1}`);

  // Detect sections (narrow CSV vs wide XLSX)
  const sections = detectSections(data, headerRow, colMap.billCode ?? 7);
  log('system', `Format: ${sections.isNarrow ? 'narrow CSV' : 'wide XLSX'}, future bill code col: ${sections.futureBillCode}, overage bill code col: ${sections.overageBillCode}`);

  const C: ColMap = {
    unit: colMap.unit ?? 0, dba: colMap.dba ?? 1, leaseId: colMap.leaseId ?? 2,
    squareFootage: colMap.squareFootage ?? 5,
    leaseType: colMap.leaseType ?? 8, unitType: colMap.unitType ?? 9,
    leaseStatus: colMap.leaseStatus ?? 10, percentInLieu: colMap.percentInLieu ?? 11,
    commencementDate: colMap.commencementDate ?? 12, originalEndDate: colMap.originalEndDate ?? 13,
    expireCloseDate: colMap.expireCloseDate ?? 14,
    billCode: colMap.billCode ?? 7, expenseDescription: colMap.expenseDescription ?? 8,
    beginDate: colMap.beginDate ?? 9, endDate: colMap.endDate ?? 10,
    monthlyAmount: colMap.monthlyAmount ?? 11, rateSF: colMap.rateSF ?? 12,
    chargeCategory: colMap.chargeCategory ?? 21, total: colMap.total ?? 80,
    variance: 81,
    futureBillCode: sections.futureBillCode,
    futureExpenseDesc: sections.futureExpenseDesc,
    futureBeginDate: sections.futureBeginDate,
    futureEndDate: sections.futureEndDate,
    futureMonthlyAmt: sections.futureMonthlyAmt,
    futureRateSF: sections.futureRateSF,
    futurePercentInc: sections.futurePercentInc,
    overageBillCode: sections.overageBillCode,
    overageBeginDate: sections.overageBeginDate,
    overageEndDate: sections.overageEndDate,
    overageBreakpoint: sections.overageBreakpoint,
    overagePercent: sections.overagePercent,
    rentBumpStart: 93, breakpointStart: 145,
    camEscStart: 177, utlEscStart: 185, retEscStart: 193,
    camBumpStart: 202, utlBumpStart: 227, retBumpStart: 252,
    categoryLabel: 277, buildingCode: 285, buildingName: 286,
  };

  // Detect category label column
  for (let r = headerRow + 1; r < Math.min(headerRow + 5, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = row.length - 1; c >= 200; c--) {
      const val = str(row[c]).toLowerCase();
      if (val === 'anchor' || val === 'outparcel' || val === 'specialty' || val === 'inline') {
        C.categoryLabel = c;
        break;
      }
    }
    if (C.categoryLabel !== 277) break;
  }

  const chargeCodeCols = sections.isNarrow ? [] : findChargeCodeColumns(data, headerRow);
  if (!sections.isNarrow) {
    log('system', `Found ${chargeCodeCols.length} charge code columns`);
  }

  let currentTenant: MallRentRollTenant | null = null;
  let currentCategory = '';

  for (let r = headerRow + 1; r < data.length; r++) {
    const row = data[r];
    if (!row || row.length === 0) continue;

    const unitVal = str(cell(row, C.unit));
    const dbaVal = str(cell(row, C.dba));
    const billCodeVal = str(cell(row, C.billCode));

    const catVal = str(cell(row, C.categoryLabel));
    if (catVal && catVal !== currentCategory) currentCategory = catVal;

    // Space type row: only unit column has a value
    if (unitVal && !dbaVal && !billCodeVal) {
      const otherVals = row.filter((v, ci) => ci !== C.unit && v !== null && v !== undefined && String(v).trim() !== '');
      if (otherVals.length === 0) {
        currentCategory = unitVal;
        log('system', `Space type detected: "${unitVal}"`);
        continue;
      }
    }

    // Total row
    const expDescVal = str(cell(row, C.expenseDescription));
    if (expDescVal.toLowerCase().includes('total')) {
      if (currentTenant) {
        currentTenant.totalMonthlyAmount = num(cell(row, C.monthlyAmount));
        currentTenant.rawRows.push([...row]);
      }
      continue;
    }

    // New tenant: unit + dba present
    if (unitVal && dbaVal) {
      if (currentTenant) tenants.push(currentTenant);

      currentTenant = {
        unit: unitVal, dba: dbaVal, leaseId: str(cell(row, C.leaseId)),
        squareFootage: num(cell(row, C.squareFootage)),
        category: currentCategory,
        leaseType: str(cell(row, C.leaseType)) || null,
        unitType: str(cell(row, C.unitType)) || null,
        leaseStatus: str(cell(row, C.leaseStatus)) || null,
        percentInLieu: cell(row, C.percentInLieu),
        commencementDate: cell(row, C.commencementDate),
        openDate: null,
        originalEndDate: cell(row, C.originalEndDate),
        expireCloseDate: cell(row, C.expireCloseDate),
        charges: [], totalMonthlyAmount: null,
        futureEscalations: [],
        overageEntries: [],
        annualChargesByCode: {},
        annualTotal: num(cell(row, C.total)),
        variance: num(cell(row, C.variance)),
        // For wide format, extract from columns; for narrow, derive later
        rentBumps: sections.isNarrow ? [] : extractBumps(row, C.rentBumpStart, 18),
        breakpoints: sections.isNarrow ? [] : extractBreakpoints(row, C.breakpointStart),
        camEscalation: sections.isNarrow ? null : extractEscSummary(row, C.camEscStart),
        utlEscalation: sections.isNarrow ? null : extractEscSummary(row, C.utlEscStart),
        retEscalation: sections.isNarrow ? null : extractEscSummary(row, C.retEscStart),
        camBumps: sections.isNarrow ? [] : extractBumps(row, C.camBumpStart, 12),
        utlBumps: sections.isNarrow ? [] : extractBumps(row, C.utlBumpStart, 12),
        retBumps: sections.isNarrow ? [] : extractBumps(row, C.retBumpStart, 12),
        buildingCode: cell(row, C.buildingCode),
        buildingName: cell(row, C.buildingName),
        rawRows: [[...row]],
      };

      // First charge line
      if (billCodeVal) {
        currentTenant.charges.push({
          billCode: billCodeVal, expenseDescription: str(cell(row, C.expenseDescription)),
          beginDate: cell(row, C.beginDate), endDate: cell(row, C.endDate),
          monthlyAmount: num(cell(row, C.monthlyAmount)), annualRateSF: num(cell(row, C.rateSF)),
          chargeCategory: str(cell(row, C.chargeCategory)) || null,
        });
      }

      // Future escalation
      const futureBill = str(cell(row, C.futureBillCode));
      if (futureBill) {
        currentTenant.futureEscalations.push({
          billCode: futureBill, expenseDescription: str(cell(row, C.futureExpenseDesc)),
          beginDate: cell(row, C.futureBeginDate), endDate: cell(row, C.futureEndDate),
          monthlyAmount: num(cell(row, C.futureMonthlyAmt)), annualRateSF: num(cell(row, C.futureRateSF)),
          percentInc: num(cell(row, C.futurePercentInc)),
        });
      }

      // Overage/breakpoint
      tryExtractOverage(row, C, currentTenant);

      // Annual charges by code (wide format only)
      for (const { col, code } of chargeCodeCols) {
        const val = num(cell(row, col));
        if (val !== null) currentTenant.annualChargesByCode[code] = val;
      }
      continue;
    }

    if (!currentTenant) continue;

    // Additional charge line
    if (billCodeVal && !unitVal) {
      currentTenant.charges.push({
        billCode: billCodeVal, expenseDescription: str(cell(row, C.expenseDescription)),
        beginDate: cell(row, C.beginDate), endDate: cell(row, C.endDate),
        monthlyAmount: num(cell(row, C.monthlyAmount)), annualRateSF: num(cell(row, C.rateSF)),
        chargeCategory: str(cell(row, C.chargeCategory)) || null,
      });
      const futureBill = str(cell(row, C.futureBillCode));
      if (futureBill) {
        currentTenant.futureEscalations.push({
          billCode: futureBill, expenseDescription: str(cell(row, C.futureExpenseDesc)),
          beginDate: cell(row, C.futureBeginDate), endDate: cell(row, C.futureEndDate),
          monthlyAmount: num(cell(row, C.futureMonthlyAmt)), annualRateSF: num(cell(row, C.futureRateSF)),
          percentInc: num(cell(row, C.futurePercentInc)),
        });
      }
      tryExtractOverage(row, C, currentTenant);
      currentTenant.rawRows.push([...row]);
      continue;
    }

    // Metadata row — also check for future escalation & overage on same row
    const metadataFound = extractMetadata(row, currentTenant, C.leaseId, C.leaseId + 2);

    // Even on metadata/other rows, check for future escalations and overage
    const futureBill = str(cell(row, C.futureBillCode));
    if (futureBill) {
      currentTenant.futureEscalations.push({
        billCode: futureBill, expenseDescription: str(cell(row, C.futureExpenseDesc)),
        beginDate: cell(row, C.futureBeginDate), endDate: cell(row, C.futureEndDate),
        monthlyAmount: num(cell(row, C.futureMonthlyAmt)), annualRateSF: num(cell(row, C.futureRateSF)),
        percentInc: num(cell(row, C.futurePercentInc)),
      });
    }
    tryExtractOverage(row, C, currentTenant);
    currentTenant.rawRows.push([...row]);
  }

  if (currentTenant) tenants.push(currentTenant);

  // Post-process: derive bumps/breakpoints from future escalations and overage
  deriveFromEscalationsAndOverage(tenants, sections.isNarrow);

  log('system', `Parsed ${tenants.length} tenants from Mall Rent Roll.`);
  const categories = new Map<string, number>();
  for (const t of tenants) categories.set(t.category, (categories.get(t.category) || 0) + 1);
  log('system', `Categories: ${[...categories.entries()].map(([k, v]) => `${k || 'Uncategorized'}: ${v}`).join(', ')}`);

  // Log sample future escalations and breakpoints for debugging
  const withFuture = tenants.filter(t => t.futureEscalations.length > 0);
  log('system', `Tenants with future escalations: ${withFuture.length}`);
  const withBreakpoints = tenants.filter(t => t.breakpoints.length > 0);
  log('system', `Tenants with breakpoints: ${withBreakpoints.length}`);

  return tenants;
}
