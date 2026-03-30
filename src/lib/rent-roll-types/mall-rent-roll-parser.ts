// src/lib/rent-roll-types/mall-rent-roll-parser.ts
//
// Deterministic parser for JDE EnterpriseOne Mall Rent Roll exports.
// Each tenant occupies a multi-row block: main row, charge lines, metadata, total, separator.
// This parser collapses each block into a single MallRentRollTenant.

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

export interface BumpPair { date: Cell; amount: Cell; }
export interface BreakpointEntry { date: Cell; amount: Cell; percent: Cell; }

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

  annualChargesByCode: Record<string, number>;
  annualTotal: number | null;
  variance: number | null;

  // Rent bumps (18 pairs)
  rentBumps: BumpPair[];
  // Breakpoints (current + 5 future)
  breakpoints: BreakpointEntry[];

  // CAM/UTL/RET escalation summaries
  camEscalation: EscalationSummary | null;
  utlEscalation: EscalationSummary | null;
  retEscalation: EscalationSummary | null;

  // CAM/UTL/RET bumps (12 pairs each)
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
  category: string;      // Rent, CAM, UTL, RET, Relief, Excluded
  reliefSubType: string; // for Relief codes: Rent, RET, etc.
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

/** Ordered list of charge codes (same order as columns in source) */
export const CHARGE_CODES = DEFAULT_CHARGE_CODE_MAPPING.map(m => m.code);

// ─── Helpers ─────────────────────────────────────────────────────────────────

function str(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('en-US');
  return String(v).trim();
}

function num(v: Cell): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = parseFloat(v.replace(/,/g, '').trim());
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
    bumps.push({
      date: cell(row, startCol + i * 2),
      amount: cell(row, startCol + i * 2 + 1),
    });
  }
  return bumps;
}

function extractBreakpoints(row: Cell[], startCol: number): BreakpointEntry[] {
  const bps: BreakpointEntry[] = [];
  // Current breakpoint: cols 145, 146, 147
  bps.push({ date: cell(row, startCol), amount: cell(row, startCol + 1), percent: cell(row, startCol + 2) });
  // 5 future breakpoints: cols 149-163 in groups of 3 (skip col 148)
  for (let i = 0; i < 5; i++) {
    const base = startCol + 4 + i * 3; // 149, 152, 155, 158, 161
    bps.push({ date: cell(row, base), amount: cell(row, base + 1), percent: cell(row, base + 2) });
  }
  return bps;
}

function extractEscSummary(row: Cell[], startCol: number): EscalationSummary | null {
  const code = cell(row, startCol + 1); // code is offset +1 from section label
  if (!code && !cell(row, startCol + 4)) return null;
  return {
    code: cell(row, startCol + 1),
    description: cell(row, startCol + 2),
    beginDate: cell(row, startCol + 3),
    endDate: cell(row, startCol + 4),
    monthlyAmount: num(cell(row, startCol + 5)),
    rateSF: num(cell(row, startCol + 6)),
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

function findHeaderRow(data: Cell[][]): { headerRow: number; colMap: Partial<ColMap> } {
  for (let r = 0; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;

    const prevRow = r > 0 ? data[r - 1] : null;
    const mergedHeaders: string[] = [];
    const maxCols = Math.max(row.length, prevRow?.length ?? 0);
    for (let c = 0; c < maxCols; c++) {
      const top = prevRow ? str(prevRow[c] ?? null).toLowerCase().replace(/:$/, '').trim() : '';
      const bot = str(row[c] ?? null).toLowerCase().replace(/:$/, '').trim();
      mergedHeaders[c] = top && bot ? `${top} ${bot}` : (bot || top);
    }

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

function findChargeCodeColumns(data: Cell[][], headerRow: number): { col: number; code: string }[] {
  const codes: { col: number; code: string }[] = [];

  // Strategy 1: Look for known charge codes (BMRP, CAFD, etc.) in rows above header
  const knownCodes = new Set(CHARGE_CODES);
  for (let offset = 1; offset <= 2; offset++) {
    const r = headerRow - offset;
    if (r < 0) continue;
    const row = data[r];
    if (!row) continue;
    for (let c = 28; c < 80 && c < row.length; c++) {
      const val = str(row[c]).toUpperCase();
      if (val && knownCodes.has(val)) {
        codes.push({ col: c, code: val });
      }
    }
    if (codes.length > 0) break;
  }
  if (codes.length > 0) return codes;

  // Strategy 2: Match descriptions from rows above header to DEFAULT_CHARGE_CODE_MAPPING
  // JDE format has descriptions like "MINIMUM RENT", "CAM FIXED" in the sub-header row
  const descMap = new Map<string, string>();
  // Build a lookup: description → code. For duplicate descriptions, track which ones are used.
  const descToCodes = new Map<string, string[]>();
  for (const m of DEFAULT_CHARGE_CODE_MAPPING) {
    const key = m.description.toUpperCase();
    const arr = descToCodes.get(key) || [];
    arr.push(m.code);
    descToCodes.set(key, arr);
  }
  // Track usage index per description for duplicates
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
    if (found >= 5) break; // found enough
  }
  if (codes.length > 0) return codes;

  // Strategy 3: Fallback — assume 28 codes start at col 29 in standard order
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

// ─── Main parser ─────────────────────────────────────────────────────────────

export function parseMallRentRoll(data: Cell[][], addLog?: LogFn): MallRentRollTenant[] {
  const log = addLog || (() => {});
  const tenants: MallRentRollTenant[] = [];

  const { headerRow, colMap } = findHeaderRow(data);
  log('system', `Header row detected at row ${headerRow + 1}`);

  // Build full column map with defaults for all sections
  const C: ColMap = {
    unit: colMap.unit ?? 2, dba: colMap.dba ?? 3, leaseId: colMap.leaseId ?? 4,
    squareFootage: colMap.squareFootage ?? 7,
    leaseType: colMap.leaseType ?? 8, unitType: colMap.unitType ?? 9,
    leaseStatus: colMap.leaseStatus ?? 10, percentInLieu: colMap.percentInLieu ?? 11,
    commencementDate: colMap.commencementDate ?? 12, originalEndDate: colMap.originalEndDate ?? 13,
    expireCloseDate: colMap.expireCloseDate ?? 14,
    billCode: colMap.billCode ?? 15, expenseDescription: colMap.expenseDescription ?? 16,
    beginDate: colMap.beginDate ?? 17, endDate: colMap.endDate ?? 18,
    monthlyAmount: colMap.monthlyAmount ?? 19, rateSF: colMap.rateSF ?? 20,
    chargeCategory: colMap.chargeCategory ?? 21, total: colMap.total ?? 80,
    variance: 81,
    futureBillCode: 84, futureExpenseDesc: 85, futureBeginDate: 86,
    futureEndDate: 87, futureMonthlyAmt: 88, futureRateSF: 89, futurePercentInc: 90,
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

  const chargeCodeCols = findChargeCodeColumns(data, headerRow);
  log('system', `Found ${chargeCodeCols.length} charge code columns: ${chargeCodeCols.slice(0, 5).map(c => `${c.code}@${c.col}`).join(', ')}${chargeCodeCols.length > 5 ? '...' : ''}`);

  let currentTenant: MallRentRollTenant | null = null;
  let currentCategory = '';

  for (let r = headerRow + 1; r < data.length; r++) {
    const row = data[r];
    if (!row || row.length === 0) continue;

    const unitVal = str(cell(row, C.unit));
    const dbaVal = str(cell(row, C.dba));
    const billCodeVal = str(cell(row, C.billCode));
    const col0 = num(cell(row, 0));

    const catVal = str(cell(row, C.categoryLabel));
    if (catVal && catVal !== currentCategory) currentCategory = catVal;

    // Space type row: only unit column has a value (e.g. "Anchor", "Inline", "Outparcel")
    // These rows have no DBA, no bill code, no sqft — just a label in the unit column
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

    // New tenant
    if (unitVal && dbaVal && (col0 === 0 || col0 === null)) {
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
        annualChargesByCode: {},
        annualTotal: num(cell(row, C.total)),
        variance: num(cell(row, C.variance)),
        // Rent bumps (18 pairs starting at col 93)
        rentBumps: extractBumps(row, C.rentBumpStart, 18),
        // Breakpoints (current + 5 future starting at col 145)
        breakpoints: extractBreakpoints(row, C.breakpointStart),
        // CAM/UTL/RET escalation summaries
        camEscalation: extractEscSummary(row, C.camEscStart),
        utlEscalation: extractEscSummary(row, C.utlEscStart),
        retEscalation: extractEscSummary(row, C.retEscStart),
        // CAM/UTL/RET bumps (12 pairs each)
        camBumps: extractBumps(row, C.camBumpStart, 12),
        utlBumps: extractBumps(row, C.utlBumpStart, 12),
        retBumps: extractBumps(row, C.retBumpStart, 12),
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

      // Annual charges by code
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
      currentTenant.rawRows.push([...row]);
      continue;
    }

    // Metadata row
    if (extractMetadata(row, currentTenant, C.leaseId, C.leaseId + 2)) {
      currentTenant.rawRows.push([...row]);
      continue;
    }

    currentTenant.rawRows.push([...row]);
  }

  if (currentTenant) tenants.push(currentTenant);

  log('system', `Parsed ${tenants.length} tenants from Mall Rent Roll.`);
  const categories = new Map<string, number>();
  for (const t of tenants) categories.set(t.category, (categories.get(t.category) || 0) + 1);
  log('system', `Categories: ${[...categories.entries()].map(([k, v]) => `${k || 'Uncategorized'}: ${v}`).join(', ')}`);

  return tenants;
}
