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
  chargeCategory: string | null; // Rent, CAM, UTL, RET, Relief, Excluded
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

  // Annual charge breakdown by individual code columns (cols 29-56)
  annualChargesByCode: Record<string, number>;
  annualTotal: number | null; // col 80

  rawRows: Cell[][];
}

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

// ─── Column index map ────────────────────────────────────────────────────────

interface ColMap {
  unit: number;
  dba: number;
  leaseId: number;
  squareFootage: number;
  leaseType: number;
  unitType: number;
  leaseStatus: number;
  percentInLieu: number;
  commencementDate: number;
  originalEndDate: number;
  expireCloseDate: number;
  billCode: number;
  expenseDescription: number;
  beginDate: number;
  endDate: number;
  monthlyAmount: number;
  rateSF: number;
  chargeCategory: number; // col after rateSF (Rent/CAM/etc)
  total: number;
  // Future escalation columns
  futureBillCode: number;
  futureExpenseDesc: number;
  futureBeginDate: number;
  futureEndDate: number;
  futureMonthlyAmt: number;
  futureRateSF: number;
  futurePercentInc: number;
  // Category label column
  categoryLabel: number;
}

// Sorted longest-first so "unit type" matches before "unit", etc.
const HEADER_KEYWORDS: [string, keyof ColMap, boolean][] = [
  // [keyword, field, exactMatch]
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
  ['unit',                'unit',                true],   // exact only — avoid matching "unit type"
  ['dba',                 'dba',                 true],
  ['code',                'billCode',            true],
  ['amount',              'monthlyAmount',       true],
];

function findHeaderRow(data: Cell[][]): { headerRow: number; colMap: Partial<ColMap> } {
  for (let r = 0; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;

    // Build merged header by combining this row with the row above (multi-row headers).
    const prevRow = r > 0 ? data[r - 1] : null;
    const mergedHeaders: string[] = [];
    const maxCols = Math.max(row.length, prevRow?.length ?? 0);
    for (let c = 0; c < maxCols; c++) {
      const top = prevRow ? str(prevRow[c] ?? null).toLowerCase().replace(/:$/, '').trim() : '';
      const bot = str(row[c] ?? null).toLowerCase().replace(/:$/, '').trim();
      // Merge: "Begin" + "Date" → "begin date"
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

    // Need at least 5 matches to consider this the header row
    if (matchCount >= 5) {
      return { headerRow: r, colMap };
    }
  }

  // Fallback: use known default positions from JDE format
  return {
    headerRow: 10,
    colMap: {
      unit: 2, dba: 3, leaseId: 4, squareFootage: 7,
      leaseType: 8, unitType: 9, leaseStatus: 10, percentInLieu: 11,
      commencementDate: 12, originalEndDate: 13, expireCloseDate: 14,
      billCode: 15, expenseDescription: 16, beginDate: 17, endDate: 18,
      monthlyAmount: 19, rateSF: 20, chargeCategory: 21,
      total: 80,
    },
  };
}

/** Detect future escalation columns and category column from section divider row */
function findSectionColumns(data: Cell[][], headerRow: number): {
  futureBillCode: number;
  futureExpenseDesc: number;
  futureBeginDate: number;
  futureEndDate: number;
  futureMonthlyAmt: number;
  futureRateSF: number;
  futurePercentInc: number;
  categoryLabel: number;
} {
  // Look for "Future Rent" section divider in the rows before the header
  // The future section typically starts around col 84
  // Also scan the header row itself for matching labels in the future section

  const result = {
    futureBillCode: 84,
    futureExpenseDesc: 85,
    futureBeginDate: 86,
    futureEndDate: 87,
    futureMonthlyAmt: 88,
    futureRateSF: 89,
    futurePercentInc: 90,
    categoryLabel: 277,
  };

  // Scan divider rows for "Future Rent" to find the start column
  for (let r = Math.max(0, headerRow - 5); r < headerRow; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 50; c < row.length; c++) {
      const val = str(row[c]).toLowerCase();
      if (val.includes('future rent') || val.includes('future rent & expense')) {
        // Found the start of future section; bill code should be at or near this col
        result.futureBillCode = c;
        result.futureExpenseDesc = c + 1;
        result.futureBeginDate = c + 2;
        result.futureEndDate = c + 3;
        result.futureMonthlyAmt = c + 4;
        result.futureRateSF = c + 5;
        result.futurePercentInc = c + 6;
        break;
      }
    }
  }

  // Scan header row for the future section labels to refine positions
  const hrow = data[headerRow];
  if (hrow) {
    // Look for Bill Code / Code in future section area (col 60+)
    for (let c = 60; c < hrow.length; c++) {
      const val = str(hrow[c]).toLowerCase();
      if ((val === 'bill code' || val === 'code') && c > 50) {
        result.futureBillCode = c;
        break;
      }
    }
  }

  // Find category label column - typically the rightmost column with category names
  // Check the first data row after headers for a column containing category labels
  for (let r = headerRow + 1; r < Math.min(headerRow + 5, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = row.length - 1; c >= 200; c--) {
      const val = str(row[c]).toLowerCase();
      if (val === 'anchor' || val === 'outparcel' || val === 'specialty') {
        result.categoryLabel = c;
        break;
      }
    }
    if (result.categoryLabel !== 277) break;
  }

  return result;
}

/** Build a map of annual charge code columns from row 9 (sub-headers) */
function findChargeCodeColumns(data: Cell[][], headerRow: number): { col: number; code: string }[] {
  // The charge code names appear in the row above the header row (or 2 rows above)
  const codes: { col: number; code: string }[] = [];
  // Check 1-2 rows above header
  for (let offset = 1; offset <= 2; offset++) {
    const r = headerRow - offset;
    if (r < 0) continue;
    const row = data[r];
    if (!row) continue;
    for (let c = 28; c < 80 && c < row.length; c++) {
      const val = str(row[c]).toUpperCase();
      if (val && /^[A-Z]{3,5}$/.test(val)) {
        codes.push({ col: c, code: val });
      }
    }
    if (codes.length > 0) break;
  }
  return codes;
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

function extractMetadata(
  row: Cell[],
  tenant: MallRentRollTenant,
  labelCol: number,
  valueCol: number,
): boolean {
  const rawLabel = str(cell(row, labelCol));
  if (!rawLabel.includes(':')) return false;

  const label = rawLabel.replace(':', '').trim().toLowerCase();
  const field = METADATA_LABELS[label];
  if (!field) return false;

  const value = cell(row, valueCol);
  (tenant as Record<string, Cell>)[field] = value;
  return true;
}

// ─── Main parser ─────────────────────────────────────────────────────────────

export function parseMallRentRoll(
  data: Cell[][],
  addLog?: LogFn,
): MallRentRollTenant[] {
  const log = addLog || (() => {});
  const tenants: MallRentRollTenant[] = [];

  // 1. Find header row and column map
  const { headerRow, colMap } = findHeaderRow(data);
  log('system', `Header row detected at row ${headerRow + 1}`);

  // 2. Find section columns (future escalations, category)
  const sectionCols = findSectionColumns(data, headerRow);
  const fullColMap: ColMap = {
    unit: colMap.unit ?? 2,
    dba: colMap.dba ?? 3,
    leaseId: colMap.leaseId ?? 4,
    squareFootage: colMap.squareFootage ?? 7,
    leaseType: colMap.leaseType ?? 8,
    unitType: colMap.unitType ?? 9,
    leaseStatus: colMap.leaseStatus ?? 10,
    percentInLieu: colMap.percentInLieu ?? 11,
    commencementDate: colMap.commencementDate ?? 12,
    originalEndDate: colMap.originalEndDate ?? 13,
    expireCloseDate: colMap.expireCloseDate ?? 14,
    billCode: colMap.billCode ?? 15,
    expenseDescription: colMap.expenseDescription ?? 16,
    beginDate: colMap.beginDate ?? 17,
    endDate: colMap.endDate ?? 18,
    monthlyAmount: colMap.monthlyAmount ?? 19,
    rateSF: colMap.rateSF ?? 20,
    chargeCategory: colMap.chargeCategory ?? 21,
    total: colMap.total ?? 80,
    ...sectionCols,
  };

  // 3. Find annual charge code columns
  const chargeCodeCols = findChargeCodeColumns(data, headerRow);

  // 4. Extract property metadata from top rows
  let propertyName = '';
  for (let r = 0; r < headerRow; r++) {
    const row = data[r];
    if (!row) continue;
    for (const c of row) {
      const v = str(c);
      if (v.toLowerCase().includes('mall') || v.toLowerCase().includes('property')) {
        // Check if this looks like a property name (not a header keyword)
        if (v.length > 10 && !v.includes('---')) {
          propertyName = v;
          break;
        }
      }
    }
  }

  // 5. Walk data rows
  let currentTenant: MallRentRollTenant | null = null;
  let currentCategory = '';
  const dataStart = headerRow + 1;

  // Sometimes there's a category row right after header (or mixed in)
  // Category rows have col 277 (or categoryLabel) with text like "Anchor"

  for (let r = dataStart; r < data.length; r++) {
    const row = data[r];
    if (!row || row.length === 0) continue;

    const unitVal = str(cell(row, fullColMap.unit));
    const dbaVal = str(cell(row, fullColMap.dba));
    const billCodeVal = str(cell(row, fullColMap.billCode));
    const leaseIdVal = str(cell(row, fullColMap.leaseId));
    const col0 = num(cell(row, 0));

    // Track category from the category label column
    const catVal = str(cell(row, fullColMap.categoryLabel));
    if (catVal && catVal !== currentCategory) {
      currentCategory = catVal;
    }

    // Check for "Total :" row
    const expDescVal = str(cell(row, fullColMap.expenseDescription));
    if (expDescVal.toLowerCase().includes('total')) {
      if (currentTenant) {
        currentTenant.totalMonthlyAmount = num(cell(row, fullColMap.monthlyAmount));
        currentTenant.rawRows.push([...row]);
      }
      continue;
    }

    // New tenant: Unit and DBA both present (col0 === 0 is also a reliable indicator)
    if (unitVal && dbaVal && (col0 === 0 || col0 === null)) {
      // Finalize previous tenant
      if (currentTenant) {
        tenants.push(currentTenant);
      }

      currentTenant = {
        unit: unitVal,
        dba: dbaVal,
        leaseId: leaseIdVal,
        squareFootage: num(cell(row, fullColMap.squareFootage)),
        category: currentCategory,

        leaseType: str(cell(row, fullColMap.leaseType)) || null,
        unitType: str(cell(row, fullColMap.unitType)) || null,
        leaseStatus: str(cell(row, fullColMap.leaseStatus)) || null,
        percentInLieu: cell(row, fullColMap.percentInLieu),
        commencementDate: cell(row, fullColMap.commencementDate),
        openDate: null,
        originalEndDate: cell(row, fullColMap.originalEndDate),
        expireCloseDate: cell(row, fullColMap.expireCloseDate),

        charges: [],
        totalMonthlyAmount: null,
        futureEscalations: [],
        annualChargesByCode: {},
        annualTotal: num(cell(row, fullColMap.total)),
        rawRows: [[...row]],
      };

      // Extract first charge line from main row
      if (billCodeVal) {
        currentTenant.charges.push({
          billCode: billCodeVal,
          expenseDescription: str(cell(row, fullColMap.expenseDescription)),
          beginDate: cell(row, fullColMap.beginDate),
          endDate: cell(row, fullColMap.endDate),
          monthlyAmount: num(cell(row, fullColMap.monthlyAmount)),
          annualRateSF: num(cell(row, fullColMap.rateSF)),
          chargeCategory: str(cell(row, fullColMap.chargeCategory)) || null,
        });
      }

      // Extract future escalation from main row
      const futureBill = str(cell(row, fullColMap.futureBillCode));
      if (futureBill) {
        currentTenant.futureEscalations.push({
          billCode: futureBill,
          expenseDescription: str(cell(row, fullColMap.futureExpenseDesc)),
          beginDate: cell(row, fullColMap.futureBeginDate),
          endDate: cell(row, fullColMap.futureEndDate),
          monthlyAmount: num(cell(row, fullColMap.futureMonthlyAmt)),
          annualRateSF: num(cell(row, fullColMap.futureRateSF)),
          percentInc: num(cell(row, fullColMap.futurePercentInc)),
        });
      }

      // Extract annual charges by code
      for (const { col, code } of chargeCodeCols) {
        const val = num(cell(row, col));
        if (val !== null) {
          currentTenant.annualChargesByCode[code] = val;
        }
      }

      continue;
    }

    // No current tenant — skip
    if (!currentTenant) continue;

    // Additional charge line: bill code present, no unit/DBA
    if (billCodeVal && !unitVal) {
      currentTenant.charges.push({
        billCode: billCodeVal,
        expenseDescription: str(cell(row, fullColMap.expenseDescription)),
        beginDate: cell(row, fullColMap.beginDate),
        endDate: cell(row, fullColMap.endDate),
        monthlyAmount: num(cell(row, fullColMap.monthlyAmount)),
        annualRateSF: num(cell(row, fullColMap.rateSF)),
        chargeCategory: str(cell(row, fullColMap.chargeCategory)) || null,
      });

      // Also check for future escalation on this row
      const futureBill = str(cell(row, fullColMap.futureBillCode));
      if (futureBill) {
        currentTenant.futureEscalations.push({
          billCode: futureBill,
          expenseDescription: str(cell(row, fullColMap.futureExpenseDesc)),
          beginDate: cell(row, fullColMap.futureBeginDate),
          endDate: cell(row, fullColMap.futureEndDate),
          monthlyAmount: num(cell(row, fullColMap.futureMonthlyAmt)),
          annualRateSF: num(cell(row, fullColMap.futureRateSF)),
          percentInc: num(cell(row, fullColMap.futurePercentInc)),
        });
      }

      currentTenant.rawRows.push([...row]);
      continue;
    }

    // Metadata row: col 4 (leaseId column) contains a label with ":"
    if (extractMetadata(row, currentTenant, fullColMap.leaseId, fullColMap.leaseId + 2)) {
      currentTenant.rawRows.push([...row]);
      continue;
    }

    // Separator row (col0 === 1 or fully empty) — just skip
    currentTenant.rawRows.push([...row]);
  }

  // Finalize last tenant
  if (currentTenant) {
    tenants.push(currentTenant);
  }

  log('system', `Parsed ${tenants.length} tenants from Mall Rent Roll.`);

  // Log category breakdown
  const categories = new Map<string, number>();
  for (const t of tenants) {
    categories.set(t.category, (categories.get(t.category) || 0) + 1);
  }
  const catSummary = [...categories.entries()].map(([k, v]) => `${k || 'Uncategorized'}: ${v}`).join(', ');
  log('system', `Categories: ${catSummary}`);

  return tenants;
}
