/**
 * Tenancy Schedule I — parser
 *
 * ═══════════════════════════════════════════════════════════════════════════════
 * STRUCTURE (derived from actual Excel file — Book2.xlsx)
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * PREAMBLE ROWS (skip):
 *   Row 1 — "Tenancy Schedule I"  (title)
 *   Row 2 — "Property: ... As of Date: ..."  (metadata)
 *
 * STACKED HEADER ROWS (merged vertically per column):
 *   Row 3 — top-level labels:   Property | Unit(s) | Lease | Lease Type | Area |
 *                                Lease From | Lease To | Term | Tenancy |
 *                                Monthly | Monthly | Annual | Annual |
 *                                Annual | Annual | Security | LOC Amount/ |
 *   Row 4 — sub-labels:         (empty)…| Years | Rent | Rent/Area | Rent |
 *                                Rent/Area | Rec./Area | Misc/Area | Deposit |
 *                                Bank Guarantee |
 *   Row 5 — sub-sub-labels:     (empty)…| Received |
 *
 *   Merged result per column (0-indexed):
 *     0  → "Property"
 *     1  → "Unit(s)"
 *     2  → "Lease"
 *     3  → "Lease Type"
 *     4  → "Area"
 *     5  → "Lease From"
 *     6  → "Lease To"
 *     7  → "Term"
 *     8  → "Tenancy Years"
 *     9  → "Monthly Rent"
 *     10 → "Monthly Rent/Area"
 *     11 → "Annual Rent"
 *     12 → "Annual Rent/Area"
 *     13 → "Annual Rec./Area"
 *     14 → "Annual Misc/Area"
 *     15 → "Security Deposit Received"
 *     16 → "LOC Amount/ Bank Guarantee"
 *
 * ROW 6 — blank separator (skipped)
 *
 * TENANT MAIN ROW (col A non-empty, e.g. "2091-2115 Faulkner Road (br40656)"):
 *   Col A  — Property name
 *   Col B  — Unit(s)
 *   Col C+ — values mapped to merged main headers
 *
 * BLANK ROW — skip (may appear after main row or between sub-sections)
 *
 * SUB-SECTION MARKER ROW (col A empty, col B has text like "Rent Steps"):
 *   Col B  — sub-section name  (e.g. "Rent Steps", "Charge Schedules")
 *   Col C+ — column labels for the rows that follow:
 *             Charge | Type | Unit | Area Label | Area | From | To |
 *             Monthly Amt | Amt/Area | Annual | Annual/Area |
 *             Management Fee | Annual Gross Amount
 *
 * SUB-SECTION DATA ROW (col A empty, col B empty, data in col C+):
 *   Values keyed by the column labels from the preceding marker row.
 *
 * TENANT BLOCK BOUNDARY:
 *   The ONLY trigger for a new tenant is col A becoming non-empty again.
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// src/lib/rent-roll-types/tenancy-schedule-parser.ts

// ─── Types ───────────────────────────────────────────────────────────────────

/** A single data row inside a sub-section, keyed by the column labels from the marker row */
export interface SubSectionRow {
  /** Column label → cell value (raw, preserving original type) */
  values: Record<string, string | number | Date | null>;
  /** The full raw Excel row */
  rawRow: (string | number | Date | null)[];
}

/** A sub-section block (e.g. "Rent Steps", "Charge Schedules") */
export interface SubSection {
  /** Label from col B of the marker row, e.g. "Rent Steps" */
  name: string;
  /** Column labels read from the marker row (col C onward, 0-indexed offset 2+) */
  columnLabels: string[];
  /** Data rows keyed by those labels */
  rows: SubSectionRow[];
}

export interface TenancyScheduleTenant {
  /** Main row values keyed by the merged main headers */
  mainRow: Record<string, string | number | Date | null>;
  /** All sub-sections in order of appearance */
  subSections: SubSection[];
  /** All raw rows belonging to this tenant (main + markers + data) */
  rawRows: (string | number | Date | null)[][];
}

// ─── Internal row kinds ───────────────────────────────────────────────────────

type RowKind =
  | { kind: 'skip' }
  | { kind: 'blank' }
  | { kind: 'tenant-main' }
  | { kind: 'sub-section-marker'; name: string }
  | { kind: 'sub-section-data' };

// ─── Helpers ──────────────────────────────────────────────────────────────────

type Cell = string | number | Date | null;
type Row = Cell[];

function cellStr(v: Cell): string {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function isBlank(row: Row | undefined): boolean {
  if (!row) return true;
  return row.every(c => c === null || c === undefined || String(c).trim() === '');
}

/**
 * Returns true if the cell value looks like it belongs in a header row:
 * a short text string, not a number, not a Date object.
 */
function isHeaderCell(v: Cell): boolean {
  if (v === null || v === undefined) return true; // empty is fine in a header row
  if (typeof v === 'number') return false;
  if (v instanceof Date) return false;
  const s = String(v).trim();
  // Long strings that look like property names / tenant names → data
  if (s.length > 40) return false;
  return true;
}

// ─── Header detection ─────────────────────────────────────────────────────────

/**
 * Detects the multi-row stacked header block.
 *
 * Algorithm:
 *  1. Scan the first 20 rows for the row whose col A (trimmed, lowercase)
 *     equals "property" — this is the first header row.
 *  2. Collect subsequent rows as header continuations while they pass
 *     `isHeaderRow()` (all non-empty cells are short text, no numbers/dates).
 *  3. Merge header rows vertically per column (join non-empty parts with space).
 *  4. Skip any trailing blank rows to find the true data start.
 *
 * Returns:
 *   labels       — merged label string per column index (empty string if none)
 *   headerEndRow — 0-based index of the first data row
 */
function detectMainHeaders(data: Row[]): {
  labels: string[];
  headerEndRow: number;
} {
  // Step 1 — find the anchor row (col A === "property")
  let anchorRow = -1;
  for (let i = 0; i < Math.min(20, data.length); i++) {
    const a = cellStr(data[i]?.[0]).toLowerCase();
    if (a === 'property' || a === 'properties') {
      anchorRow = i;
      break;
    }
  }
  if (anchorRow === -1) {
    return { labels: [], headerEndRow: 0 };
  }

  // Step 2 — collect header continuation rows
  const headerRows: Row[] = [data[anchorRow]];
  let headerEndRow = anchorRow + 1;

  for (let i = anchorRow + 1; i < Math.min(anchorRow + 8, data.length); i++) {
    const row = data[i];
    if (isBlank(row)) {
      // Blank row immediately after headers — consume and stop
      headerEndRow = i + 1;
      break;
    }
    // A row is a header continuation only if every non-empty cell passes isHeaderCell
    const nonEmpty = (row ?? []).filter(c => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonEmpty.length > 0 && nonEmpty.every(isHeaderCell)) {
      headerRows.push(row);
      headerEndRow = i + 1;
    } else {
      // Hit a data row — stop without consuming it
      break;
    }
  }

  // Step 3 — skip any remaining blank rows (filter/separator rows)
  while (headerEndRow < data.length && isBlank(data[headerEndRow])) {
    headerEndRow++;
  }

  // Step 4 — merge vertically per column
  const maxCols = headerRows.reduce((m, r) => Math.max(m, r?.length ?? 0), 0);
  const labels: string[] = [];
  for (let col = 0; col < maxCols; col++) {
    const parts: string[] = [];
    for (const hRow of headerRows) {
      const v = col < (hRow?.length ?? 0) ? cellStr(hRow[col]) : '';
      if (v) parts.push(v);
    }
    labels.push(parts.join(' ').trim());
  }

  return { labels, headerEndRow };
}

// ─── Row classification ───────────────────────────────────────────────────────

function classifyRow(row: Row | undefined, rowIndex: number, headerEndRow: number): RowKind {
  // Rows before data start are always skipped
  if (rowIndex < headerEndRow) return { kind: 'skip' };

  if (isBlank(row)) return { kind: 'blank' };

  const colA = cellStr(row![0]);
  const colB = cellStr(row![1]);

  // Tenant main row: col A is non-empty
  if (colA) {
    return { kind: 'tenant-main' };
  }

  // Sub-section marker: col A empty, col B has text
  // The same row also carries column labels starting at col C (index 2)
  if (!colA && colB) {
    return { kind: 'sub-section-marker', name: colB };
  }

  // Sub-section data row: col A and col B both empty, data somewhere from col C onward
  if (!colA && !colB) {
    const hasData = (row ?? []).slice(2).some(c => cellStr(c) !== '');
    if (hasData) return { kind: 'sub-section-data' };
  }

  return { kind: 'blank' };
}

// ─── Main exported parser ─────────────────────────────────────────────────────

export function parseTenancySchedule(
  data: Row[],
  addLog?: (type: 'system' | 'flag', msg: string) => void,
): TenancyScheduleTenant[] {
  const log = addLog ?? (() => {});

  // ── 1. Detect headers ──────────────────────────────────────────────────────
  const { labels: mainHeaders, headerEndRow } = detectMainHeaders(data);

  if (mainHeaders.length === 0) {
    log('flag', 'Tenancy Schedule: could not detect header rows. Is col A of the header row "Property"?');
    return [];
  }

  const namedHeaders = mainHeaders.filter(l => l);
  log('system', `Tenancy Schedule: ${namedHeaders.length} main columns detected, data starts at row ${headerEndRow + 1}.`);
  log('system', `Main headers: ${namedHeaders.join(' | ')}`);

  // ── 2. Parse tenant blocks ─────────────────────────────────────────────────
  const tenants: TenancyScheduleTenant[] = [];
  let current: TenancyScheduleTenant | null = null;
  let currentSubSection: SubSection | null = null;

  const finalise = () => {
    if (!current) return;
    if (currentSubSection) {
      current.subSections.push(currentSubSection);
      currentSubSection = null;
    }
    tenants.push(current);
    current = null;
  };

  for (let i = 0; i < data.length; i++) {
    const row = data[i] ?? [];
    const classified = classifyRow(row, i, headerEndRow);

    switch (classified.kind) {
      case 'skip':
      case 'blank':
        continue;

      // ── Tenant main row ────────────────────────────────────────────────────
      case 'tenant-main': {
        // Close the previous tenant (if any)
        finalise();

        // Map cell values to merged main headers
        const mainRow: Record<string, Cell> = {};
        for (let col = 0; col < mainHeaders.length; col++) {
          const label = mainHeaders[col];
          if (!label) continue;
          mainRow[label] = col < row.length ? (row[col] ?? null) : null;
        }

        current = { mainRow, subSections: [], rawRows: [row] };
        break;
      }

      // ── Sub-section marker row ─────────────────────────────────────────────
      case 'sub-section-marker': {
        if (!current) {
          log('flag', `Row ${i + 1}: sub-section marker "${classified.name}" found before any tenant row — skipped.`);
          continue;
        }

        // Close the previous sub-section
        if (currentSubSection) {
          current.subSections.push(currentSubSection);
        }

        // Column labels live in col C onward (index 2+) on the marker row itself
        const columnLabels: string[] = [];
        for (let col = 2; col < row.length; col++) {
          columnLabels.push(cellStr(row[col]));
        }
        // Trim trailing empty labels
        while (columnLabels.length > 0 && !columnLabels[columnLabels.length - 1]) {
          columnLabels.pop();
        }

        currentSubSection = {
          name: classified.name,
          columnLabels,
          rows: [],
        };
        current.rawRows.push(row);
        break;
      }

      // ── Sub-section data row ───────────────────────────────────────────────
      case 'sub-section-data': {
        if (!current) continue;
        current.rawRows.push(row);

        // If data arrives before any marker, open an unnamed sub-section
        if (!currentSubSection) {
          log('flag', `Row ${i + 1}: data row found before any sub-section marker — opening unnamed section.`);
          currentSubSection = { name: '(unnamed)', columnLabels: [], rows: [] };
        }

        // Map cell values to this sub-section's column labels
        // Labels start at col C (index 2); label index j → sheet column j + 2
        const values: Record<string, Cell> = {};
        for (let j = 0; j < currentSubSection.columnLabels.length; j++) {
          const label = currentSubSection.columnLabels[j];
          if (!label) continue;
          const colIdx = j + 2;
          values[label] = colIdx < row.length ? (row[colIdx] ?? null) : null;
        }

        currentSubSection.rows.push({ values, rawRow: row });
        break;
      }
    }
  }

  // Close the last open tenant
  finalise();

  // ── 3. Summary logs ────────────────────────────────────────────────────────
  log('system', `${tenants.length} tenant block(s) parsed.`);

  const subSectionNames = new Set<string>();
  for (const t of tenants) {
    for (const s of t.subSections) subSectionNames.add(s.name);
  }
  if (subSectionNames.size > 0) {
    log('system', `Sub-section types found: ${[...subSectionNames].join(', ')}.`);
  }

  return tenants;
}