/**
 * Tenancy Schedule I — parser (extraction policy).
 *
 * ═══════════════════════════════════════════════════════════════════════════════
 * STRUCTURE RULES
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * HEADER:
 *   - Multi-row stacked headers at the top of the sheet.
 *     e.g. row 3: "Annual"  row 4: "Misc/Area"  → merged label "Annual Misc/Area"
 *   - We detect these dynamically by finding consecutive text-heavy rows
 *     before the first data row. We merge vertically per column.
 *
 * TENANT MAIN ROW:
 *   - Column A (0) has the property name → signals a new tenant.
 *   - All values on this row map to the merged main headers.
 *
 * BLANK ROWS:
 *   - May appear after the main row or between sub-sections. Just skip.
 *
 * SUB-SECTION MARKER ROW:
 *   - Col A (0) is empty, Col B (1) has a label like "Rent Steps",
 *     "Charge Schedules", or any other sub-header text.
 *   - The SAME row also contains column labels from col C (2) onward.
 *     e.g. "Charge | Type | Unit | Area Label | Area | From | To | ..."
 *   - We read these labels dynamically — they become the keys for data rows
 *     in that sub-section.
 *
 * SUB-SECTION DATA ROW:
 *   - Col A empty, Col B empty, data in col C+ onward.
 *   - Values are keyed by the column labels read from the marker row above.
 *
 * TENANT BLOCK BOUNDARY:
 *   - Next row with col A non-empty starts a new tenant.
 *   - No blank-row-based splitting — only col A drives tenant boundaries.
 *
 * ═══════════════════════════════════════════════════════════════════════════════
 */
//src/lib/rent-roll-types/tenancy-schedule-parser.ts
import { getCellValue } from '../col-utils';

// ─── Types ──────────────────────────────────────────────────────────────────

/** A single data row inside a sub-section, keyed by the column labels from the marker row */
export interface SubSectionRow {
  /** Column label → cell value (raw, preserving type) */
  values: Record<string, string | number | null>;
  /** The full raw Excel row */
  rawRow: (string | number | null)[];
}

/** A sub-section block (e.g. "Rent Steps", "Charge Schedules") */
export interface SubSection {
  /** The label from col B, e.g. "Rent Steps" */
  name: string;
  /** Column labels read from the marker row (col C onward) */
  columnLabels: string[];
  /** Data rows keyed by those labels */
  rows: SubSectionRow[];
}

export interface TenancyScheduleTenant {
  /** Main row values keyed by the merged main headers */
  mainRow: Record<string, string | number | null>;
  /** All sub-sections in order of appearance */
  subSections: SubSection[];
  /** All raw rows belonging to this tenant (main + marker + data) */
  rawRows: (string | number | null)[][];
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function isBlankRow(row: (string | number | null)[] | undefined): boolean {
  if (!row) return true;
  return row.every(c => c === null || c === undefined || String(c).trim() === '');
}

function cellStr(val: string | number | null): string {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

// ─── Header detection ───────────────────────────────────────────────────────

/**
 * Detect the multi-row stacked main headers.
 *
 * Strategy:
 *   1. Find the first row that looks like a header — multiple text-only cells,
 *      and starts with something like "Property" in col A.
 *   2. Collect consecutive rows that are also header-like (text cells that
 *      fill gaps below the first header row — the "stacked" sub-labels).
 *   3. Merge vertically: for each column, join non-empty cells with a space.
 *
 * Returns: { labels: string[] (merged label per column), headerEndRow: number }
 */
function detectMainHeaders(data: (string | number | null)[][]): {
  labels: string[];
  headerEndRow: number;
} {
  // Find the first header row: scan first 10 rows for one where col 0 has "Property"
  let headerStartRow = -1;
  for (let i = 0; i < Math.min(15, data.length); i++) {
    const row = data[i];
    if (!row) continue;
    const col0 = cellStr(row[0]).toLowerCase();
    if (col0 === 'property' || col0 === 'properties') {
      headerStartRow = i;
      break;
    }
  }

  if (headerStartRow === -1) {
    // Fallback: use the first row that has >=4 non-empty text cells
    for (let i = 0; i < Math.min(15, data.length); i++) {
      const row = data[i];
      if (!row) continue;
      const textCells = row.filter(c => {
        const s = cellStr(c);
        return s !== '' && isNaN(Number(s));
      });
      if (textCells.length >= 4) {
        headerStartRow = i;
        break;
      }
    }
  }

  if (headerStartRow === -1) {
    return { labels: [], headerEndRow: 0 };
  }

  // Determine max columns
  const maxCols = data.slice(0, headerStartRow + 5).reduce((m, r) => Math.max(m, r?.length ?? 0), 0);

  // Collect header rows: starting from headerStartRow, keep going while the row
  // is text-heavy and NOT a data row (col 0 doesn't look like a property address)
  const headerRows: (string | number | null)[][] = [data[headerStartRow]];
  let headerEndRow = headerStartRow + 1;

  for (let i = headerStartRow + 1; i < Math.min(headerStartRow + 5, data.length); i++) {
    const row = data[i];
    if (!row || isBlankRow(row)) {
      // A blank row after headers could be a filter row — skip it and stop
      headerEndRow = i + 1;
      break;
    }

    const col0 = cellStr(row[0]);

    // If col 0 has a long string (looks like a property name), this is data — stop
    if (col0.length > 20) break;
    // If col 0 is "Property" again, skip (duplicate)
    if (col0.toLowerCase() === 'property') continue;

    // Check if this row looks like a continuation header (sub-labels):
    // mostly short text cells, many cells empty (filling gaps from row above)
    const filled = row.filter(c => cellStr(c) !== '');
    const allShortText = filled.every(c => {
      const s = cellStr(c);
      return s.length < 30 && isNaN(Number(s.replace(/[,$%]/g, '')));
    });

    if (allShortText && filled.length >= 1) {
      // Even a single-cell row is a valid header continuation
      // (e.g. "Received" stacked under "Security" + "Deposit")
      headerRows.push(row);
      headerEndRow = i + 1;
    } else {
      break;
    }
  }

  // Skip any blank/filter rows after the last header row
  while (headerEndRow < data.length && isBlankRow(data[headerEndRow])) {
    headerEndRow++;
  }

  // Merge headers vertically: for each column, join non-empty cells top-to-bottom
  const labels: string[] = [];
  for (let col = 0; col < maxCols; col++) {
    const parts: string[] = [];
    for (const hRow of headerRows) {
      const val = col < hRow.length ? cellStr(hRow[col]) : '';
      if (val) parts.push(val);
    }
    labels.push(parts.join(' ').trim());
  }

  return { labels, headerEndRow };
}

// ─── Row classification ─────────────────────────────────────────────────────

type RowKind = 'tenant-main' | 'sub-section-marker' | 'sub-section-data' | 'blank' | 'skip';

function classifyRow(
  row: (string | number | null)[] | undefined,
  headerEndRow: number,
  rowIndex: number,
): { kind: RowKind; markerName?: string } {
  if (rowIndex < headerEndRow) return { kind: 'skip' };
  if (!row || isBlankRow(row)) return { kind: 'blank' };

  const colA = cellStr(row[0]);
  const colB = cellStr(row[1]);

  // Main tenant row: col A (Property) is non-empty
  if (colA) {
    return { kind: 'tenant-main' };
  }

  // Sub-section marker: col A empty, col B has text, col C+ also has text labels
  // (the marker row doubles as the column header for that sub-section)
  if (!colA && colB) {
    // Check if col C onward has multiple text cells (column labels)
    const restCells = row.slice(2).filter(c => cellStr(c) !== '');
    if (restCells.length >= 3) {
      return { kind: 'sub-section-marker', markerName: colB };
    }
    // If col B has text but no labels follow, it's still a marker (just no data columns)
    // This handles edge cases — treat as marker anyway
    return { kind: 'sub-section-marker', markerName: colB };
  }

  // Sub-section data row: col A empty, col B empty, data in col C+
  if (!colA && !colB) {
    const hasData = row.slice(2).some(c => cellStr(c) !== '');
    if (hasData) return { kind: 'sub-section-data' };
  }

  return { kind: 'blank' };
}

// ─── Main parser ────────────────────────────────────────────────────────────

export function parseTenancySchedule(
  data: (string | number | null)[][],
  addLog?: (type: 'system' | 'flag', msg: string) => void,
): TenancyScheduleTenant[] {
  const log = addLog || (() => {});

  // 1. Detect main headers
  const { labels: mainHeaders, headerEndRow } = detectMainHeaders(data);
  log('system', `Tenancy Schedule: ${mainHeaders.filter(l => l).length} main columns detected, data starts at row ${headerEndRow + 1}`);
  if (mainHeaders.length > 0) {
    log('system', `Main headers: ${mainHeaders.filter(l => l).join(' | ')}`);
  }

  const tenants: TenancyScheduleTenant[] = [];
  let current: TenancyScheduleTenant | null = null;
  let currentSubSection: SubSection | null = null;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const { kind, markerName } = classifyRow(row, headerEndRow, i);

    switch (kind) {
      case 'skip':
      case 'blank':
        continue;

      case 'tenant-main': {
        // Finalize previous sub-section and tenant
        if (current) {
          if (currentSubSection) current.subSections.push(currentSubSection);
          tenants.push(current);
        }

        // Build main row values keyed by merged headers
        const mainRow: Record<string, string | number | null> = {};
        for (let col = 0; col < mainHeaders.length; col++) {
          const label = mainHeaders[col];
          if (!label) continue;
          mainRow[label] = (row && col < row.length) ? (row[col] ?? null) : null;
        }

        current = {
          mainRow,
          subSections: [],
          rawRows: [row],
        };
        currentSubSection = null;
        break;
      }

      case 'sub-section-marker': {
        if (!current) continue;

        // Finalize previous sub-section
        if (currentSubSection) {
          current.subSections.push(currentSubSection);
        }

        // Read column labels from col C (index 2) onward on this marker row
        const columnLabels: string[] = [];
        if (row) {
          for (let col = 2; col < row.length; col++) {
            columnLabels.push(cellStr(row[col]));
          }
        }
        // Trim trailing empty labels
        while (columnLabels.length > 0 && !columnLabels[columnLabels.length - 1]) {
          columnLabels.pop();
        }

        currentSubSection = {
          name: markerName!,
          columnLabels,
          rows: [],
        };

        current.rawRows.push(row);
        break;
      }

      case 'sub-section-data': {
        if (!current) continue;
        current.rawRows.push(row);

        if (!currentSubSection) {
          // Data before any marker — create an unnamed sub-section
          currentSubSection = { name: '(unknown)', columnLabels: [], rows: [] };
        }

        // Map cell values to column labels from the marker row
        const values: Record<string, string | number | null> = {};
        const labels = currentSubSection.columnLabels;
        if (row) {
          for (let j = 0; j < labels.length; j++) {
            const label = labels[j];
            if (!label) continue;
            const colIdx = j + 2; // labels start at col C (index 2)
            values[label] = colIdx < row.length ? (row[colIdx] ?? null) : null;
          }
        }

        currentSubSection.rows.push({
          values,
          rawRow: row,
        });
        break;
      }
    }
  }

  // Push last tenant
  if (current) {
    if (currentSubSection) current.subSections.push(currentSubSection);
    tenants.push(current);
  }

  log('system', `${tenants.length} tenancy schedule tenant blocks found.`);

  // Log sub-section summary
  const subSectionNames = new Set<string>();
  for (const t of tenants) {
    for (const s of t.subSections) {
      subSectionNames.add(s.name);
    }
  }
  if (subSectionNames.size > 0) {
    log('system', `Sub-sections found: ${[...subSectionNames].join(', ')}`);
  }

  return tenants;
}
