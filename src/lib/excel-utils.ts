import * as XLSX from 'xlsx';
import type { TenantObject, ParsingInstruction, GroupSpan } from './types';
import { COLUMN_GROUPS } from './types';
import { colLetterToIndex, getCellValue, indexToColLetter } from './col-utils';

/**
 * Pick the best sheet from a workbook.
 * Prefers sheets whose first 15 rows contain rent-roll header keywords.
 * Falls back to the sheet with the most rows.
 */
function pickBestSheet(workbook: XLSX.WorkBook): string {
  const KEYWORDS = ['unit', 'dba', 'lease', 'tenant', 'suite', 'rent', 'sqft', 'square footage'];

  let bestName = workbook.SheetNames[0];
  let bestScore = -1;

  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    const sample = XLSX.utils.sheet_to_json<(string | number | Date | null)[]>(sheet, {
      header: 1, defval: null, raw: true, range: 0,
    });

    // Score by keyword matches in the first 15 rows
    let score = 0;
    const limit = Math.min(15, sample.length);
    for (let r = 0; r < limit; r++) {
      const row = sample[r];
      if (!row) continue;
      for (const cell of row) {
        if (typeof cell === 'string') {
          const lower = cell.toLowerCase();
          for (const kw of KEYWORDS) {
            if (lower.includes(kw)) score++;
          }
        }
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestName = name;
    }
  }

  return bestName;
}

export function readExcelFile(file: File): Promise<{
  data: (string | number | Date | null)[][];
  totalRows: number;
  fileName: string;
  fileSize: number;
}> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const arrayBuffer = e.target?.result;
        if (!arrayBuffer) throw new Error('Failed to read file');

        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
        const sheetName = pickBestSheet(workbook);
        const sheet = workbook.Sheets[sheetName];

        const jsonData = XLSX.utils.sheet_to_json<(string | number | Date | null)[]>(sheet, {
          header: 1,
          defval: null,
          raw: true,
        });

        resolve({
          data: jsonData as (string | number | Date | null)[][],
          totalRows: jsonData.length,
          fileName: file.name,
          fileSize: file.size,
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

export function exportToExcel(
  tenants: TenantObject[],
  fileName: string,
  instruction: ParsingInstruction,
  groupSpans: GroupSpan[],
  columnLabels: Record<number, string>
): void {
  const cm = instruction.column_map;

  // Build rows from raw data using column mapping
  const rows = tenants.map(t => {
    const primaryRow = t.rawRows[0] || [];
    const base: Record<string, unknown> = {
      'Suite ID': t.suite_id,
      'Tenant Name': t.tenant_name,
    };

    // Scalar groups
    for (const group of COLUMN_GROUPS.filter(g => g.id !== 'identity' && !g.collection)) {
      const span = groupSpans.find(s => s.groupId === group.id);
      if (!span) continue;
      const parts: string[] = [];
      for (let c = span.startCol; c <= span.endCol; c++) {
        const label = columnLabels[c] || indexToColLetter(c);
        const val = c < primaryRow.length ? primaryRow[c] : null;
        if (val !== null && val !== undefined && String(val).trim()) {
          parts.push(`${label}: ${String(val).trim()}`);
        }
      }
      base[group.label] = parts.join(' | ');
    }

    // Collection groups
    for (const group of COLUMN_GROUPS.filter(g => g.collection)) {
      const span = groupSpans.find(s => s.groupId === group.id);
      if (!span) continue;
      const entries: string[] = [];
      for (const row of t.rawRows) {
        const parts: string[] = [];
        let hasData = false;
        for (let c = span.startCol; c <= span.endCol; c++) {
          const label = columnLabels[c] || indexToColLetter(c);
          const val = c < row.length ? row[c] : null;
          if (val !== null && val !== undefined && String(val).trim()) {
            parts.push(`${label}: ${String(val).trim()}`);
            hasData = true;
          }
        }
        if (hasData) entries.push(parts.join(' | '));
      }
      base[group.label] = entries.join('; ');
    }

    base['Notes'] = t.notes;
    return base;
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Parsed Rent Roll');

  const outputName = fileName.replace(/\.(xlsx|xls|csv)$/i, '') + '_parsed.xlsx';
  XLSX.writeFile(wb, outputName);
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)}KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}
