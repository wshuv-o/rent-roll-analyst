import * as XLSX from 'xlsx';
import type { TenantObject } from './types';
import { COLUMN_GROUPS } from './types';

export function readExcelFile(file: File): Promise<{
  data: (string | number | null)[][];
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
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const jsonData = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
          header: 1,
          defval: null,
          raw: false,
        });

        resolve({
          data: jsonData as (string | number | null)[][],
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

export function exportToExcel(tenants: TenantObject[], fileName: string): void {
  // Collect all group ids present
  const groupIds = new Set<string>();
  for (const t of tenants) {
    for (const gid of Object.keys(t.groups)) groupIds.add(gid);
  }

  // Order by COLUMN_GROUPS definition
  const orderedGroups: string[] = COLUMN_GROUPS
    .filter(g => g.id !== 'identity' && groupIds.has(g.id))
    .map(g => g.id);
  for (const gid of groupIds) {
    if (!orderedGroups.includes(gid)) orderedGroups.push(gid);
  }

  const rows = tenants.map(t => {
    const base: Record<string, unknown> = {
      'Suite ID': t.suite_id,
      'Tenant Name': t.tenant_name,
    };

    for (const gid of orderedGroups) {
      const groupLabel = COLUMN_GROUPS.find(g => g.id === gid)?.label || gid;
      const entries = t.groups[gid];
      if (!entries || entries.length === 0) {
        base[groupLabel] = '';
        continue;
      }
      // Serialize: each entry as "key: val | key: val", entries separated by ";"
      base[groupLabel] = entries.map(entry =>
        Object.entries(entry)
          .filter(([, v]) => v !== null && v !== '')
          .map(([k, v]) => `${k}: ${v}`)
          .join(' | ')
      ).join('; ');
    }

    base['Notes'] = t.notes;
    return base;
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Parsed Rent Roll');

  const colWidths = Object.keys(rows[0] || {}).map(key => ({
    wch: Math.max(key.length, ...rows.map(r => String((r as Record<string, unknown>)[key] || '').length)).toString().length > 50 ? 50 : Math.max(key.length, ...rows.map(r => String((r as Record<string, unknown>)[key] || '').length))
  }));
  ws['!cols'] = colWidths;

  const outputName = fileName.replace(/\.(xlsx|xls)$/i, '') + '_parsed.xlsx';
  XLSX.writeFile(wb, outputName);
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)}KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}
