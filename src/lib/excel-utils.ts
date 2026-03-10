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
  const scalarGroupDefs = COLUMN_GROUPS.filter(g => g.id !== 'identity' && !g.collection);
  const collectionGroupDefs = COLUMN_GROUPS.filter(g => g.collection);

  // Check which groups have data
  const scalarIds = new Set<string>();
  const collectionIds = new Set<string>();
  for (const t of tenants) {
    for (const gid of Object.keys(t.scalars)) scalarIds.add(gid);
    for (const gid of Object.keys(t.collections)) collectionIds.add(gid);
  }

  const activeScalars = scalarGroupDefs.filter(g => scalarIds.has(g.id));
  const activeCollections = collectionGroupDefs.filter(g => collectionIds.has(g.id));

  const serializeRecord = (rec: Record<string, string | number | null>) =>
    Object.entries(rec)
      .filter(([, v]) => v !== null && v !== '')
      .map(([k, v]) => `${k}: ${v}`)
      .join(' | ');

  const rows = tenants.map(t => {
    const base: Record<string, unknown> = {
      'Suite ID': t.suite_id,
      'Tenant Name': t.tenant_name,
    };
    for (const g of activeScalars) {
      base[g.label] = t.scalars[g.id] ? serializeRecord(t.scalars[g.id]) : '';
    }
    for (const g of activeCollections) {
      const entries = t.collections[g.id];
      base[g.label] = entries ? entries.map(serializeRecord).join('; ') : '';
    }
    base['Notes'] = t.notes;
    return base;
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Parsed Rent Roll');

  const outputName = fileName.replace(/\.(xlsx|xls)$/i, '') + '_parsed.xlsx';
  XLSX.writeFile(wb, outputName);
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)}KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}
