import * as XLSX from 'xlsx';
import type { TenantObject } from './types';

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

        // Convert to array of arrays, preserving structure
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
  // Flatten tenant objects for Excel
  const rows = tenants.map(t => ({
    'Suite ID': t.suite_id,
    'Tenant Name': t.tenant_name,
    'Lease Start': t.lease_start,
    'Lease End': t.lease_end,
    'GLA (SF)': t.gla_sqft,
    'Monthly Base Rent': t.monthly_base_rent,
    'Base Rent PSF': t.base_rent_psf,
    'Recurring Charges': t.recurring_charges.map(rc =>
      `${rc.code}: $${rc.amount ?? 'N/A'}`
    ).join('; '),
    'Future Rent Increases': t.future_rent_increases.map(fr =>
      `${fr.effective_date}: $${fr.monthly_amount ?? 'N/A'}`
    ).join('; '),
    'Notes': t.notes,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Parsed Rent Roll');

  // Auto-size columns
  const colWidths = Object.keys(rows[0] || {}).map(key => ({
    wch: Math.max(key.length, ...rows.map(r => String((r as Record<string, unknown>)[key] || '').length))
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
