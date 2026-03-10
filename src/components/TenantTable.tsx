import { useMemo } from 'react';
import type { TenantObject, ColumnGroupId } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';
import { exportToExcel } from '@/lib/excel-utils';
import { Download } from 'lucide-react';

const GROUP_COLORS: Record<ColumnGroupId, string> = {
  'identity': 'text-group-identity',
  'lease': 'text-group-lease',
  'space': 'text-group-space',
  'base-rent': 'text-group-base-rent',
  'charges': 'text-group-charges',
  'future-rent': 'text-group-future-rent',
};

interface TenantTableProps {
  tenants: TenantObject[];
  fileName: string;
}

/** Format a group's collected rows into a readable cell */
function formatGroupCell(rows: Record<string, string | number | null>[] | undefined): string {
  if (!rows || rows.length === 0) return '—';
  return rows.map(entry => {
    const parts = Object.entries(entry)
      .filter(([, v]) => v !== null && v !== '')
      .map(([k, v]) => `${k}: ${v}`);
    return parts.join(' | ');
  }).join('; ');
}

export function TenantTable({ tenants, fileName }: TenantTableProps) {
  // Determine which groups are present across all tenants (excluding identity)
  const presentGroups = useMemo(() => {
    const groupIds = new Set<string>();
    for (const t of tenants) {
      for (const gid of Object.keys(t.groups)) groupIds.add(gid);
    }
    // Order them by COLUMN_GROUPS definition order
    const ordered = COLUMN_GROUPS
      .filter(g => g.id !== 'identity' && groupIds.has(g.id))
      .map(g => g);
    // Also include any unknown group ids (from custom expansions)
    for (const gid of groupIds) {
      if (!ordered.find(g => g.id === gid)) {
        ordered.push({ id: gid as ColumnGroupId, label: gid, fields: [], fieldLabels: {} });
      }
    }
    return ordered;
  }, [tenants]);

  return (
    <div className="flex flex-col gap-3">
      <div className="flex items-center justify-between">
        <span className="font-mono text-sm text-muted-foreground">
          {tenants.length} tenant{tenants.length !== 1 ? 's' : ''} parsed
        </span>
        <button
          onClick={() => exportToExcel(tenants, fileName)}
          className="flex items-center gap-2 px-3 py-1.5 text-xs font-mono rounded-sm bg-secondary text-secondary-foreground hover:bg-secondary/80 transition-colors"
        >
          <Download className="w-3.5 h-3.5" />
          Download Excel
        </button>
      </div>
      <div className="overflow-auto max-h-[calc(100vh-280px)] border border-panel-border rounded-sm">
        <table className="w-full text-xs font-mono">
          <thead className="sticky top-0 bg-card">
            <tr className="border-b border-panel-border">
              <th className="text-left p-2 text-muted-foreground font-semibold">#</th>
              <th className="text-left p-2 text-muted-foreground font-semibold">Suite</th>
              <th className="text-left p-2 text-muted-foreground font-semibold">Tenant</th>
              {presentGroups.map(g => (
                <th key={g.id} className={`text-left p-2 font-semibold ${GROUP_COLORS[g.id] || 'text-muted-foreground'}`}>
                  {g.label}
                </th>
              ))}
              <th className="text-left p-2 text-muted-foreground font-semibold">Notes</th>
            </tr>
          </thead>
          <tbody>
            {tenants.map((t, i) => (
              <tr key={i} className="border-b border-panel-border/50 hover:bg-muted/30">
                <td className="p-2 text-muted-foreground">{i + 1}</td>
                <td className="p-2">{t.suite_id}</td>
                <td className="p-2">{t.tenant_name}</td>
                {presentGroups.map(g => (
                  <td key={g.id} className="p-2 max-w-[300px]">
                    <div className="whitespace-pre-wrap text-[11px]">
                      {formatGroupCell(t.groups[g.id])}
                    </div>
                  </td>
                ))}
                <td className="p-2 text-muted-foreground">{t.notes || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
