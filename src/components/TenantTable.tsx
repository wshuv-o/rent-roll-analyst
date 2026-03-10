import type { TenantObject } from '@/lib/types';
import { exportToExcel } from '@/lib/excel-utils';
import { Download } from 'lucide-react';

interface TenantTableProps {
  tenants: TenantObject[];
  fileName: string;
}

export function TenantTable({ tenants, fileName }: TenantTableProps) {
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
              <th className="text-left p-2 text-muted-foreground font-semibold">Lease Start</th>
              <th className="text-left p-2 text-muted-foreground font-semibold">Lease End</th>
              <th className="text-right p-2 text-muted-foreground font-semibold">GLA (SF)</th>
              <th className="text-right p-2 text-muted-foreground font-semibold">Monthly Rent</th>
              <th className="text-right p-2 text-muted-foreground font-semibold">Rent PSF</th>
              <th className="text-left p-2 text-muted-foreground font-semibold">Charges</th>
            </tr>
          </thead>
          <tbody>
            {tenants.map((t, i) => (
              <tr key={i} className="border-b border-panel-border/50 hover:bg-muted/30">
                <td className="p-2 text-muted-foreground">{i + 1}</td>
                <td className="p-2">{t.suite_id}</td>
                <td className="p-2">{t.tenant_name}</td>
                <td className="p-2">{t.lease_start}</td>
                <td className="p-2">{t.lease_end}</td>
                <td className="p-2 text-right tabular-nums">{t.gla_sqft?.toLocaleString() ?? '—'}</td>
                <td className="p-2 text-right tabular-nums">
                  {t.monthly_base_rent !== null ? `$${t.monthly_base_rent.toLocaleString()}` : '—'}
                </td>
                <td className="p-2 text-right tabular-nums">
                  {t.base_rent_psf !== null ? `$${t.base_rent_psf.toFixed(2)}` : '—'}
                </td>
                <td className="p-2 text-muted-foreground">
                  {t.recurring_charges.length > 0
                    ? t.recurring_charges.map(rc => rc.code).filter(Boolean).join(', ') || `${t.recurring_charges.length} charge(s)`
                    : '—'}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
