import { useMemo } from 'react';
import type { TenantObject, ParsingInstruction, GroupSpan, ColumnGroupId, CustomGroup } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';
import { exportToExcel } from '@/lib/excel-utils';
import { exportTemplatizedRentRoll } from '@/lib/template-export';
import { getCellValue, indexToColLetter } from '@/lib/col-utils';
import { Download, FileSpreadsheet, ArrowLeft } from 'lucide-react';

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
  instruction: ParsingInstruction;
  groupSpans: GroupSpan[];
  columnLabels: Record<number, string>;
  customGroups?: CustomGroup[];
  onBack?: () => void;
}

function formatSpanValues(
  rows: (string | number | Date | null)[][],
  span: GroupSpan,
  columnLabels: Record<number, string>,
  collection: boolean
): string {
  if (collection) {
    const entries: string[] = [];
    for (const row of rows) {
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
    return entries.length > 0 ? entries.join('\n') : '—';
  } else {
    const row = rows[0] || [];
    const parts: string[] = [];
    for (let c = span.startCol; c <= span.endCol; c++) {
      const label = columnLabels[c] || indexToColLetter(c);
      const val = c < row.length ? row[c] : null;
      if (val !== null && val !== undefined && String(val).trim()) {
        parts.push(`${label}: ${String(val).trim()}`);
      }
    }
    return parts.length > 0 ? parts.join(' | ') : '—';
  }
}

export function TenantTable({ tenants, fileName, instruction, groupSpans, columnLabels, customGroups = [], onBack }: TenantTableProps) {
  const displayGroups = useMemo(() => {
    return groupSpans
      .filter(s => s.groupId !== 'identity')
      .map(s => {
        const builtIn = COLUMN_GROUPS.find(g => g.id === s.groupId);
        const custom = customGroups.find(cg => cg.id === s.groupId);
        return {
          span: s,
          label: builtIn?.label || custom?.label || s.groupId,
          colorClass: GROUP_COLORS[s.groupId as ColumnGroupId] || 'text-muted-foreground',
        };
      });
  }, [groupSpans, customGroups]);

  return (
    <div className="flex flex-col gap-3">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          {onBack && (
            <button
              onClick={onBack}
              className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-mono rounded-sm text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              Back
            </button>
          )}
          <span className="font-mono text-sm text-muted-foreground">
            {tenants.length} tenant{tenants.length !== 1 ? 's' : ''} parsed
          </span>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => exportToExcel(tenants, fileName, instruction, groupSpans, columnLabels)}
            className="flex items-center gap-2 px-3 py-1.5 text-xs font-mono rounded-sm bg-secondary text-secondary-foreground hover:bg-secondary/80 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Download Raw
          </button>
          <button
            onClick={() => exportTemplatizedRentRoll(tenants, fileName, instruction, groupSpans, columnLabels, customGroups)}
            className="flex items-center gap-2 px-3 py-1.5 text-xs font-mono rounded-sm bg-accent text-accent-foreground hover:bg-accent/80 border border-border transition-colors"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            Download Template
          </button>
        </div>
      </div>
      <div className="overflow-auto max-h-[calc(100vh-280px)] border border-panel-border rounded-sm">
        <table className="w-full text-xs font-mono">
          <thead className="sticky top-0 bg-card">
            <tr className="border-b border-panel-border">
              <th className="text-left p-2 text-muted-foreground font-semibold">#</th>
              <th className="text-left p-2 text-muted-foreground font-semibold">Suite</th>
              <th className="text-left p-2 text-muted-foreground font-semibold">Tenant</th>
              {displayGroups.map(g => (
                <th key={g.span.groupId} className={`text-left p-2 font-semibold ${g.colorClass}`}>
                  {g.label}{g.span.collection ? ' ⟨list⟩' : ''}
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
                {displayGroups.map(g => (
                  <td key={g.span.groupId} className="p-2 max-w-[300px]">
                    <div className="whitespace-pre-wrap text-[11px]">
                      {formatSpanValues(t.rawRows, g.span, columnLabels, g.span.collection)}
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
