import type { ParsingInstruction, ColumnGroupId } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';
import { Check, RotateCcw, Loader2 } from 'lucide-react';

interface ColumnMappingToolbarProps {
  instruction: ParsingInstruction | null;
  isAnalyzing: boolean;
  onConfirm: () => void;
  onReset: () => void;
}

const GROUP_DOT_COLORS: Record<ColumnGroupId, string> = {
  'identity': 'bg-group-identity',
  'lease': 'bg-group-lease',
  'space': 'bg-group-space',
  'base-rent': 'bg-group-base-rent',
  'charges': 'bg-group-charges',
  'future-rent': 'bg-group-future-rent',
};

function colLetterClean(letter: string): string {
  return letter?.toUpperCase().replace(/[^A-Z]/g, '') || '—';
}

export function ColumnMappingToolbar({ instruction, isAnalyzing, onConfirm, onReset }: ColumnMappingToolbarProps) {
  if (isAnalyzing) {
    return (
      <div className="shrink-0 px-3 py-2 border-t border-panel-border bg-card flex items-center gap-2">
        <Loader2 className="w-3.5 h-3.5 animate-spin text-log-thinking" />
        <span className="text-[11px] font-mono text-muted-foreground">AI analyzing spreadsheet layout...</span>
      </div>
    );
  }

  if (!instruction) return null;

  return (
    <div className="shrink-0 border-t border-panel-border bg-card">
      {/* Group legend */}
      <div className="px-3 py-2 flex flex-wrap gap-x-4 gap-y-1 items-center">
        {COLUMN_GROUPS.map(group => {
          const mappedFields = group.fields
            .map(f => ({ field: f, col: instruction.column_map[f] }))
            .filter(f => f.col);

          if (mappedFields.length === 0) return null;

          return (
            <div key={group.id} className="flex items-center gap-1.5">
              <span className={`w-2 h-2 rounded-full ${GROUP_DOT_COLORS[group.id]}`} />
              <span className="text-[10px] font-heading uppercase tracking-wider text-muted-foreground">
                {group.label}
              </span>
              <span className="text-[10px] font-mono text-foreground/60">
                ({mappedFields.map(f => colLetterClean(f.col)).join(', ')})
              </span>
            </div>
          );
        })}
      </div>

      {/* Info + actions */}
      <div className="px-3 py-2 border-t border-panel-border/50 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <span className="text-[11px] font-mono text-muted-foreground">
            Data starts at row <span className="text-foreground font-semibold">{instruction.data_starts_at_row ?? '?'}</span>
          </span>
          <span className={`text-[11px] font-mono ${
            instruction.confidence === 'high' ? 'text-log-output' :
            instruction.confidence === 'medium' ? 'text-log-thinking' : 'text-log-flag'
          }`}>
            Confidence: {instruction.confidence}
          </span>
          {instruction.notes && (
            <span className="text-[10px] font-mono text-muted-foreground truncate max-w-[300px]" title={instruction.notes}>
              {instruction.notes}
            </span>
          )}
        </div>
        <div className="flex items-center gap-2">
          <span className="text-[10px] text-muted-foreground font-mono">
            Drag columns to reassign
          </span>
          <button
            onClick={onReset}
            className="flex items-center gap-1 px-2 py-1 text-[11px] font-mono rounded-sm bg-secondary text-secondary-foreground hover:bg-secondary/80 transition-colors"
          >
            <RotateCcw className="w-3 h-3" />
            Re-analyze
          </button>
          <button
            onClick={onConfirm}
            className="flex items-center gap-1 px-3 py-1 text-[11px] font-mono rounded-sm bg-log-output text-primary-foreground hover:opacity-90 transition-colors font-semibold"
          >
            <Check className="w-3.5 h-3.5" />
            Confirm & Parse
          </button>
        </div>
      </div>
    </div>
  );
}
