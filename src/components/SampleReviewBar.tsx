import { Slider } from '@/components/ui/slider';

interface SampleReviewBarProps {
  sampleRows: number;
  sampleCols: number;
  maxAvailableRows: number;
  maxAvailableCols: number;
  onRowsChange: (n: number) => void;
  onColsChange: (n: number) => void;
  onConfirmSend: () => void;
  isProcessing: boolean;
}

const MAX_ROWS = 50;
const MAX_COLS = 100;

export function SampleReviewBar({
  sampleRows,
  sampleCols,
  maxAvailableRows,
  maxAvailableCols,
  onRowsChange,
  onColsChange,
  onConfirmSend,
  isProcessing,
}: SampleReviewBarProps) {
  const rowLimit = Math.min(MAX_ROWS, maxAvailableRows);
  const colLimit = Math.min(MAX_COLS, maxAvailableCols);

  return (
    <div className="shrink-0 border-t border-panel-border bg-card px-4 py-2.5 flex items-center gap-6">
      <div className="flex items-center gap-2 flex-1 min-w-0">
        <div className="flex items-center gap-3 flex-1">
          <label className="text-[10px] font-mono text-muted-foreground uppercase tracking-wider whitespace-nowrap">
            Rows
          </label>
          <Slider
            min={1}
            max={rowLimit}
            step={1}
            value={[sampleRows]}
            onValueChange={([v]) => onRowsChange(v)}
            className="w-32"
          />
          <span className="text-[11px] font-mono text-foreground min-w-[2.5rem] text-right">
            {sampleRows}
          </span>
        </div>

        <div className="w-px h-4 bg-panel-border" />

        <div className="flex items-center gap-3 flex-1">
          <label className="text-[10px] font-mono text-muted-foreground uppercase tracking-wider whitespace-nowrap">
            Cols
          </label>
          <Slider
            min={1}
            max={colLimit}
            step={1}
            value={[sampleCols]}
            onValueChange={([v]) => onColsChange(v)}
            className="w-32"
          />
          <span className="text-[11px] font-mono text-foreground min-w-[2.5rem] text-right">
            {sampleCols}
          </span>
        </div>
      </div>

      <div className="text-[10px] font-mono text-muted-foreground">
        {sampleRows} × {sampleCols} cells will be anonymized &amp; sent
      </div>

      <button
        onClick={onConfirmSend}
        disabled={isProcessing}
        className="px-4 py-1.5 text-[11px] font-mono font-semibold bg-primary text-primary-foreground rounded-sm hover:opacity-90 transition-opacity disabled:opacity-40"
      >
        Send to AI
      </button>
    </div>
  );
}
