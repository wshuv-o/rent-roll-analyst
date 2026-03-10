import { useRef, useEffect, useMemo, useState, useCallback } from 'react';
import type { ParsingInstruction, ColumnGroupId, COLUMN_GROUPS as ColumnGroupsType } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';

interface SpreadsheetViewerProps {
  data: (string | number | null)[][];
  instruction: ParsingInstruction | null;
  headerRows: number[];
  onColumnAssign?: (colIndex: number, field: string) => void;
}

// Map column letter to 0-based index
function colLetterToIndex(letter: string): number {
  if (!letter) return -1;
  const upper = letter.toUpperCase().trim().replace(/[^A-Z]/g, '');
  if (!upper) return -1;
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

function indexToColLetter(idx: number): string {
  let letter = '';
  let n = idx;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

// Group color CSS classes
const GROUP_COLORS: Record<ColumnGroupId, { border: string; bg: string; text: string }> = {
  'identity': { border: 'border-group-identity', bg: 'bg-group-identity-bg', text: 'text-group-identity' },
  'lease': { border: 'border-group-lease', bg: 'bg-group-lease-bg', text: 'text-group-lease' },
  'space': { border: 'border-group-space', bg: 'bg-group-space-bg', text: 'text-group-space' },
  'base-rent': { border: 'border-group-base-rent', bg: 'bg-group-base-rent-bg', text: 'text-group-base-rent' },
  'charges': { border: 'border-group-charges', bg: 'bg-group-charges-bg', text: 'text-group-charges' },
  'future-rent': { border: 'border-group-future-rent', bg: 'bg-group-future-rent-bg', text: 'text-group-future-rent' },
};

export function SpreadsheetViewer({ data, instruction, headerRows, onColumnAssign }: SpreadsheetViewerProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [dragStart, setDragStart] = useState<number | null>(null);
  const [dragEnd, setDragEnd] = useState<number | null>(null);
  const [showAssignMenu, setShowAssignMenu] = useState<{ colIndices: number[]; x: number; y: number } | null>(null);

  const maxCols = useMemo(() => data.reduce((max, row) => Math.max(max, row.length), 0), [data]);
  const visibleRows = useMemo(() => Math.min(data.length, 200), [data]);

  // Build column → group mapping from instruction
  const colGroupMap = useMemo(() => {
    const map = new Map<number, { groupId: ColumnGroupId; field: string; fieldLabel: string }>();
    if (!instruction) return map;

    for (const group of COLUMN_GROUPS) {
      for (const field of group.fields) {
        const letter = instruction.column_map[field];
        if (letter) {
          const idx = colLetterToIndex(letter);
          if (idx >= 0) {
            map.set(idx, {
              groupId: group.id,
              field,
              fieldLabel: group.fieldLabels[field] || field,
            });
          }
        }
      }
    }
    return map;
  }, [instruction]);

  // Build group spans for header border rendering
  const groupSpans = useMemo(() => {
    const spans: { groupId: ColumnGroupId; label: string; startCol: number; endCol: number }[] = [];
    if (!instruction) return spans;

    for (const group of COLUMN_GROUPS) {
      const indices: number[] = [];
      for (const field of group.fields) {
        const letter = instruction.column_map[field];
        if (letter) {
          const idx = colLetterToIndex(letter);
          if (idx >= 0) indices.push(idx);
        }
      }
      if (indices.length > 0) {
        spans.push({
          groupId: group.id,
          label: group.label,
          startCol: Math.min(...indices),
          endCol: Math.max(...indices),
        });
      }
    }
    return spans.sort((a, b) => a.startCol - b.startCol);
  }, [instruction]);

  const headerRowSet = useMemo(() => new Set(headerRows), [headerRows]);
  const dataStartRow = instruction?.data_starts_at_row ? instruction.data_starts_at_row - 1 : 0;

  // Handle column drag selection
  const handleColMouseDown = useCallback((colIdx: number, e: React.MouseEvent) => {
    if (!onColumnAssign) return;
    e.preventDefault();
    setDragStart(colIdx);
    setDragEnd(colIdx);
  }, [onColumnAssign]);

  const handleColMouseEnter = useCallback((colIdx: number) => {
    if (dragStart !== null) {
      setDragEnd(colIdx);
    }
  }, [dragStart]);

  const handleColMouseUp = useCallback((e: React.MouseEvent) => {
    if (dragStart !== null && dragEnd !== null) {
      const start = Math.min(dragStart, dragEnd);
      const end = Math.max(dragStart, dragEnd);
      const indices = Array.from({ length: end - start + 1 }, (_, i) => start + i);
      setShowAssignMenu({ colIndices: indices, x: e.clientX, y: e.clientY });
    }
    setDragStart(null);
    setDragEnd(null);
  }, [dragStart, dragEnd]);

  const isDragSelected = useCallback((colIdx: number) => {
    if (dragStart === null || dragEnd === null) return false;
    const start = Math.min(dragStart, dragEnd);
    const end = Math.max(dragStart, dragEnd);
    return colIdx >= start && colIdx <= end;
  }, [dragStart, dragEnd]);

  // Close menu on outside click
  useEffect(() => {
    const handler = () => setShowAssignMenu(null);
    if (showAssignMenu) {
      document.addEventListener('click', handler, { once: true });
    }
    return () => document.removeEventListener('click', handler);
  }, [showAssignMenu]);

  const handleAssignField = useCallback((field: string) => {
    if (!showAssignMenu || !onColumnAssign) return;
    for (const colIdx of showAssignMenu.colIndices) {
      onColumnAssign(colIdx, field);
    }
    setShowAssignMenu(null);
  }, [showAssignMenu, onColumnAssign]);

  return (
    <div className="relative flex flex-col h-full">
      {/* Group header band */}
      {groupSpans.length > 0 && (
        <div className="shrink-0 overflow-x-auto" style={{ marginLeft: '40px' }}>
          <div className="flex relative h-6" style={{ minWidth: `${maxCols * 100}px` }}>
            {groupSpans.map(span => (
              <div
                key={span.groupId}
                className={`absolute h-full flex items-center justify-center text-[10px] font-heading uppercase tracking-wider border-t-2 border-l-2 border-r-2 rounded-t-sm ${GROUP_COLORS[span.groupId].border} ${GROUP_COLORS[span.groupId].bg} ${GROUP_COLORS[span.groupId].text}`}
                style={{
                  left: `${span.startCol * 100}px`,
                  width: `${(span.endCol - span.startCol + 1) * 100}px`,
                }}
              >
                {span.label}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Spreadsheet grid */}
      <div
        ref={containerRef}
        className="flex-1 overflow-auto"
        onMouseUp={handleColMouseUp}
      >
        <table className="border-collapse text-[11px] font-mono" style={{ minWidth: `${maxCols * 100 + 40}px` }}>
          {/* Column letters header */}
          <thead className="sticky top-0 z-10">
            <tr className="bg-card">
              <th className="w-[40px] min-w-[40px] p-0 border-r border-b border-panel-border bg-card sticky left-0 z-20" />
              {Array.from({ length: maxCols }, (_, c) => {
                const groupInfo = colGroupMap.get(c);
                const selected = isDragSelected(c);
                return (
                  <th
                    key={c}
                    className={`min-w-[100px] max-w-[180px] p-1 border-r border-b border-panel-border text-center select-none cursor-crosshair transition-colors ${
                      selected ? 'bg-muted' : groupInfo ? GROUP_COLORS[groupInfo.groupId].bg : 'bg-card'
                    } ${groupInfo ? `border-b-2 ${GROUP_COLORS[groupInfo.groupId].border}` : ''}`}
                    onMouseDown={(e) => handleColMouseDown(c, e)}
                    onMouseEnter={() => handleColMouseEnter(c)}
                  >
                    <div className="text-muted-foreground text-[10px]">{indexToColLetter(c)}</div>
                    {groupInfo && (
                      <div className={`text-[9px] ${GROUP_COLORS[groupInfo.groupId].text} truncate`}>
                        {groupInfo.fieldLabel}
                      </div>
                    )}
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {data.slice(0, visibleRows).map((row, rowIdx) => {
              const isHeader = headerRowSet.has(rowIdx);
              const isDataStart = rowIdx === dataStartRow;
              return (
                <tr
                  key={rowIdx}
                  className={`${isHeader ? 'bg-group-header-row' : ''} ${isDataStart ? 'border-t-2 border-t-log-output' : ''}`}
                >
                  {/* Row number */}
                  <td className="w-[40px] min-w-[40px] px-1 py-0.5 text-right text-[10px] text-muted-foreground border-r border-b border-panel-border bg-card sticky left-0 z-10 select-none">
                    {rowIdx + 1}
                  </td>
                  {Array.from({ length: maxCols }, (_, c) => {
                    const val = c < row.length ? row[c] : null;
                    const groupInfo = colGroupMap.get(c);
                    const selected = isDragSelected(c);
                    return (
                      <td
                        key={c}
                        className={`min-w-[100px] max-w-[180px] px-1.5 py-0.5 border-r border-b border-panel-border/50 truncate ${
                          selected ? 'bg-muted' : groupInfo ? GROUP_COLORS[groupInfo.groupId].bg : ''
                        } ${isHeader ? 'font-semibold text-foreground' : 'text-foreground/80'} ${
                          groupInfo ? `border-l ${GROUP_COLORS[groupInfo.groupId].border}/30` : ''
                        }`}
                        title={val !== null && val !== undefined ? String(val) : ''}
                      >
                        {val !== null && val !== undefined ? String(val) : ''}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>

        {data.length > visibleRows && (
          <div className="py-2 text-center text-[10px] text-muted-foreground font-mono">
            Showing {visibleRows} of {data.length} rows
          </div>
        )}
      </div>

      {/* Assign menu (appears after drag-select) */}
      {showAssignMenu && (
        <div
          className="fixed z-50 bg-card border border-panel-border rounded-sm shadow-lg py-1 min-w-[200px]"
          style={{ left: showAssignMenu.x, top: showAssignMenu.y }}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="px-2 py-1 text-[10px] text-muted-foreground font-heading uppercase tracking-wider border-b border-panel-border">
            Assign {showAssignMenu.colIndices.length > 1 ? `Columns ${indexToColLetter(showAssignMenu.colIndices[0])}–${indexToColLetter(showAssignMenu.colIndices[showAssignMenu.colIndices.length - 1])}` : `Column ${indexToColLetter(showAssignMenu.colIndices[0])}`}
          </div>
          {COLUMN_GROUPS.map(group => (
            <div key={group.id}>
              <div className={`px-2 pt-1.5 pb-0.5 text-[9px] font-heading uppercase tracking-wider ${GROUP_COLORS[group.id].text}`}>
                {group.label}
              </div>
              {group.fields.map((field, fi) => (
                <button
                  key={field}
                  className="w-full text-left px-3 py-1 text-[11px] font-mono text-foreground/80 hover:bg-muted transition-colors"
                  onClick={() => {
                    // Assign sequentially: first selected col gets first field, etc.
                    const colIdx = showAssignMenu.colIndices[Math.min(fi, showAssignMenu.colIndices.length - 1)];
                    onColumnAssign?.(colIdx, field);
                    if (fi === group.fields.length - 1 || fi === showAssignMenu.colIndices.length - 1) {
                      setShowAssignMenu(null);
                    }
                  }}
                >
                  {group.fieldLabels[field]}
                </button>
              ))}
            </div>
          ))}
          <div className="border-t border-panel-border mt-1">
            <button
              className="w-full text-left px-3 py-1 text-[11px] font-mono text-destructive hover:bg-muted transition-colors"
              onClick={() => {
                for (const colIdx of showAssignMenu.colIndices) {
                  onColumnAssign?.(colIdx, '');
                }
                setShowAssignMenu(null);
              }}
            >
              Clear Assignment
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
