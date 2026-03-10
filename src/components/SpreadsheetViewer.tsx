import { useRef, useMemo, useState, useCallback, useEffect } from 'react';
import type { ParsingInstruction, ColumnGroupId, GroupSpan } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';

interface SpreadsheetViewerProps {
  data: (string | number | null)[][];
  instruction: ParsingInstruction | null;
  headerRows: number[];
  groupSpans: GroupSpan[];
  columnAliases?: Record<number, string>;
  onColumnAssign?: (colIndex: number, field: string) => void;
  onCustomFieldAssign?: (colIndex: number, fieldName: string) => void;
  onGroupResize?: (groupId: ColumnGroupId, startCol: number, endCol: number) => void;
  onColumnRename?: (colIndex: number, name: string) => void;
}

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

const GROUP_COLORS: Record<ColumnGroupId, { border: string; bg: string; text: string; hsl: string }> = {
  'identity':    { border: 'border-group-identity',    bg: 'bg-group-identity-bg',    text: 'text-group-identity',    hsl: 'var(--group-identity)' },
  'lease':       { border: 'border-group-lease',       bg: 'bg-group-lease-bg',       text: 'text-group-lease',       hsl: 'var(--group-lease)' },
  'space':       { border: 'border-group-space',       bg: 'bg-group-space-bg',       text: 'text-group-space',       hsl: 'var(--group-space)' },
  'base-rent':   { border: 'border-group-base-rent',   bg: 'bg-group-base-rent-bg',   text: 'text-group-base-rent',   hsl: 'var(--group-base-rent)' },
  'charges':     { border: 'border-group-charges',     bg: 'bg-group-charges-bg',     text: 'text-group-charges',     hsl: 'var(--group-charges)' },
  'future-rent': { border: 'border-group-future-rent', bg: 'bg-group-future-rent-bg', text: 'text-group-future-rent', hsl: 'var(--group-future-rent)' },
};

const COL_WIDTH = 100;
const ROW_NUM_WIDTH = 40;

export function SpreadsheetViewer({
  data,
  instruction,
  headerRows,
  groupSpans,
  columnAliases = {},
  onColumnAssign,
  onCustomFieldAssign,
  onGroupResize,
  onColumnRename,
}: SpreadsheetViewerProps) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const [assignMenuCol, setAssignMenuCol] = useState<{ colIndex: number; x: number; y: number } | null>(null);
  const [customFieldInput, setCustomFieldInput] = useState('');

  // Drag-resize state
  const [resizing, setResizing] = useState<{
    groupId: ColumnGroupId;
    edge: 'left' | 'right';
    originalStart: number;
    originalEnd: number;
    currentCol: number;
  } | null>(null);

  const maxCols = useMemo(() => data.reduce((max, row) => Math.max(max, row.length), 0), [data]);
  const visibleRows = useMemo(() => Math.min(data.length, 200), [data]);

  // Build column → group + field mapping (includes custom columns)
  const colFieldMap = useMemo(() => {
    const map = new Map<number, { groupId: ColumnGroupId | 'custom'; field: string; fieldLabel: string }>();
    if (!instruction) return map;
    for (const group of COLUMN_GROUPS) {
      for (const field of group.fields) {
        const letter = instruction.column_map[field];
        if (letter) {
          const idx = colLetterToIndex(letter);
          if (idx >= 0) {
            map.set(idx, { groupId: group.id, field, fieldLabel: group.fieldLabels[field] || field });
          }
        }
      }
    }
    // Custom columns
    if (instruction.custom_columns) {
      for (const [fieldName, letter] of Object.entries(instruction.custom_columns)) {
        if (letter) {
          const idx = colLetterToIndex(letter);
          if (idx >= 0) {
            map.set(idx, { groupId: 'custom', field: fieldName, fieldLabel: fieldName });
          }
        }
      }
    }
    return map;
  }, [instruction]);

  // Compute live spans during resize
  const liveSpans = useMemo(() => {
    if (!resizing) return groupSpans;
    return groupSpans.map(s => {
      if (s.groupId !== resizing.groupId) return s;
      if (resizing.edge === 'left') {
        return { ...s, startCol: Math.min(resizing.currentCol, s.endCol) };
      } else {
        return { ...s, endCol: Math.max(resizing.currentCol, s.startCol) };
      }
    });
  }, [groupSpans, resizing]);

  const headerRowSet = useMemo(() => new Set(headerRows), [headerRows]);
  const dataStartRow = instruction?.data_starts_at_row ? instruction.data_starts_at_row - 1 : 0;

  // Resize drag handlers
  const handleResizeStart = useCallback((groupId: ColumnGroupId, edge: 'left' | 'right', e: React.MouseEvent) => {
    if (!onGroupResize) return;
    e.preventDefault();
    e.stopPropagation();
    const span = groupSpans.find(s => s.groupId === groupId);
    if (!span) return;
    setResizing({
      groupId, edge,
      originalStart: span.startCol,
      originalEnd: span.endCol,
      currentCol: edge === 'left' ? span.startCol : span.endCol,
    });
  }, [groupSpans, onGroupResize]);

  useEffect(() => {
    if (!resizing) return;

    const handleMouseMove = (e: MouseEvent) => {
      const container = scrollRef.current;
      if (!container) return;
      const rect = container.getBoundingClientRect();
      const scrollLeft = container.scrollLeft;
      const x = e.clientX - rect.left + scrollLeft - ROW_NUM_WIDTH;
      const col = Math.max(0, Math.min(maxCols - 1, Math.floor(x / COL_WIDTH)));
      setResizing(prev => prev ? { ...prev, currentCol: col } : null);
    };

    const handleMouseUp = () => {
      if (resizing && onGroupResize) {
        const span = groupSpans.find(s => s.groupId === resizing.groupId);
        if (span) {
          const newStart = resizing.edge === 'left' ? Math.min(resizing.currentCol, span.endCol) : span.startCol;
          const newEnd = resizing.edge === 'right' ? Math.max(resizing.currentCol, span.startCol) : span.endCol;
          onGroupResize(resizing.groupId, newStart, newEnd);
        }
      }
      setResizing(null);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizing, onGroupResize, groupSpans, maxCols]);

  // Column header click → open assign menu
  const handleColHeaderClick = useCallback((colIndex: number, e: React.MouseEvent) => {
    if (!onColumnAssign) return;
    // Don't open menu if clicking on the editable name area
    const target = e.target as HTMLElement;
    if (target.contentEditable === 'true') return;
    e.preventDefault();
    setAssignMenuCol({ colIndex, x: e.clientX, y: e.clientY });
  }, [onColumnAssign]);

  // Close assign menu on outside click
  useEffect(() => {
    if (!assignMenuCol) return;
    const handler = () => setAssignMenuCol(null);
    document.addEventListener('click', handler, { once: true });
    return () => document.removeEventListener('click', handler);
  }, [assignMenuCol]);

  const handleFieldAssign = useCallback((field: string) => {
    if (!assignMenuCol || !onColumnAssign) return;
    onColumnAssign(assignMenuCol.colIndex, field);
    setAssignMenuCol(null);
  }, [assignMenuCol, onColumnAssign]);

  // Get live group membership during resize
  const getLiveGroupId = useCallback((colIdx: number): ColumnGroupId | undefined => {
    for (const span of liveSpans) {
      if (colIdx >= span.startCol && colIdx <= span.endCol) return span.groupId;
    }
    return undefined;
  }, [liveSpans]);

  const tableWidth = maxCols * COL_WIDTH + ROW_NUM_WIDTH;

  return (
    <div className="relative flex flex-col h-full">
      {/* Instruction banner */}
      {onColumnAssign && (
        <div className="shrink-0 px-3 py-1.5 bg-muted/50 border-b border-panel-border text-[10px] font-mono text-muted-foreground">
          <span className="text-foreground/70 font-semibold">How to reassign:</span>{' '}
          Drag colored group edges ← → to include/exclude columns. Click column letter to assign a field. Double-click column name to rename. Right-click to clear.
        </div>
      )}

      {/* Single scroll container for group headers + grid */}
      <div ref={scrollRef} className="flex-1 overflow-auto" style={{ cursor: resizing ? 'col-resize' : undefined }}>
        <div style={{ minWidth: `${tableWidth}px` }}>
          {/* Group header band */}
          {liveSpans.length > 0 && (
            <div className="sticky top-0 z-30 flex h-7" style={{ paddingLeft: `${ROW_NUM_WIDTH}px` }}>
              <div className="relative w-full" style={{ minWidth: `${maxCols * COL_WIDTH}px` }}>
                {liveSpans.map(span => {
                  const colors = GROUP_COLORS[span.groupId];
                  const group = COLUMN_GROUPS.find(g => g.id === span.groupId);
                  return (
                    <div
                      key={span.groupId}
                      className={`absolute h-full flex items-center justify-center text-[10px] font-heading uppercase tracking-wider border-t-2 border-l-2 border-r-2 rounded-t-sm ${colors.border} ${colors.bg} ${colors.text}`}
                      style={{
                        left: `${span.startCol * COL_WIDTH}px`,
                        width: `${(span.endCol - span.startCol + 1) * COL_WIDTH}px`,
                      }}
                    >
                      {/* Left drag handle */}
                      {onGroupResize && (
                        <div
                          className="absolute left-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-foreground/10 z-10"
                          onMouseDown={(e) => handleResizeStart(span.groupId, 'left', e)}
                        />
                      )}
                      <span className="pointer-events-none select-none">{group?.label}</span>
                      {/* Right drag handle */}
                      {onGroupResize && (
                        <div
                          className="absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-foreground/10 z-10"
                          onMouseDown={(e) => handleResizeStart(span.groupId, 'right', e)}
                        />
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Table */}
          <table className="border-collapse text-[11px] font-mono" style={{ tableLayout: 'fixed', width: `${maxCols * COL_WIDTH + ROW_NUM_WIDTH}px` }}>
            <thead className="sticky z-20" style={{ top: liveSpans.length > 0 ? '28px' : '0px' }}>
              <tr className="bg-card">
                <th className="w-[40px] min-w-[40px] p-0 border-r border-b border-panel-border bg-card sticky left-0 z-30" />
                {Array.from({ length: maxCols }, (_, c) => {
                  const groupId = getLiveGroupId(c);
                  const fieldInfo = colFieldMap.get(c);
                  const colors = groupId ? GROUP_COLORS[groupId] : null;
                  const alias = columnAliases[c];
                  return (
                    <th
                      key={c}
                      className={`min-w-[100px] max-w-[180px] p-1 border-r border-b border-panel-border text-center select-none transition-colors ${
                        colors ? `${colors.bg} border-b-2 ${colors.border}` : 'bg-card'
                      } ${onColumnAssign ? 'cursor-pointer hover:bg-muted/50' : ''}`}
                      style={{ width: `${COL_WIDTH}px` }}
                      onClick={(e) => handleColHeaderClick(c, e)}
                      onContextMenu={(e) => {
                        if (!onColumnAssign) return;
                        e.preventDefault();
                        onColumnAssign(c, '');
                      }}
                    >
                      {/* Column letter — click to assign */}
                      <div className="text-muted-foreground text-[10px]">
                        {indexToColLetter(c)}
                      </div>

                      {/* Editable column name alias */}
                      {onColumnRename && (
                        <div
                          contentEditable
                          suppressContentEditableWarning
                          spellCheck={false}
                          title="Double-click to rename column"
                          className={`text-[9px] outline-none truncate cursor-text rounded px-0.5 hover:bg-foreground/5 focus:bg-foreground/10 focus:ring-1 focus:ring-foreground/20 ${
                            alias ? 'text-foreground/70' : 'text-muted-foreground/40'
                          }`}
                          onFocus={(e) => {
                            // Select all on focus for easy overwrite
                            const el = e.currentTarget;
                            const range = document.createRange();
                            range.selectNodeContents(el);
                            const sel = window.getSelection();
                            sel?.removeAllRanges();
                            sel?.addRange(range);
                          }}
                          onBlur={(e) => {
                            onColumnRename(c, e.currentTarget.textContent || '');
                          }}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              e.currentTarget.blur();
                            }
                            if (e.key === 'Escape') {
                              // Restore previous value on escape
                              e.currentTarget.textContent = alias || '';
                              e.currentTarget.blur();
                            }
                            // Stop click-to-assign from firing while editing
                            e.stopPropagation();
                          }}
                          onClick={(e) => e.stopPropagation()}
                        >
                          {alias || ''}
                        </div>
                      )}

                      {/* Field label from column_map */}
                      {fieldInfo && (
                        <div className={`text-[9px] ${fieldInfo.groupId === 'custom' ? 'text-accent-foreground' : GROUP_COLORS[fieldInfo.groupId].text} truncate`}>
                          {fieldInfo.fieldLabel}
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
                    <td className="w-[40px] min-w-[40px] px-1 py-0.5 text-right text-[10px] text-muted-foreground border-r border-b border-panel-border bg-card sticky left-0 z-10 select-none">
                      {rowIdx + 1}
                    </td>
                    {Array.from({ length: maxCols }, (_, c) => {
                      const val = c < row.length ? row[c] : null;
                      const groupId = getLiveGroupId(c);
                      const colors = groupId ? GROUP_COLORS[groupId] : null;
                      return (
                        <td
                          key={c}
                          className={`min-w-[100px] max-w-[180px] px-1.5 py-0.5 border-r border-b border-panel-border/50 truncate ${
                            colors ? colors.bg : ''
                          } ${isHeader ? 'font-semibold text-foreground' : 'text-foreground/80'} ${
                            colors ? `border-l ${colors.border}/30` : ''
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
      </div>

      {/* Assign menu */}
      {assignMenuCol && (
        <div
          className="fixed z-50 bg-card border border-panel-border rounded-sm shadow-lg py-1 min-w-[200px] max-h-[400px] overflow-y-auto"
          style={{ left: assignMenuCol.x, top: assignMenuCol.y }}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="px-2 py-1 text-[10px] text-muted-foreground font-heading uppercase tracking-wider border-b border-panel-border">
            Assign Column {indexToColLetter(assignMenuCol.colIndex)}
          </div>
          {COLUMN_GROUPS.map(group => (
            <div key={group.id}>
              <div className={`px-2 pt-1.5 pb-0.5 text-[9px] font-heading uppercase tracking-wider ${GROUP_COLORS[group.id].text}`}>
                {group.label}
              </div>
              {group.fields.map(field => (
                <button
                  key={field}
                  className="w-full text-left px-3 py-1 text-[11px] font-mono text-foreground/80 hover:bg-muted transition-colors"
                  onClick={() => handleFieldAssign(field)}
                >
                  {group.fieldLabels[field]}
                </button>
              ))}
            </div>
          ))}

          {/* Existing custom fields */}
          {instruction?.custom_columns && Object.keys(instruction.custom_columns).length > 0 && (
            <div className="border-t border-panel-border mt-1">
              <div className="px-2 pt-1.5 pb-0.5 text-[9px] font-heading uppercase tracking-wider text-accent-foreground">
                Custom Fields
              </div>
              {Object.keys(instruction.custom_columns).map(fieldName => (
                <button
                  key={fieldName}
                  className="w-full text-left px-3 py-1 text-[11px] font-mono text-foreground/80 hover:bg-muted transition-colors"
                  onClick={() => {
                    if (onCustomFieldAssign) onCustomFieldAssign(assignMenuCol.colIndex, fieldName);
                    setAssignMenuCol(null);
                  }}
                >
                  {fieldName}
                </button>
              ))}
            </div>
          )}

          {/* Add new custom field */}
          <div className="border-t border-panel-border mt-1 px-2 py-1.5">
            <div className="text-[9px] font-heading uppercase tracking-wider text-muted-foreground mb-1">
              Add Custom Field
            </div>
            <form
              className="flex gap-1"
              onSubmit={(e) => {
                e.preventDefault();
                const name = customFieldInput.trim();
                if (name && onCustomFieldAssign) {
                  onCustomFieldAssign(assignMenuCol.colIndex, name);
                  setCustomFieldInput('');
                  setAssignMenuCol(null);
                }
              }}
            >
              <input
                type="text"
                value={customFieldInput}
                onChange={(e) => setCustomFieldInput(e.target.value)}
                placeholder="e.g. code, status..."
                className="flex-1 px-1.5 py-0.5 text-[11px] font-mono bg-muted border border-panel-border rounded-sm outline-none focus:ring-1 focus:ring-foreground/20"
                autoFocus
                onClick={(e) => e.stopPropagation()}
              />
              <button
                type="submit"
                className="px-2 py-0.5 text-[10px] font-mono bg-accent text-accent-foreground rounded-sm hover:opacity-90"
              >
                Add
              </button>
            </form>
          </div>

          <div className="border-t border-panel-border mt-1">
            <button
              className="w-full text-left px-3 py-1 text-[11px] font-mono text-destructive hover:bg-muted transition-colors"
              onClick={() => handleFieldAssign('')}
            >
              Clear Assignment
            </button>
          </div>
        </div>
      )}
    </div>
  );
}