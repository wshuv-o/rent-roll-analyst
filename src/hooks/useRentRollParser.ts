import { useState, useCallback, useRef } from 'react';
import type { LogEntry, LogType, TenantObject, ParsingInstruction, WorkflowStep, GroupSpan, ColumnGroupId } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';
import { readExcelFile, formatFileSize } from '@/lib/excel-utils';
import { anonymizeSheet, detectHeaderRows } from '@/lib/anonymizer';
import { buildSample } from '@/lib/sample-builder';
import { parseSheet } from '@/lib/parser';
import { streamAnalysis } from '@/lib/ai-stream';

let logIdCounter = 0;

function indexToColLetter(idx: number): string {
  let letter = '';
  let n = idx;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

function colLetterToIdx(letter: string): number {
  if (!letter) return -1;
  const upper = letter.toUpperCase().trim().replace(/[^A-Z]/g, '');
  if (!upper) return -1;
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

// Derive group spans from instruction's column_map
function deriveGroupSpans(instruction: ParsingInstruction): GroupSpan[] {
  const spans: GroupSpan[] = [];
  for (const group of COLUMN_GROUPS) {
    const indices: number[] = [];
    for (const field of group.fields) {
      const letter = instruction.column_map[field];
      if (letter) {
        const idx = colLetterToIdx(letter);
        if (idx >= 0) indices.push(idx);
      }
    }
    if (indices.length > 0) {
      spans.push({
        groupId: group.id,
        startCol: Math.min(...indices),
        endCol: Math.max(...indices),
      });
    }
  }
  return spans.sort((a, b) => a.startCol - b.startCol);
}

export function useRentRollParser() {
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [tenants, setTenants] = useState<TenantObject[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileName, setFileName] = useState('');
  const [step, setStep] = useState<WorkflowStep>('upload');

  const [sheetData, setSheetData] = useState<(string | number | null)[][]>([]);
  const [headerRows, setHeaderRows] = useState<number[]>([]);
  const [instruction, setInstruction] = useState<ParsingInstruction | null>(null);
  const [groupSpans, setGroupSpans] = useState<GroupSpan[]>([]);
  const [totalRows, setTotalRows] = useState(0);

  // Column aliases: colIndex → custom display name
  const [columnAliases, setColumnAliases] = useState<Record<number, string>>({});

  const sheetDataRef = useRef<(string | number | null)[][]>([]);
  const streamingEntryRef = useRef<string | null>(null);

  // NOTE: No useEffect deriving groupSpans from instruction.
  // groupSpans are set explicitly after AI response or reset.
  // This prevents drag-resize snapping back.

  const addLog = useCallback((type: LogType, message: string, isStreaming = false): string => {
    const id = `log-${++logIdCounter}`;
    setLogs(prev => [...prev, { id, type, message, timestamp: new Date(), isStreaming }]);
    return id;
  }, []);

  const updateLog = useCallback((id: string, update: Partial<LogEntry>) => {
    setLogs(prev => prev.map(l => l.id === id ? { ...l, ...update } : l));
  }, []);

  const appendToLog = useCallback((id: string, text: string) => {
    setLogs(prev => prev.map(l => l.id === id ? { ...l, message: l.message + text } : l));
  }, []);

  const loadFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setTenants([]);
    setInstruction(null);
    setGroupSpans([]);
    setColumnAliases({});
    setFileName(file.name);
    setStep('analyzing');

    addLog('system', `File received: ${file.name} — ${formatFileSize(file.size)} — reading sheet...`);

    let data: (string | number | null)[][];
    let rows: number;

    try {
      const result = await readExcelFile(file);
      data = result.data;
      rows = result.totalRows;
    } catch (err) {
      addLog('flag', `Failed to read file: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setIsProcessing(false);
      setStep('upload');
      return;
    }

    sheetDataRef.current = data;
    setSheetData(data);
    setTotalRows(rows);

    const headerRowIndices = detectHeaderRows(data);
    setHeaderRows(headerRowIndices);
    addLog('system', `${rows} rows loaded. Detected header rows: ${headerRowIndices.map(i => i + 1).join(', ')}`);

    const { anonymized, stats } = anonymizeSheet(data, headerRowIndices);
    addLog('system', `Anonymized sample: ${stats.names} names, ${stats.suites} suite IDs masked.`);

    const { html, contextNote, sampleRanges } = buildSample(anonymized, rows);
    addLog('system', `Sample: ${sampleRanges}. Sending to AI...`);

    let instructionJson: ParsingInstruction | null = null;

    await new Promise<void>((resolve) => {
      let currentStreamId: string | null = null;

      streamAnalysis(html, contextNote, {
        onSection: (type: LogType) => {
          if (currentStreamId) updateLog(currentStreamId, { isStreaming: false });
          currentStreamId = addLog(type, '', true);
          streamingEntryRef.current = currentStreamId;
        },
        onToken: (_type: LogType, token: string) => {
          if (currentStreamId) {
            appendToLog(currentStreamId, token);
          } else {
            currentStreamId = addLog('thinking', token, true);
            streamingEntryRef.current = currentStreamId;
          }
        },
        onInstruction: (json: string) => {
          try {
            instructionJson = JSON.parse(json) as ParsingInstruction;
            addLog('output', `Parsing instruction received. Confidence: ${instructionJson.confidence}.`);
          } catch {
            addLog('flag', 'Failed to parse AI instruction JSON.');
          }
        },
        onDone: () => {
          if (currentStreamId) updateLog(currentStreamId, { isStreaming: false });
          resolve();
        },
        onError: (error: string) => {
          addLog('flag', error);
          resolve();
        },
      });
    });

    if (!instructionJson) {
      addLog('flag', 'No parsing instruction received. You can manually assign columns.');
      setStep('confirm');
      setIsProcessing(false);
      return;
    }

    console.log('[HOOK] AI instruction:', JSON.stringify(instructionJson, null, 2));
    setInstruction(instructionJson);
    // Set spans explicitly here — NOT via useEffect — to avoid resize snap-back
    setGroupSpans(deriveGroupSpans(instructionJson));
    setStep('confirm');
    setIsProcessing(false);
    addLog('system', 'Review the highlighted columns. Drag group edges to adjust, click column headers to assign fields, then "Confirm & Parse".');
  }, [addLog, updateLog, appendToLog]);

  // Handle group span resize (dragging edges).
  // Also auto-assigns newly included columns to unassigned fields in that group.
  const handleGroupResize = useCallback((groupId: ColumnGroupId, newStartCol: number, newEndCol: number) => {
    // Update visual spans immediately
    setGroupSpans(prev => {
      const updated = prev.map(s =>
        s.groupId === groupId ? { ...s, startCol: newStartCol, endCol: newEndCol } : s
      );
      return updated.sort((a, b) => a.startCol - b.startCol);
    });

    // Update column_map:
    // 1. Remove fields outside new range
    // 2. Auto-assign newly included columns to unassigned fields in this group
    setInstruction(prev => {
      if (!prev) return prev;
      const group = COLUMN_GROUPS.find(g => g.id === groupId);
      if (!group) return prev;

      const newMap = { ...prev.column_map } as Record<string, string>;

      // Step 1: Remove fields that fell outside the new range
      for (const field of group.fields) {
        const idx = colLetterToIdx(newMap[field] || '');
        if (idx >= 0 && (idx < newStartCol || idx > newEndCol)) {
          newMap[field] = '';
        }
      }

      // Step 2: Collect which cols in the new range are already assigned within this group
      const assignedColsInGroup = new Set(
        group.fields
          .map(f => colLetterToIdx(newMap[f] || ''))
          .filter(i => i >= 0)
      );

      // Step 3: Find fields in this group that still have no assignment
      const unassignedFields = group.fields.filter(f => !newMap[f]);

      // Step 4: For each col in new range not yet assigned to anything globally,
      // bind it to the next unassigned field in this group
      let fieldCursor = 0;
      for (let col = newStartCol; col <= newEndCol; col++) {
        if (fieldCursor >= unassignedFields.length) break;
        if (assignedColsInGroup.has(col)) continue;

        // Check col isn't assigned to any field across ALL groups
        const usedGlobally = Object.values(newMap).some(v => colLetterToIdx(v) === col);
        if (!usedGlobally) {
          newMap[unassignedFields[fieldCursor]] = indexToColLetter(col);
          fieldCursor++;
        }
      }

      return { ...prev, column_map: newMap as ParsingInstruction['column_map'] };
    });
  }, []);

  // Column field assignment via click menu
  const handleColumnAssign = useCallback((colIndex: number, field: string) => {
    setInstruction(prev => {
      if (!prev) {
        // No instruction yet — create a blank one
        const blank: ParsingInstruction = {
          header_rows: headerRows,
          data_starts_at_row: (headerRows.length > 0 ? headerRows[headerRows.length - 1] + 2 : 1),
          column_map: {
            suite_id: '', tenant_name: '', lease_start: '', lease_end: '',
            gla_sqft: '', monthly_base_rent: '', base_rent_psf: '',
            recurring_charge_code: '', recurring_charge_amount: '', recurring_charge_psf: '',
            future_rent_date: '', future_rent_amount: '', future_rent_psf: '',
          },
          new_tenant_rule: 'suite_id column non-empty',
          skip_row_patterns: [],
          addon_space_patterns: [],
          confidence: 'medium',
          notes: 'Manual assignment',
        };
        if (field) {
          (blank.column_map as Record<string, string>)[field] = indexToColLetter(colIndex);
        }
        return blank;
      }

      const newMap = { ...prev.column_map } as Record<string, string>;

      if (field) {
        // Clear any other field that was pointing to this column
        for (const key of Object.keys(newMap)) {
          if (colLetterToIdx(newMap[key]) === colIndex) {
            newMap[key] = '';
          }
        }
        newMap[field] = indexToColLetter(colIndex);
      } else {
        // Clear assignment — field is empty string means "unassign this col"
        for (const key of Object.keys(newMap)) {
          if (colLetterToIdx(newMap[key]) === colIndex) {
            newMap[key] = '';
          }
        }
      }

      // Re-derive group spans from the updated map so colored bands stay in sync
      const updatedInstruction = { ...prev, column_map: newMap as ParsingInstruction['column_map'] };
      setGroupSpans(deriveGroupSpans(updatedInstruction));

      return updatedInstruction;
    });
  }, [headerRows]);

  // Rename a column header for display (does not affect column_map keys)
  const handleColumnRename = useCallback((colIndex: number, name: string) => {
    setColumnAliases(prev => {
      if (!name.trim()) {
        // Remove alias if cleared
        const next = { ...prev };
        delete next[colIndex];
        return next;
      }
      return { ...prev, [colIndex]: name.trim() };
    });
  }, []);

  const confirmAndParse = useCallback(() => {
    if (!instruction) {
      addLog('flag', 'No column mapping defined. Please assign columns first.');
      return;
    }

    setStep('parsing');
    setIsProcessing(true);
    addLog('system', `Parsing full sheet... ${totalRows} rows.`);

    const data = sheetDataRef.current;
    const finalTenants = parseSheet(data, instruction, addLog);
    addLog('system', `${finalTenants.length} tenant blocks found.`);

    setTenants(finalTenants);
    setStep('done');
    setIsProcessing(false);
  }, [instruction, totalRows, addLog]);

  const resetToUpload = useCallback(() => {
    setStep('upload');
    setSheetData([]);
    setInstruction(null);
    setGroupSpans([]);
    setTenants([]);
    setHeaderRows([]);
    setLogs([]);
    setColumnAliases({});
  }, []);

  const reAnalyze = useCallback(() => {
    setInstruction(null);
    setGroupSpans([]);
    setTenants([]);
    setColumnAliases({});
    setStep('analyzing');
    const data = sheetDataRef.current;
    if (data.length === 0) return;

    setIsProcessing(true);
    const headerRowIndices = detectHeaderRows(data);
    const { anonymized } = anonymizeSheet(data, headerRowIndices);
    const { html, contextNote, sampleRanges } = buildSample(anonymized, totalRows);
    addLog('system', `Re-analyzing... Sample: ${sampleRanges}`);

    streamAnalysis(html, contextNote, {
      onSection: (type: LogType) => {
        addLog(type, '', true);
      },
      onToken: (_type: LogType, _token: string) => {},
      onInstruction: (json: string) => {
        try {
          const instructionJson = JSON.parse(json) as ParsingInstruction;
          setInstruction(instructionJson);
          // Set spans explicitly — NOT via useEffect
          setGroupSpans(deriveGroupSpans(instructionJson));
          addLog('output', `New instruction received. Confidence: ${instructionJson.confidence}.`);
        } catch {
          addLog('flag', 'Failed to parse AI instruction JSON.');
        }
      },
      onDone: () => {
        setStep('confirm');
        setIsProcessing(false);
        addLog('system', 'Review updated column mapping.');
      },
      onError: (error: string) => {
        addLog('flag', error);
        setStep('confirm');
        setIsProcessing(false);
      },
    });
  }, [totalRows, addLog]);

  return {
    logs, tenants, isProcessing, fileName, step,
    sheetData, headerRows, instruction, groupSpans,
    columnAliases,
    loadFile, handleColumnAssign, handleGroupResize,
    handleColumnRename,
    confirmAndParse, resetToUpload, reAnalyze,
  };
}