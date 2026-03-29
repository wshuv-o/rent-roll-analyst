//src/hooks/useRentRollParser.ts
import { useState, useCallback, useRef } from 'react';
import type { LogEntry, LogType, TenantObject, ParsingInstruction, WorkflowStep, GroupSpan, ColumnGroupId, CustomGroup } from '@/lib/types';
import { COLUMN_GROUPS } from '@/lib/types';
import { readExcelFile, formatFileSize } from '@/lib/excel-utils';
import { anonymizeSheet, detectHeaderRows } from '@/lib/anonymizer';
import { buildSample } from '@/lib/sample-builder';
import { parseSheet } from '@/lib/parser';
import { streamAnalysis } from '@/lib/ai-stream';
import { indexToColLetter, colLetterToIndex } from '@/lib/col-utils';
import { parseTenancySchedule } from '@/lib/rent-roll-types/tenancy-schedule-parser';
import type { TenancyScheduleTenant } from '@/lib/rent-roll-types/tenancy-schedule-parser';

let logIdCounter = 0;

const colLetterToIdx = colLetterToIndex;

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
        collection: group.collection,
      });
    }
  }
  return spans.sort((a, b) => a.startCol - b.startCol);
}

/** Compute the last non-empty column across the first N rows */
function computeEffectiveCols(data: (string | number | null)[][], maxRows: number): number {
  let effCols = 0;
  const limit = Math.min(maxRows, data.length);
  for (let r = 0; r < limit; r++) {
    const row = data[r];
    for (let c = row.length - 1; c >= 0; c--) {
      if (row[c] !== null && row[c] !== undefined && String(row[c]).trim() !== '') {
        effCols = Math.max(effCols, c + 1);
        break;
      }
    }
  }
  return effCols;
}

const MAX_SAMPLE_ROWS = 50;
const MAX_SAMPLE_COLS = 100;

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

  // Sample selection bounds
  const [sampleRows, setSampleRows] = useState(15);
  const [sampleCols, setSampleCols] = useState(0);
  const [maxAvailableCols, setMaxAvailableCols] = useState(0);

  // Column aliases: colIndex → custom display name
  const [columnAliases, setColumnAliases] = useState<Record<number, string>>({});
  // Custom groups created by user
  const [customGroups, setCustomGroups] = useState<CustomGroup[]>([]);

  const [sentSampleHtml, setSentSampleHtml] = useState<string | null>(null);
  const [selectedRentRollType, setSelectedRentRollType] = useState<string | null>(null);

  const sheetDataRef = useRef<(string | number | null)[][]>([]);
  const streamingEntryRef = useRef<string | null>(null);

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

  // Step 1: Load file → go to review-sample
  const loadFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setTenants([]);
    setInstruction(null);
    setGroupSpans([]);
    setColumnAliases({});
    setCustomGroups([]);
    setFileName(file.name);

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

    // Compute effective column count (last non-empty column across first 50 rows)
    const effCols = computeEffectiveCols(data, MAX_SAMPLE_ROWS);
    const totalCols = data.reduce((max, row) => Math.max(max, row.length), 0);
    setMaxAvailableCols(Math.min(totalCols, MAX_SAMPLE_COLS));
    setSampleRows(Math.min(15, rows));
    setSampleCols(Math.min(effCols, MAX_SAMPLE_COLS));

    setStep('type-confirm');
    setIsProcessing(false);
  }, [addLog]);

  // Tenancy schedule result (separate from regular TenantObject[])
  const [tenancyScheduleTenants, setTenancyScheduleTenants] = useState<TenancyScheduleTenant[]>([]);

  // Step 1b: User selects rent roll type → branch by type
  const confirmRentRollType = useCallback((typeId: string) => {
    setSelectedRentRollType(typeId);

    if (typeId === 'tenancy-schedule') {
      // Tenancy schedule: skip AI, parse directly with hardcoded rules
      addLog('system', `Rent roll type selected: Tenancy Schedule. Parsing with built-in rules...`);
      setStep('parsing');
      setIsProcessing(true);

      const data = sheetDataRef.current;
      const result = parseTenancySchedule(data, addLog);
      setTenancyScheduleTenants(result);

      // Convert to TenantObject[] for the existing TenantTable display.
      // Main row keys are dynamic (from merged headers), so find the right ones.
      const converted: TenantObject[] = result.map(t => {
        const mr = t.mainRow;
        // Find the unit and tenant name by looking for common header labels
        const unitKey = Object.keys(mr).find(k => /unit/i.test(k)) || '';
        const leaseKey = Object.keys(mr).find(k => /lease(?!.*type|.*from|.*to)/i.test(k)) || '';
        return {
          suite_id: String(mr[unitKey] ?? ''),
          tenant_name: String(mr[leaseKey] ?? ''),
          rawRows: t.rawRows,
          notes: '',
        };
      });
      setTenants(converted);

      setStep('done');
      setIsProcessing(false);
    } else {
      // Regular: go through AI flow
      addLog('system', `Rent roll type selected: ${typeId}. Adjust sample area then click "Send to AI".`);
      setStep('review-sample');
    }
  }, [addLog]);

  // Step 2: User confirms sample bounds → send to AI
  const sendSampleToAI = useCallback(async () => {
    setStep('analyzing');
    setIsProcessing(true);

    const data = sheetDataRef.current;
    const { anonymized, stats } = anonymizeSheet(data, headerRows);
    addLog('system', `Anonymized sample: ${stats.names} names, ${stats.suites} suite IDs masked.`);

    const { html, contextNote, sampleRanges } = buildSample(anonymized, totalRows, sampleRows, sampleCols);
    setSentSampleHtml(html);
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
    setGroupSpans(deriveGroupSpans(instructionJson));
    setStep('confirm');
    setIsProcessing(false);
    addLog('system', 'Review the highlighted columns. Drag group edges to adjust, click column headers to assign fields, then "Confirm & Parse".');
  }, [headerRows, totalRows, sampleRows, sampleCols, addLog, updateLog, appendToLog]);

  // Handle group span resize (dragging edges).
  const handleGroupResize = useCallback((groupId: ColumnGroupId | string, newStartCol: number, newEndCol: number) => {
    setGroupSpans(prev => {
      const updated = prev.map(s =>
        s.groupId === groupId ? { ...s, startCol: newStartCol, endCol: newEndCol } : s
      );
      return updated.sort((a, b) => a.startCol - b.startCol);
    });
  }, []);

  // Column field assignment via click menu
  const handleColumnAssign = useCallback((colIndex: number, field: string) => {
    setInstruction((prev: ParsingInstruction | null) => {
      if (!prev) {
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
        for (const key of Object.keys(newMap)) {
          if (colLetterToIdx(newMap[key]) === colIndex) {
            newMap[key] = '';
          }
        }
        newMap[field] = indexToColLetter(colIndex);
      } else {
        for (const key of Object.keys(newMap)) {
          if (colLetterToIdx(newMap[key]) === colIndex) {
            newMap[key] = '';
          }
        }
        const newCustom = { ...(prev.custom_columns || {}) };
        for (const key of Object.keys(newCustom)) {
          if (colLetterToIdx(newCustom[key]) === colIndex) {
            delete newCustom[key];
          }
        }
        const updatedInstruction = { ...prev, column_map: newMap as ParsingInstruction['column_map'], custom_columns: newCustom };
        setGroupSpans(deriveGroupSpans(updatedInstruction));
        return updatedInstruction;
      }

      const updatedInstruction = { ...prev, column_map: newMap as ParsingInstruction['column_map'] };
      setGroupSpans(deriveGroupSpans(updatedInstruction));
      return updatedInstruction;
    });
  }, [headerRows]);

  const handleCustomFieldAssign = useCallback((colIndex: number, fieldName: string) => {
    setInstruction((prev: ParsingInstruction | null) => {
      const base = prev || {
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
        confidence: 'medium' as const,
        notes: 'Manual assignment',
        custom_columns: {},
      };

      const newCustom = { ...(base.custom_columns || {}) };
      for (const key of Object.keys(newCustom)) {
        if (colLetterToIdx(newCustom[key]) === colIndex) {
          delete newCustom[key];
        }
      }
      const newMap = { ...base.column_map } as Record<string, string>;
      for (const key of Object.keys(newMap)) {
        if (colLetterToIdx(newMap[key]) === colIndex) {
          newMap[key] = '';
        }
      }
      newCustom[fieldName] = indexToColLetter(colIndex);

      return { ...base, column_map: newMap as ParsingInstruction['column_map'], custom_columns: newCustom };
    });
  }, [headerRows]);

  const handleColumnRename = useCallback((colIndex: number, name: string) => {
    setColumnAliases(prev => {
      if (!name.trim()) {
        const next = { ...prev };
        delete next[colIndex];
        return next;
      }
      return { ...prev, [colIndex]: name.trim() };
    });
  }, []);

  const handleCreateCustomGroup = useCallback((colIndex: number, groupName: string, collection: boolean) => {
    const groupId = `custom-${groupName.toLowerCase().replace(/\s+/g, '-')}-${Date.now()}`;
    const newGroup: CustomGroup = { id: groupId, label: groupName, collection };
    setCustomGroups(prev => [...prev, newGroup]);
    setGroupSpans(prev => [
      ...prev,
      { groupId, startCol: colIndex, endCol: colIndex, collection },
    ].sort((a, b) => a.startCol - b.startCol));
  }, []);

  const buildColumnLabels = useCallback((): Record<number, string> => {
    const labels: Record<number, string> = {};
    const data = sheetDataRef.current;
    const lastHeaderIdx = headerRows.length > 0 ? headerRows[headerRows.length - 1] : -1;
    const headerRow = lastHeaderIdx >= 0 && lastHeaderIdx < data.length ? data[lastHeaderIdx] : null;

    const maxCols = data.reduce((max, row) => Math.max(max, row.length), 0);
    for (let col = 0; col < maxCols; col++) {
      if (columnAliases[col]) {
        labels[col] = columnAliases[col];
      } else if (headerRow && col < headerRow.length && headerRow[col] !== null && headerRow[col] !== undefined) {
        labels[col] = String(headerRow[col]).trim();
      } else {
        labels[col] = indexToColLetter(col);
      }
    }
    return labels;
  }, [columnAliases, headerRows]);

  const confirmAndParse = useCallback(() => {
    if (!instruction) {
      addLog('flag', 'No column mapping defined. Please assign columns first.');
      return;
    }

    setStep('parsing');
    setIsProcessing(true);

    addLog('system', `Parsing full sheet... ${totalRows} rows, ${groupSpans.length} groups.`);

    const data = sheetDataRef.current;
    const finalTenants = parseSheet(data, instruction, addLog);
    addLog('system', `${finalTenants.length} tenant blocks found.`);

    setTenants(finalTenants);
    setStep('done');
    setIsProcessing(false);
  }, [instruction, totalRows, addLog, groupSpans, buildColumnLabels]);

  const resetToUpload = useCallback(() => {
    setStep('upload');
    setSheetData([]);
    setInstruction(null);
    setGroupSpans([]);
    setTenants([]);
    setHeaderRows([]);
    setLogs([]);
    setColumnAliases({});
    setCustomGroups([]);
    setSentSampleHtml(null);
    setSelectedRentRollType(null);
    setTenancyScheduleTenants([]);
  }, []);

  const reAnalyze = useCallback(() => {
    setInstruction(null);
    setGroupSpans([]);
    setTenants([]);
    setColumnAliases({});
    setCustomGroups([]);
    // Go back to review-sample so user can adjust bounds before re-sending
    setStep('review-sample');
  }, []);

  const goBackToConfirm = useCallback(() => {
    setStep('confirm');
  }, []);

  return {
    logs, tenants, isProcessing, fileName, step,
    sheetData, headerRows, instruction, groupSpans,
    columnAliases, customGroups, sentSampleHtml,
    sampleRows, sampleCols, maxAvailableCols, totalRows,
    setSampleRows, setSampleCols,
    selectedRentRollType, tenancyScheduleTenants,
    loadFile, sendSampleToAI, confirmRentRollType,
    handleColumnAssign, handleCustomFieldAssign, handleGroupResize,
    handleColumnRename, handleCreateCustomGroup,
    confirmAndParse, resetToUpload, reAnalyze, goBackToConfirm,
    buildColumnLabels,
  };
}
