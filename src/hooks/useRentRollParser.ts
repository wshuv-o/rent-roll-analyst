import { useState, useCallback, useRef } from 'react';
import type { LogEntry, LogType, TenantObject, ParsingInstruction, AnonymizationMapping, WorkflowStep } from '@/lib/types';
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

export function useRentRollParser() {
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [tenants, setTenants] = useState<TenantObject[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileName, setFileName] = useState('');
  const [step, setStep] = useState<WorkflowStep>('upload');

  // Raw sheet data for spreadsheet viewer
  const [sheetData, setSheetData] = useState<(string | number | null)[][]>([]);
  const [headerRows, setHeaderRows] = useState<number[]>([]);
  const [instruction, setInstruction] = useState<ParsingInstruction | null>(null);
  const [totalRows, setTotalRows] = useState(0);

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

  // Step 1: Load file and show in spreadsheet
  const loadFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setTenants([]);
    setInstruction(null);
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

    // Detect headers
    const headerRowIndices = detectHeaderRows(data);
    setHeaderRows(headerRowIndices);
    addLog('system', `${rows} rows loaded. Detected header rows: ${headerRowIndices.map(i => i + 1).join(', ')}`);

    // Anonymize for AI sample only
    const { anonymized, stats } = anonymizeSheet(data, headerRowIndices);
    addLog('system', `Anonymized sample: ${stats.names} names, ${stats.suites} suite IDs masked.`);

    // Build sample & send to AI
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
    setStep('confirm');
    setIsProcessing(false);
    addLog('system', 'Review the highlighted columns. Drag-select columns to reassign, then click "Confirm & Parse".');
  }, [addLog, updateLog, appendToLog]);

  // Column reassignment handler
  const handleColumnAssign = useCallback((colIndex: number, field: string) => {
    setInstruction(prev => {
      if (!prev) {
        // Create a blank instruction
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

      const newMap = { ...prev.column_map };
      // Clear any existing assignment for this field
      if (field) {
        // Clear previous column that had this field
        for (const [key, val] of Object.entries(newMap)) {
          const existingIdx = colLetterToIdx(val);
          if (existingIdx === colIndex) {
            (newMap as Record<string, string>)[key] = '';
          }
        }
        (newMap as Record<string, string>)[field] = indexToColLetter(colIndex);
      } else {
        // Clearing: find what was assigned to this column
        for (const [key, val] of Object.entries(newMap)) {
          const existingIdx = colLetterToIdx(val);
          if (existingIdx === colIndex) {
            (newMap as Record<string, string>)[key] = '';
          }
        }
      }

      return { ...prev, column_map: newMap };
    });
  }, [headerRows]);

  // Step 2: Confirm and parse
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

  // Re-analyze: go back to upload
  const resetToUpload = useCallback(() => {
    setStep('upload');
    setSheetData([]);
    setInstruction(null);
    setTenants([]);
    setHeaderRows([]);
    setLogs([]);
  }, []);

  // Re-analyze with same file
  const reAnalyze = useCallback(() => {
    // Reset to confirm state but trigger new AI analysis
    setInstruction(null);
    setTenants([]);
    setStep('analyzing');
    // Re-run the AI part with existing data
    const data = sheetDataRef.current;
    if (data.length === 0) return;

    setIsProcessing(true);
    const headerRowIndices = detectHeaderRows(data);
    const { anonymized } = anonymizeSheet(data, headerRowIndices);
    const { html, contextNote, sampleRanges } = buildSample(anonymized, totalRows);
    addLog('system', `Re-analyzing... Sample: ${sampleRanges}`);

    let instructionJson: ParsingInstruction | null = null;

    streamAnalysis(html, contextNote, {
      onSection: (type: LogType) => {
        addLog(type, '', true);
      },
      onToken: (_type: LogType, token: string) => {
        // simplified for re-analysis
      },
      onInstruction: (json: string) => {
        try {
          instructionJson = JSON.parse(json) as ParsingInstruction;
          setInstruction(instructionJson);
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
    logs,
    tenants,
    isProcessing,
    fileName,
    step,
    sheetData,
    headerRows,
    instruction,
    loadFile,
    handleColumnAssign,
    confirmAndParse,
    resetToUpload,
    reAnalyze,
  };
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
