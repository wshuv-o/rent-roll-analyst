import { useState, useCallback, useRef } from 'react';
import type { LogEntry, LogType, TenantObject, ParsingInstruction, AnonymizationMapping } from '@/lib/types';
import { readExcelFile, formatFileSize } from '@/lib/excel-utils';
import { anonymizeSheet, deanonymize } from '@/lib/anonymizer';
import { buildSample } from '@/lib/sample-builder';
import { parseSheet } from '@/lib/parser';
import { streamAnalysis } from '@/lib/ai-stream';

let logIdCounter = 0;

export function useRentRollParser() {
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [tenants, setTenants] = useState<TenantObject[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileName, setFileName] = useState('');

  // Store refs for data that persists across the flow
  const sheetDataRef = useRef<(string | number | null)[][]>([]);
  const mappingRef = useRef<AnonymizationMapping | null>(null);
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

  const processFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setTenants([]);
    setFileName(file.name);

    // Step 1 — Read file
    addLog('system', `File received: ${file.name} — ${formatFileSize(file.size)} — reading sheet...`);

    let data: (string | number | null)[][];
    let totalRows: number;

    try {
      const result = await readExcelFile(file);
      data = result.data;
      totalRows = result.totalRows;
      sheetDataRef.current = data;
    } catch (err) {
      addLog('flag', `Failed to read file: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setIsProcessing(false);
      return;
    }

    // Step 2 — Anonymization
    // Auto-detect header rows by scanning for header-like content
    const headerRowIndices = detectHeaderRows(data);
    addLog('system', `Detected header rows: ${headerRowIndices.map(i => i + 1).join(', ')}`);
    const { anonymized, mapping, stats } = anonymizeSheet(data, headerRowIndices);
    mappingRef.current = mapping;
    sheetDataRef.current = data; // Keep original for de-anonymization

    addLog('system',
      `Anonymization complete. ${stats.names} tenant names replaced. ${stats.amounts + stats.suites} values masked. Mapping stored in memory.`
    );

    // Step 3 — Sample
    const { html, contextNote, sampleRanges } = buildSample(anonymized, totalRows);
    addLog('system', `Sample prepared: ${sampleRanges}. Sending to AI agent...`);

    // Step 4 — AI Analysis (streaming)
    let instructionJson: ParsingInstruction | null = null;

    await new Promise<void>((resolve) => {
      let currentStreamId: string | null = null;

      streamAnalysis(html, contextNote, {
        onSection: (type: LogType) => {
          // Finalize previous streaming entry
          if (currentStreamId) {
            updateLog(currentStreamId, { isStreaming: false });
          }
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
            addLog('output', `Parsing instruction object received. Confidence: ${instructionJson.confidence}.`);
          } catch {
            addLog('flag', 'Failed to parse AI instruction JSON.');
          }
        },
        onDone: () => {
          if (currentStreamId) {
            updateLog(currentStreamId, { isStreaming: false });
          }
          resolve();
        },
        onError: (error: string) => {
          addLog('flag', error);
          resolve();
        },
      });
    });

    if (!instructionJson) {
      addLog('flag', 'No parsing instruction received from AI. Cannot proceed.');
      setIsProcessing(false);
      return;
    }

    // Check confidence
    if ((instructionJson as ParsingInstruction).confidence === 'low') {
      addLog('flag', 'AI is not confident about the layout. Review the flags before continuing.');
    }

    // Step 5 — Parse full sheet
    addLog('system', `Parsing full sheet... ${totalRows} rows processed.`);
    const parsedTenants = parseSheet(anonymized, instructionJson);
    addLog('system', `${parsedTenants.length} tenant blocks found.`);

    // Step 6 — De-anonymization
    const finalTenants = deanonymize(parsedTenants, mapping);
    addLog('system', 'De-anonymization complete. Real values restored from memory mapping.');

    setTenants(finalTenants);
    setIsProcessing(false);
  }, [addLog, updateLog, appendToLog]);

  return {
    logs,
    tenants,
    isProcessing,
    fileName,
    processFile,
  };
}
