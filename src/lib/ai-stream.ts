import type { LogType } from './types';

const CHAT_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/analyze-rent-roll`;

export interface StreamCallbacks {
  onSection: (type: LogType, text: string) => void;
  onToken: (type: LogType, token: string) => void;
  onInstruction: (json: string) => void;
  onDone: () => void;
  onError: (error: string) => void;
}

export async function streamAnalysis(
  sampleHtml: string,
  contextNote: string,
  callbacks: StreamCallbacks
): Promise<void> {
  const resp = await fetch(CHAT_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY}`,
    },
    body: JSON.stringify({ sampleHtml, contextNote }),
  });

  if (!resp.ok || !resp.body) {
    if (resp.status === 429) {
      callbacks.onError('Rate limit exceeded. Please wait a moment and try again.');
      return;
    }
    if (resp.status === 402) {
      callbacks.onError('AI usage credits exhausted. Please add credits to continue.');
      return;
    }
    callbacks.onError(`AI request failed (${resp.status})`);
    return;
  }

  const reader = resp.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let fullText = '';
  let currentSection: LogType = 'thinking';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    let newlineIndex: number;
    while ((newlineIndex = buffer.indexOf('\n')) !== -1) {
      let line = buffer.slice(0, newlineIndex);
      buffer = buffer.slice(newlineIndex + 1);

      if (line.endsWith('\r')) line = line.slice(0, -1);
      if (line.startsWith(':') || line.trim() === '') continue;
      if (!line.startsWith('data: ')) continue;

      const jsonStr = line.slice(6).trim();
      if (jsonStr === '[DONE]') {
        // Process any remaining full text for instruction extraction
        extractInstruction(fullText, callbacks);
        callbacks.onDone();
        return;
      }

      try {
        const parsed = JSON.parse(jsonStr);
        const content = parsed.choices?.[0]?.delta?.content as string | undefined;
        if (content) {
          fullText += content;
          // Detect section changes
          const sectionResult = detectSection(content, currentSection);
          if (sectionResult.newSection !== currentSection) {
            currentSection = sectionResult.newSection;
            callbacks.onSection(currentSection, '');
          }
          // Emit cleaned token
          const cleanedToken = cleanSectionTags(content);
          if (cleanedToken) {
            callbacks.onToken(currentSection, cleanedToken);
          }
        }
      } catch {
        buffer = line + '\n' + buffer;
        break;
      }
    }
  }

  // Final flush
  extractInstruction(fullText, callbacks);
  callbacks.onDone();
}

function detectSection(text: string, current: LogType): { newSection: LogType } {
  const upper = text.toUpperCase();
  if (upper.includes('[THINKING]')) return { newSection: 'thinking' };
  if (upper.includes('[GROUPING]')) return { newSection: 'grouping' };
  if (upper.includes('[PARSING INSTRUCTION]')) return { newSection: 'output' };
  if (upper.includes('[FLAGS]') || upper.includes('[FLAG]')) return { newSection: 'flag' };
  return { newSection: current };
}

function cleanSectionTags(text: string): string {
  return text
    .replace(/\[THINKING\]/gi, '')
    .replace(/\[GROUPING\]/gi, '')
    .replace(/\[PARSING INSTRUCTION\]/gi, '')
    .replace(/\[FLAGS?\]/gi, '');
}

function extractInstruction(fullText: string, callbacks: StreamCallbacks): void {
  // Try to find JSON object in the text
  const jsonMatch = fullText.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
  if (jsonMatch) {
    callbacks.onInstruction(jsonMatch[1]);
    return;
  }

  // Try to find raw JSON
  const braceStart = fullText.indexOf('"header_rows"');
  if (braceStart !== -1) {
    // Walk back to find opening brace
    let start = fullText.lastIndexOf('{', braceStart);
    if (start !== -1) {
      let depth = 0;
      let end = start;
      for (let i = start; i < fullText.length; i++) {
        if (fullText[i] === '{') depth++;
        if (fullText[i] === '}') depth--;
        if (depth === 0) { end = i + 1; break; }
      }
      const jsonStr = fullText.slice(start, end);
      try {
        JSON.parse(jsonStr);
        callbacks.onInstruction(jsonStr);
      } catch {
        // Not valid JSON
      }
    }
  }
}
