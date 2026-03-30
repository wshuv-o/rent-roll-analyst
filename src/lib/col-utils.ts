/** Convert 0-based column index to Excel letter (0→A, 25→Z, 26→AA) */
export function indexToColLetter(idx: number): string {
  let letter = '';
  let n = idx;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

/** Convert Excel column letter to 0-based index (A→0, Z→25, AA→26) */
export function colLetterToIndex(letter: string): number {
  if (!letter) return -1;
  const upper = letter.toUpperCase().trim().replace(/[^A-Z]/g, '');
  if (!upper) return -1;
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

/** Get cell value as string from a row */
export function getCellValue(row: (string | number | Date | null)[], colIdx: number): string {
  if (colIdx < 0 || colIdx >= row.length) return '';
  const val = row[colIdx];
  return val !== null && val !== undefined ? String(val).trim() : '';
}

/** Get raw cell value from a row (preserving type) */
export function getRawCellValue(row: (string | number | Date | null)[], colIdx: number): string | number | Date | null {
  if (colIdx < 0 || colIdx >= row.length) return null;
  return row[colIdx] ?? null;
}
