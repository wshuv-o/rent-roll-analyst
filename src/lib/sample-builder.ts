/**
 * Build a representative sample from the anonymized sheet data.
 * Always includes rows 1-30 and the last 10 rows.
 * Returns an HTML table string preserving visual structure.
 */
export function buildSample(
  data: (string | number | null)[][],
  totalRows: number
): { html: string; contextNote: string; sampleRanges: string } {
  const firstEnd = Math.min(15, totalRows);

  const topRows = data.slice(0, firstEnd);

  const maxCols = data.reduce((max, row) => Math.max(max, row.length), 0);

  // Column letters
  const colLetters = Array.from({ length: maxCols }, (_, i) => {
    let letter = '';
    let n = i;
    while (n >= 0) {
      letter = String.fromCharCode(65 + (n % 26)) + letter;
      n = Math.floor(n / 26) - 1;
    }
    return letter;
  });

  function rowsToHtml(rows: (string | number | null)[][], startIdx: number): string {
    return rows.map((row, i) => {
      const rowNum = startIdx + i + 1;
      const cells = Array.from({ length: maxCols }, (_, c) => {
        const val = c < row.length ? (row[c] ?? '') : '';
        return `<td>${String(val)}</td>`;
      }).join('');
      return `<tr><td class="row-num">${rowNum}</td>${cells}</tr>`;
    }).join('\n');
  }

  const headerRow = `<tr><th></th>${colLetters.map(l => `<th>${l}</th>`).join('')}</tr>`;

  const tableBody = rowsToHtml(topRows, 0);

  const html = `<table border="1" cellpadding="4" cellspacing="0">\n<thead>${headerRow}</thead>\n<tbody>\n${tableBody}\n</tbody>\n</table>`;

  const rangeStr = `rows 1–${firstEnd}`;

  const contextNote = `This sheet has ${totalRows} total rows. You are seeing ${rangeStr}. The full sheet will be processed after you confirm the layout.`;

  return { html, contextNote, sampleRanges: rangeStr };
}
