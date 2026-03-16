

## Summary

Three changes:

1. **Anonymization is already in place** — the sample sent to AI is anonymized via `anonymizeSheet()`. The AI response contains column letters (A, B, C…) and structural rules, not actual data — so there's nothing to "restore." No change needed here.

2. **Reduce sample to 15 rows** — currently `sample-builder.ts` sends rows 1–30 + last 10. Change to send only rows 1–15 (no bottom rows).

3. **Add "View Sent Data" button** — a button in the toolbar/header area that opens a dialog showing the exact anonymized HTML that was sent to AI.

## Technical Plan

### A. Reduce sample to 15 rows (`src/lib/sample-builder.ts`)
- Change `firstEnd` from `Math.min(30, totalRows)` to `Math.min(15, totalRows)`
- Remove the bottom-rows logic entirely (no last 10 rows)
- Update context note accordingly

### B. Store sent sample for viewing (`src/hooks/useRentRollParser.ts`)
- Add state: `sentSampleHtml: string | null`
- After building the sample HTML (line 133), store it in state
- Expose `sentSampleHtml` from the hook

### C. "View Sent Data" button + dialog (`src/pages/Index.tsx`)
- Add a button in the header bar (next to "New File") that appears when `sentSampleHtml` is available
- On click, open a Dialog showing the raw anonymized HTML rendered in a scrollable container
- Use existing shadcn Dialog component

