// src/components/MappingDialog.tsx
import { useState, useRef, useEffect, useMemo } from 'react';

// ─── Constants ────────────────────────────────────────────────────────────────

export const DEFAULT_CATEGORIES = [
  'Rent',
  'Opex',
  'Utility',
  'Management',
  'Insurance',
  'Tax',
  'Excluded',
];

// ─── Types ────────────────────────────────────────────────────────────────────

export interface UniqueChargePair {
  charge: string;
  chargeType: string;
}

export interface MappingDialogProps {
  uniquePairs: UniqueChargePair[];
  onClose: () => void;
  onExport: (mappings: Record<string, string>, categories: string[]) => void;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

export function pairKey(charge: string, chargeType: string): string {
  return `${charge}\x00${chargeType}`;
}

function autoSuggest(chargeType: string, categories: string[]): string {
  const ct = chargeType.toLowerCase();
  if (ct.includes('rent')) return categories.find(c => /^rent$/i.test(c))     ?? categories[0] ?? '';
  if (ct.includes('cam'))  return categories.find(c => /^opex$/i.test(c))     ?? categories[0] ?? '';
  return                          categories.find(c => /^excluded$/i.test(c)) ?? categories[0] ?? '';
}

const cc = (...parts: (string | false | undefined | null)[]) => parts.filter(Boolean).join(' ');

const btnPrimary =
  'px-3 py-1.5 text-[11px] font-mono rounded bg-primary text-primary-foreground hover:bg-primary/90 transition-colors disabled:opacity-50';
const btnSecondary =
  'px-3 py-1.5 text-[11px] font-mono border border-panel-border rounded hover:border-muted-foreground transition-colors';

// Highlight the query substring inside text
function HighlightMatch({ text, query }: { text: string; query: string }) {
  if (!query) return <>{text}</>;
  const lo  = text.toLowerCase().indexOf(query.toLowerCase());
  if (lo < 0) return <>{text}</>;
  return (
    <>
      {text.slice(0, lo)}
      <span className="font-semibold underline">{text.slice(lo, lo + query.length)}</span>
      {text.slice(lo + query.length)}
    </>
  );
}

// ─── Component ────────────────────────────────────────────────────────────────

export function MappingDialog({ uniquePairs, onClose, onExport }: MappingDialogProps) {
  // ── Step 1: categories ─────────────────────────────────────────────────────
  const [step, setStep]       = useState<'categories' | 'mapping'>('categories');
  const [categories, setCategories] = useState<string[]>([...DEFAULT_CATEGORIES]);
  const [newCat, setNewCat]   = useState('');

  // ── Step 2: mappings ───────────────────────────────────────────────────────
  const [mappings, setMappings] = useState<Record<string, string>>(() => {
    const m: Record<string, string> = {};
    for (const p of uniquePairs) {
      m[pairKey(p.charge, p.chargeType)] = autoSuggest(p.chargeType, DEFAULT_CATEGORIES);
    }
    return m;
  });

  // Excel-like selection
  const [selRows, setSelRows] = useState<Set<number>>(new Set([0]));
  const [anchor, setAnchor]   = useState(0);
  const [focused, setFocused] = useState(0);

  // Typeahead state (replaces dropdown)
  const [typeBuffer, setTypeBuffer] = useState('');
  const [typeHlIdx, setTypeHlIdx]   = useState(0);

  // Internal clipboard (Ctrl+C / Ctrl+V)
  const [clipboard, setClipboard] = useState<string | null>(null);

  const containerRef = useRef<HTMLDivElement>(null);
  const rowRefs      = useRef<(HTMLTableRowElement | null)[]>([]);

  // Filtered categories for typeahead
  const filteredCats = useMemo(
    () => typeBuffer
      ? categories.filter(c => c.toLowerCase().includes(typeBuffer.toLowerCase()))
      : [],
    [typeBuffer, categories],
  );

  // Reset highlight index when filter changes
  useEffect(() => { setTypeHlIdx(0); }, [typeBuffer]);

  // Auto-scroll focused row into view
  useEffect(() => {
    if (step === 'mapping') rowRefs.current[focused]?.scrollIntoView({ block: 'nearest' });
  }, [focused, step]);

  // ── Selection helpers ──────────────────────────────────────────────────────

  const makeRange = (a: number, b: number): Set<number> => {
    const lo = Math.min(a, b), hi = Math.max(a, b);
    return new Set(Array.from({ length: hi - lo + 1 }, (_, i) => lo + i));
  };

  const selectOne = (ri: number) => {
    setAnchor(ri); setFocused(ri); setSelRows(new Set([ri])); setTypeBuffer('');
  };

  const extendTo = (ri: number) => {
    setFocused(ri); setSelRows(makeRange(anchor, ri)); setTypeBuffer('');
  };

  const moveFocus = (delta: number, extend: boolean) => {
    const next = Math.max(0, Math.min(uniquePairs.length - 1, focused + delta));
    if (extend) extendTo(next); else selectOne(next);
  };

  // ── Mapping helpers ────────────────────────────────────────────────────────

  /** Apply `val` to all selected rows (or just focused if nothing selected) */
  const applyValue = (val: string) => {
    setMappings(prev => {
      const next = { ...prev };
      const rows = selRows.size > 0 ? selRows : new Set([focused]);
      for (const ri of rows) {
        next[pairKey(uniquePairs[ri].charge, uniquePairs[ri].chargeType)] = val;
      }
      return next;
    });
    setTypeBuffer('');
    containerRef.current?.focus();
  };

  /** Confirm the currently highlighted typeahead option and advance focus */
  const confirmTypeahead = () => {
    if (filteredCats.length > 0) {
      applyValue(filteredCats[Math.min(typeHlIdx, filteredCats.length - 1)]);
      // Move focus down (Excel Enter behaviour)
      const next = Math.min(focused + 1, uniquePairs.length - 1);
      selectOne(next);
    } else {
      setTypeBuffer('');
    }
  };

  // ── Keyboard ───────────────────────────────────────────────────────────────

  const handleKeyDown = (e: React.KeyboardEvent) => {
    const ctrl = e.ctrlKey || e.metaKey;

    // ── Ctrl shortcuts (always active) ──────────────────────────────────────
    if (ctrl) {
      if (e.key === 'c') {
        e.preventDefault();
        const val = mappings[pairKey(uniquePairs[focused].charge, uniquePairs[focused].chargeType)] ?? '';
        setClipboard(val);
        return;
      }
      if (e.key === 'v') {
        e.preventDefault();
        if (clipboard !== null) applyValue(clipboard);
        return;
      }
      // Ctrl+↓ / Ctrl+↑ → jump to last / first row
      if (!e.shiftKey && e.key === 'ArrowDown') {
        e.preventDefault();
        selectOne(uniquePairs.length - 1);
        return;
      }
      if (!e.shiftKey && e.key === 'ArrowUp') {
        e.preventDefault();
        selectOne(0);
        return;
      }
      // Ctrl+Shift+↓ → extend selection to last row
      if (e.shiftKey && e.key === 'ArrowDown') {
        e.preventDefault();
        const last = uniquePairs.length - 1;
        setFocused(last); setSelRows(makeRange(anchor, last));
        return;
      }
      // Ctrl+Shift+↑ → extend selection to first row
      if (e.shiftKey && e.key === 'ArrowUp') {
        e.preventDefault();
        setFocused(0); setSelRows(makeRange(anchor, 0));
        return;
      }
      return;
    }

    // ── Typeahead mode ───────────────────────────────────────────────────────
    if (typeBuffer !== '') {
      if (e.key === 'Escape')    { e.preventDefault(); setTypeBuffer(''); return; }
      if (e.key === 'Enter')     { e.preventDefault(); confirmTypeahead(); return; }
      if (e.key === 'Backspace') { e.preventDefault(); setTypeBuffer(p => p.slice(0, -1)); return; }
      if (e.key === 'ArrowDown') { e.preventDefault(); setTypeHlIdx(i => Math.min(i + 1, filteredCats.length - 1)); return; }
      if (e.key === 'ArrowUp')   { e.preventDefault(); setTypeHlIdx(i => Math.max(i - 1, 0)); return; }
      if (e.key.length === 1)    { e.preventDefault(); setTypeBuffer(p => p + e.key); return; }
      return;
    }

    // ── Navigation mode ──────────────────────────────────────────────────────
    if (e.key === 'ArrowDown')  { e.preventDefault(); moveFocus(1,  e.shiftKey); return; }
    if (e.key === 'ArrowUp')    { e.preventDefault(); moveFocus(-1, e.shiftKey); return; }
    if (e.key === 'Delete')     { e.preventDefault(); applyValue(''); return; }
    if (e.key === 'Escape')     { e.preventDefault(); setSelRows(new Set()); return; }

    // Start typing → open typeahead
    if (e.key.length === 1 && !e.altKey) {
      e.preventDefault();
      setTypeBuffer(e.key);
    }
  };

  // ── Step 1 ─────────────────────────────────────────────────────────────────

  if (step === 'categories') {
    const addCat = () => {
      const t = newCat.trim();
      if (t && !categories.includes(t)) { setCategories(p => [...p, t]); setNewCat(''); }
    };
    return (
      <Overlay>
        <Panel width={400}>
          <PanelHeader title="Step 1 — Mapping Categories" onClose={onClose} />
          <div className="p-4 flex flex-col gap-3">
            <p className="text-[11px] font-mono text-muted-foreground leading-relaxed">
              Define the categories charges will be mapped to. Add or remove before proceeding.
            </p>
            <div className="flex flex-col gap-1 max-h-60 overflow-y-auto pr-0.5">
              {categories.map((cat, i) => (
                <div key={i} className="flex items-center justify-between px-2.5 py-1.5 rounded border border-panel-border bg-muted/20 group">
                  <span className="text-[11px] font-mono">{cat}</span>
                  <button
                    onClick={() => setCategories(p => p.filter(c => c !== cat))}
                    className="text-muted-foreground/40 group-hover:text-red-400 transition-colors text-[10px] px-1 leading-none"
                  >
                    ✕
                  </button>
                </div>
              ))}
            </div>
            <div className="flex gap-1.5">
              <input
                value={newCat}
                onChange={e => setNewCat(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && addCat()}
                placeholder="New category name…"
                className="flex-1 px-2.5 py-1.5 text-[11px] font-mono bg-muted/20 border border-panel-border rounded focus:outline-none focus:border-muted-foreground transition-colors"
              />
              <button onClick={addCat} className={btnSecondary}>Add</button>
            </div>
          </div>
          <PanelFooter>
            <button onClick={onClose} className={btnSecondary}>Cancel</button>
            <button
              onClick={() => setStep('mapping')}
              className={btnPrimary}
              disabled={categories.length === 0}
            >
              Next → ({uniquePairs.length} charges)
            </button>
          </PanelFooter>
        </Panel>
      </Overlay>
    );
  }

  // ── Step 2 ─────────────────────────────────────────────────────────────────

  const isTyping = typeBuffer !== '';

  return (
    <Overlay>
      <Panel width={580} maxHeight="82vh">
        {/* Header */}
        <PanelHeader
          title="Step 2 — Map Charges"
          onClose={onClose}
          left={
            <button
              onClick={() => { setStep('categories'); setTypeBuffer(''); }}
              className="text-[11px] font-mono text-muted-foreground hover:text-foreground transition-colors"
            >
              ← Back
            </button>
          }
        />

        {/* Hint bar */}
        <div className="px-4 py-1.5 text-[10px] font-mono text-muted-foreground bg-muted/10 border-b border-panel-border shrink-0 flex items-center gap-2 flex-wrap">
          <span>Click / ↑↓ navigate</span>
          <span className="opacity-40">·</span>
          <span>Shift+↑↓ / Ctrl+Shift+↑↓ range</span>
          <span className="opacity-40">·</span>
          <span>Type to search</span>
          <span className="opacity-40">·</span>
          <span>Del clear</span>
          <span className="opacity-40">·</span>
          <span>Ctrl+C / Ctrl+V</span>
          {clipboard !== null && (
            <>
              <span className="opacity-40">·</span>
              <span className="text-blue-400">clipboard: {clipboard}</span>
            </>
          )}
        </div>

        {/* Table */}
        <div
          ref={containerRef}
          tabIndex={0}
          onKeyDown={handleKeyDown}
          className="flex-1 overflow-auto focus:outline-none min-h-0"
        >
          <table className="text-[11px] font-mono border-collapse w-full">
            <thead className="sticky top-0 z-10">
              <tr>
                <th className="px-3 py-1.5 text-left border border-panel-border bg-muted/60 text-muted-foreground font-medium whitespace-nowrap w-40">
                  Code
                </th>
                <th className="px-3 py-1.5 text-left border border-panel-border bg-muted/60 text-muted-foreground font-medium whitespace-nowrap w-32">
                  Lease Type
                </th>
                <th className="px-3 py-1.5 text-left border border-panel-border bg-muted/60 text-muted-foreground font-medium whitespace-nowrap">
                  Mapping
                </th>
              </tr>
            </thead>
            <tbody>
              {uniquePairs.map((pair, ri) => {
                const key   = pairKey(pair.charge, pair.chargeType);
                const val   = mappings[key] ?? '';
                const isSel = selRows.has(ri);
                const isFoc = focused === ri;
                const showTypeahead = isFoc && isTyping;

                return (
                  <tr
                    key={ri}
                    ref={el => { rowRefs.current[ri] = el; }}
                    className={cc(isSel ? 'bg-blue-500/10' : 'hover:bg-muted/20')}
                    style={{ userSelect: 'none', cursor: 'pointer' }}
                    onClick={e => {
                      e.shiftKey ? extendTo(ri) : selectOne(ri);
                      containerRef.current?.focus();
                    }}
                  >
                    {/* Col A: Code — read-only */}
                    <td className="px-3 py-1 border border-panel-border whitespace-nowrap">
                      {pair.charge}
                    </td>

                    {/* Col B: Lease Type — read-only */}
                    <td className="px-3 py-1 border border-panel-border whitespace-nowrap">
                      {pair.chargeType}
                    </td>

                    {/* Col C: Mapping — typeahead */}
                    <td
                      className={cc(
                        'px-3 py-0.5 border relative',
                        isFoc
                          ? 'border-blue-500 outline outline-1 outline-blue-500 z-[1]'
                          : isSel
                          ? 'border-blue-500/60'
                          : 'border-panel-border',
                      )}
                    >
                      {/* Cell content */}
                      {showTypeahead ? (
                        /* Typing mode: show buffer + cursor */
                        <div className="flex items-center gap-0.5 min-h-[20px]">
                          <span className="text-blue-300">{typeBuffer}</span>
                          <span className="w-px h-3.5 bg-blue-400 animate-pulse" />
                        </div>
                      ) : (
                        /* Display mode */
                        <div className="flex items-center justify-between gap-2 min-h-[20px]">
                          <span className={val ? 'text-foreground' : 'text-muted-foreground/30'}>
                            {val || '—'}
                          </span>
                          {isFoc && (
                            <span className="text-muted-foreground/30 text-[9px] shrink-0">type to edit</span>
                          )}
                        </div>
                      )}

                      {/* Typeahead dropdown */}
                      {showTypeahead && (
                        <div
                          className="absolute left-[-1px] right-[-1px] top-full z-50 bg-background border border-blue-500 border-t-0 rounded-b shadow-xl"
                          onMouseDown={e => e.preventDefault()}
                        >
                          {filteredCats.length > 0 ? (
                            filteredCats.map((cat, ci) => (
                              <div
                                key={cat}
                                className={cc(
                                  'px-3 py-1.5 cursor-pointer text-[11px] font-mono',
                                  ci === typeHlIdx
                                    ? 'bg-blue-500 text-white'
                                    : 'text-foreground hover:bg-muted/40',
                                )}
                                onMouseEnter={() => setTypeHlIdx(ci)}
                                onClick={e => { e.stopPropagation(); applyValue(cat); }}
                              >
                                <HighlightMatch text={cat} query={typeBuffer} />
                              </div>
                            ))
                          ) : (
                            <div className="px-3 py-1.5 text-[11px] font-mono text-muted-foreground/50 italic">
                              No match — press Esc to cancel
                            </div>
                          )}
                        </div>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Footer */}
        <PanelFooter>
          <span className="text-[10px] font-mono text-muted-foreground mr-auto">
            {selRows.size > 1
              ? `${selRows.size} of ${uniquePairs.length} selected`
              : `${uniquePairs.length} charge types`}
          </span>
          <button onClick={onClose} className={btnSecondary}>Cancel</button>
          <button onClick={() => onExport(mappings, categories)} className={btnPrimary}>
            ↓ Export Excel
          </button>
        </PanelFooter>
      </Panel>
    </Overlay>
  );
}

// ─── Layout primitives ────────────────────────────────────────────────────────

function Overlay({ children }: { children: React.ReactNode }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm">
      {children}
    </div>
  );
}

function Panel({
  children,
  width,
  maxHeight,
}: {
  children: React.ReactNode;
  width: number;
  maxHeight?: string;
}) {
  return (
    <div
      className="bg-background border border-panel-border rounded-lg flex flex-col shadow-2xl"
      style={{ width, maxHeight: maxHeight ?? 'unset' }}
    >
      {children}
    </div>
  );
}

function PanelHeader({
  title,
  onClose,
  left,
}: {
  title: string;
  onClose: () => void;
  left?: React.ReactNode;
}) {
  return (
    <div className="flex items-center justify-between px-4 py-3 border-b border-panel-border shrink-0">
      <div className="flex items-center gap-3">
        {left}
        <h2 className="text-[12px] font-mono font-semibold">{title}</h2>
      </div>
      <button
        onClick={onClose}
        className="text-muted-foreground hover:text-foreground transition-colors text-[11px] leading-none"
      >
        ✕
      </button>
    </div>
  );
}

function PanelFooter({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex items-center justify-end gap-2 px-4 py-3 border-t border-panel-border shrink-0">
      {children}
    </div>
  );
}
