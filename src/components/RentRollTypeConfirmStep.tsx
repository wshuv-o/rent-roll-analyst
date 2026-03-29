//src/components/RentRollTypeConfirmStep.tsx
import { useState } from 'react';
import { RENT_ROLL_TYPES } from '@/lib/rent-roll-types';

// ─── SVG Illustrations ────────────────────────────────────────────────────────

const IllustrationRegular = () => (
  <svg viewBox="0 0 160 100" className="w-full h-full" aria-hidden>
    <rect x="4" y="4" width="152" height="10" rx="1" fill="currentColor" opacity="0.5" />
    {[0,1,2,3,4,5].map((i) => (
      <g key={i} transform={`translate(0, ${18 + i * 13})`}>
        <rect x="4" y="0" width="22" height="9" rx="1"
          fill="currentColor" opacity={i % 2 === 0 ? 0.7 : 0.12} />
        <rect x="30" y="0" width="40" height="9" rx="1"
          fill="currentColor" opacity={i % 2 === 0 ? 0.6 : 0.20} />
        {[0,1,2,3].map((j) => (
          <rect key={j} x={74 + j * 21} y="0" width="17" height="9" rx="1"
            fill="currentColor" opacity={i % 2 === 0 ? 0.45 : 0.25} />
        ))}
      </g>
    ))}
  </svg>
);

const IllustrationTenancySchedule = () => (
  <svg viewBox="0 0 160 100" className="w-full h-full" aria-hidden>
    <rect x="4" y="4" width="152" height="8" rx="1" fill="currentColor" opacity="0.5" />
    <rect x="4" y="16" width="152" height="8" rx="1" fill="currentColor" opacity="0.65" />
    <rect x="4" y="27" width="36" height="7" rx="1" fill="#f59e0b" opacity="0.7" />
    <rect x="4" y="37" width="152" height="7" rx="1" fill="currentColor" opacity="0.30" />
    {[0,1].map(i => (
      <g key={i} transform={`translate(0, ${47 + i * 9})`}>
        <rect x="28" y="0" width="128" height="7" rx="1" fill="currentColor" opacity="0.35" />
      </g>
    ))}
    <rect x="4" y="66" width="52" height="7" rx="1" fill="#f59e0b" opacity="0.7" />
    <rect x="4" y="76" width="152" height="7" rx="1" fill="currentColor" opacity="0.30" />
    <rect x="28" y="86" width="128" height="7" rx="1" fill="currentColor" opacity="0.35" />
  </svg>
);

const ILLUSTRATIONS: Record<string, React.ReactNode> = {
  'regular': <IllustrationRegular />,
  'tenancy-schedule': <IllustrationTenancySchedule />,
};

// ─── Component ────────────────────────────────────────────────────────────────

interface Props {
  fileName: string;
  onProceed: (typeId: string) => void;
}

export function RentRollTypeConfirmStep({ fileName, onProceed }: Props) {
  const [selected, setSelected] = useState('');

  return (
    <div className="flex-1 flex flex-col items-center justify-center p-8 gap-6">
      <div className="text-center space-y-1">
        <p className="text-[11px] font-mono text-muted-foreground">{fileName}</p>
        <h2 className="text-sm font-heading tracking-wide text-foreground">
          Select the rent roll type
        </h2>
      </div>

      <div className="grid grid-cols-2 gap-4 w-full max-w-lg">
        {RENT_ROLL_TYPES.map((type) => {
          const isSelected = selected === type.id;
          const disabled = !type.implemented;
          return (
            <button
              key={type.id}
              onClick={() => !disabled && setSelected(type.id)}
              disabled={disabled}
              className={[
                'relative flex flex-col rounded-lg border p-3 text-left transition-all gap-3',
                disabled
                  ? 'border-panel-border bg-background opacity-50 cursor-not-allowed'
                  : isSelected
                    ? 'border-primary bg-primary/10 shadow-md'
                    : 'border-panel-border hover:border-muted-foreground bg-background',
              ].join(' ')}
            >
              {disabled && (
                <span className="absolute -top-2 left-3 text-[9px] font-mono px-1.5 py-0.5 rounded bg-muted text-muted-foreground border border-panel-border">
                  coming soon
                </span>
              )}

              <div className={[
                'w-full aspect-[16/10] rounded overflow-hidden flex items-center justify-center',
                isSelected ? 'text-primary' : 'text-muted-foreground',
              ].join(' ')}>
                {ILLUSTRATIONS[type.id]}
              </div>

              <div>
                <p className={[
                  'text-[11px] font-mono font-medium',
                  isSelected ? 'text-primary' : 'text-foreground',
                ].join(' ')}>
                  {type.label}
                </p>
                <p className="text-[10px] text-muted-foreground leading-snug mt-0.5">
                  {type.description}
                </p>
              </div>

              {isSelected && (
                <div className="absolute top-2 right-2 w-4 h-4 rounded-full bg-primary flex items-center justify-center">
                  <svg viewBox="0 0 10 10" className="w-2.5 h-2.5 text-primary-foreground" fill="none">
                    <polyline points="1.5,5 4,7.5 8.5,2.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                </div>
              )}
            </button>
          );
        })}
      </div>

      <button
        onClick={() => selected && onProceed(selected)}
        disabled={!selected}
        className={[
          'px-6 py-2 text-[11px] font-mono rounded border transition-all',
          selected
            ? 'bg-primary text-primary-foreground border-primary hover:opacity-90'
            : 'text-muted-foreground border-panel-border cursor-not-allowed opacity-50',
        ].join(' ')}
      >
        Proceed
      </button>
    </div>
  );
}
