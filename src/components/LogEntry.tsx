import { LogType } from '@/lib/types';

interface LogEntryProps {
  type: LogType;
  message: string;
  timestamp: Date;
  isStreaming?: boolean;
}

const badgeConfig: Record<LogType, { label: string; colorClass: string }> = {
  system: { label: 'SYS', colorClass: 'text-log-system' },
  thinking: { label: 'AI', colorClass: 'text-log-thinking' },
  grouping: { label: 'MAP', colorClass: 'text-log-grouping' },
  output: { label: 'OUT', colorClass: 'text-log-output' },
  flag: { label: 'FLG', colorClass: 'text-log-flag' },
};

export function LogEntryComponent({ type, message, timestamp, isStreaming }: LogEntryProps) {
  const badge = badgeConfig[type];
  const time = timestamp.toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });

  return (
    <div className="log-entry-appear flex gap-1.5 py-0.5 px-2 text-[11px] leading-snug font-mono">
      <span className="text-muted-foreground shrink-0 tabular-nums">{time}</span>
      <span className={`shrink-0 font-semibold ${badge.colorClass}`}>{badge.label}</span>
      <span className="text-foreground/80 break-words min-w-0 line-clamp-3">
        {message}
        {isStreaming && <span className="typing-cursor" />}
      </span>
    </div>
  );
}
