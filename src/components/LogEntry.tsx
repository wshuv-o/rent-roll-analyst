import { LogType } from '@/lib/types';

interface LogEntryProps {
  type: LogType;
  message: string;
  timestamp: Date;
  isStreaming?: boolean;
}

const badgeConfig: Record<LogType, { emoji: string; label: string; colorClass: string }> = {
  system: { emoji: '🔵', label: 'SYSTEM', colorClass: 'text-log-system' },
  thinking: { emoji: '🟡', label: 'THINKING', colorClass: 'text-log-thinking' },
  grouping: { emoji: '🟠', label: 'GROUPING', colorClass: 'text-log-grouping' },
  output: { emoji: '🟢', label: 'OUTPUT', colorClass: 'text-log-output' },
  flag: { emoji: '🔴', label: 'FLAG', colorClass: 'text-log-flag' },
};

export function LogEntryComponent({ type, message, timestamp, isStreaming }: LogEntryProps) {
  const badge = badgeConfig[type];
  const time = timestamp.toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });

  return (
    <div className="log-entry-appear flex gap-3 py-1.5 px-3 text-[13px] leading-relaxed font-mono">
      <span className="text-muted-foreground shrink-0 tabular-nums">{time}</span>
      <span className={`shrink-0 font-semibold ${badge.colorClass}`}>
        {badge.emoji} {badge.label}
      </span>
      <span className="text-foreground/90 break-words min-w-0">
        {message}
        {isStreaming && <span className="typing-cursor" />}
      </span>
    </div>
  );
}
