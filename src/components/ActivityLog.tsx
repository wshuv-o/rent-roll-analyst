import { useRef, useEffect } from 'react';
import type { LogEntry } from '@/lib/types';
import { LogEntryComponent } from './LogEntry';

interface ActivityLogProps {
  entries: LogEntry[];
}

export function ActivityLog({ entries }: ActivityLogProps) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const isAutoScrolling = useRef(true);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el || !isAutoScrolling.current) return;
    el.scrollTop = el.scrollHeight;
  }, [entries]);

  const handleScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 50;
    isAutoScrolling.current = atBottom;
  };

  return (
    <div className="flex flex-col h-full">
      <div className="px-2 py-1.5 border-b border-panel-border">
        <h2 className="font-heading text-[11px] uppercase tracking-wider text-muted-foreground">
          Activity
        </h2>
      </div>
      <div
        ref={scrollRef}
        onScroll={handleScroll}
        className="flex-1 overflow-y-auto"
      >
        {entries.length === 0 ? (
          <div className="flex items-center justify-center h-full text-muted-foreground text-[11px] font-mono">
            Waiting...
          </div>
        ) : (
          <div className="py-1">
            {entries.map(entry => (
              <LogEntryComponent
                key={entry.id}
                type={entry.type}
                message={entry.message}
                timestamp={entry.timestamp}
                isStreaming={entry.isStreaming}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
