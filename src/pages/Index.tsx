import { UploadPanel } from '@/components/UploadPanel';
import { ActivityLog } from '@/components/ActivityLog';
import { useRentRollParser } from '@/hooks/useRentRollParser';

const Index = () => {
  const { logs, tenants, isProcessing, fileName, processFile } = useRentRollParser();

  return (
    <div className="h-screen flex flex-col bg-background">
      {/* Header */}
      <div className="shrink-0 border-b border-panel-border px-4 py-2 flex items-center justify-between">
        <h1 className="font-heading text-base tracking-wide text-foreground">
          RENT ROLL PARSER
        </h1>
        <span className="font-mono text-[11px] text-muted-foreground">
          v1.0 — glass-box engine
        </span>
      </div>

      {/* Two-panel layout */}
      <div className="flex-1 flex min-h-0">
        {/* Left panel — 40% */}
        <div className="w-[40%] border-r border-panel-border">
          <UploadPanel
            onFileSelect={processFile}
            isProcessing={isProcessing}
            tenants={tenants}
            fileName={fileName}
          />
        </div>

        {/* Right panel — 60% */}
        <div className="w-[60%]">
          <ActivityLog entries={logs} />
        </div>
      </div>

      {/* Min-width guard */}
      <div className="fixed inset-0 bg-background flex items-center justify-center min-[1200px]:hidden z-50">
        <div className="text-center font-mono text-sm text-muted-foreground p-8">
          <p>Rent Roll Parser requires a</p>
          <p>minimum width of 1200px.</p>
        </div>
      </div>
    </div>
  );
};

export default Index;
