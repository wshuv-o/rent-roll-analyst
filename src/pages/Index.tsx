import { UploadPanel } from '@/components/UploadPanel';
import { ActivityLog } from '@/components/ActivityLog';
import { useRentRollParser } from '@/hooks/useRentRollParser';

const Index = () => {
  const { logs, tenants, isProcessing, fileName, processFile } = useRentRollParser();

  return (
    <div className="h-screen flex flex-col bg-background">
      {/* Header */}
      <div className="shrink-0 border-b border-panel-border px-4 py-1.5 flex items-center justify-between">
        <h1 className="font-heading text-sm tracking-wide text-foreground">
          RENT ROLL PARSER
        </h1>
        <span className="font-mono text-[10px] text-muted-foreground">v1.0</span>
      </div>

      {/* Main layout: Excel dominant, log sidebar */}
      <div className="flex-1 flex min-h-0">
        {/* Main panel — Excel & output */}
        <div className="flex-1 min-w-0 border-r border-panel-border">
          <UploadPanel
            onFileSelect={processFile}
            isProcessing={isProcessing}
            tenants={tenants}
            fileName={fileName}
          />
        </div>

        {/* Right sidebar — compact activity log */}
        <div className="w-[280px] shrink-0">
          <ActivityLog entries={logs} />
        </div>
      </div>
    </div>
  );
};

export default Index;
