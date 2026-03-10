import { FileUpload } from '@/components/FileUpload';
import { SpreadsheetViewer } from '@/components/SpreadsheetViewer';
import { ColumnMappingToolbar } from '@/components/ColumnMappingToolbar';
import { TenantTable } from '@/components/TenantTable';
import { ActivityLog } from '@/components/ActivityLog';
import { useRentRollParser } from '@/hooks/useRentRollParser';

const Index = () => {
  const {
    logs, tenants, isProcessing, fileName, step,
    sheetData, headerRows, instruction, groupSpans,
    columnAliases,
    loadFile, handleColumnAssign, handleCustomFieldAssign, handleGroupResize,
    handleColumnRename,
    confirmAndParse, resetToUpload, reAnalyze,
  } = useRentRollParser();

  return (
    <div className="h-screen flex flex-col bg-background">
      {/* Header */}
      <div className="shrink-0 border-b border-panel-border px-4 py-1.5 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <h1 className="font-heading text-sm tracking-wide text-foreground">
            RENT ROLL PARSER
          </h1>
          {step !== 'upload' && fileName && (
            <span className="text-[11px] font-mono text-muted-foreground">
              — {fileName}
            </span>
          )}
        </div>
        <div className="flex items-center gap-3">
          {step !== 'upload' && (
            <button
              onClick={resetToUpload}
              className="text-[11px] font-mono text-muted-foreground hover:text-foreground transition-colors"
            >
              New File
            </button>
          )}
          <span className="font-mono text-[10px] text-muted-foreground">v2.0</span>
        </div>
      </div>

      {/* Main layout */}
      <div className="flex-1 flex min-h-0">
        {/* Main panel */}
        <div className="flex-1 min-w-0 flex flex-col border-r border-panel-border">
          {step === 'upload' && (
            <div className="flex-1 flex items-center justify-center p-8">
              <div className="w-full max-w-md">
                <FileUpload onFileSelect={loadFile} isProcessing={isProcessing} />
              </div>
            </div>
          )}

          {(step === 'analyzing' || step === 'confirm') && sheetData.length > 0 && (
            <>
              <div className="flex-1 min-h-0">
                <SpreadsheetViewer
                  data={sheetData}
                  instruction={instruction}
                  headerRows={headerRows}
                  groupSpans={groupSpans}
                  columnAliases={columnAliases}
                  onColumnAssign={step === 'confirm' ? handleColumnAssign : undefined}
                  onCustomFieldAssign={step === 'confirm' ? handleCustomFieldAssign : undefined}
                  onGroupResize={step === 'confirm' ? handleGroupResize : undefined}
                  onColumnRename={step === 'confirm' ? handleColumnRename : undefined}
                />
              </div>
              <ColumnMappingToolbar
                instruction={instruction}
                isAnalyzing={step === 'analyzing'}
                onConfirm={confirmAndParse}
                onReset={reAnalyze}
              />
            </>
          )}

          {step === 'parsing' && (
            <div className="flex-1 flex items-center justify-center">
              <span className="text-sm font-mono text-muted-foreground animate-pulse">
                Parsing tenants...
              </span>
            </div>
          )}

          {step === 'done' && tenants.length > 0 && (
            <div className="flex-1 overflow-y-auto p-4">
              <TenantTable tenants={tenants} fileName={fileName} />
            </div>
          )}

          {step === 'done' && tenants.length === 0 && (
            <div className="flex-1 flex items-center justify-center">
              <span className="text-sm font-mono text-log-flag">
                0 tenants found. Try adjusting column assignments.
              </span>
            </div>
          )}
        </div>

        {/* Activity log sidebar */}
        <div className="w-[280px] shrink-0">
          <ActivityLog entries={logs} />
        </div>
      </div>
    </div>
  );
};

export default Index;