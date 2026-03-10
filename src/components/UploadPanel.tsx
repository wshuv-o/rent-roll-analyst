import { FileUpload } from './FileUpload';
import { TenantTable } from './TenantTable';
import type { TenantObject } from '@/lib/types';

interface UploadPanelProps {
  onFileSelect: (file: File) => void;
  isProcessing: boolean;
  tenants: TenantObject[];
  fileName: string;
}

export function UploadPanel({ onFileSelect, isProcessing, tenants, fileName }: UploadPanelProps) {
  return (
    <div className="flex flex-col h-full">
      <div className="px-4 py-3 border-b border-panel-border">
        <h2 className="font-heading text-sm uppercase tracking-wider text-muted-foreground">
          Upload & Output
        </h2>
      </div>
      <div className="flex-1 overflow-y-auto p-4">
        <FileUpload onFileSelect={onFileSelect} isProcessing={isProcessing} />
        {tenants.length > 0 && (
          <div className="mt-4">
            <TenantTable tenants={tenants} fileName={fileName} />
          </div>
        )}
      </div>
    </div>
  );
}
