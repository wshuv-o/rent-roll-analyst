import { useCallback } from 'react';
import { Upload } from 'lucide-react';

interface FileUploadProps {
  onFileSelect: (file: File) => void;
  isProcessing: boolean;
}

export function FileUpload({ onFileSelect, isProcessing }: FileUploadProps) {
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && /\.(xlsx|xls|csv)$/i.test(file.name)) {
      onFileSelect(file);
    }
  }, [onFileSelect]);

  const handleChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) onFileSelect(file);
  }, [onFileSelect]);

  return (
    <div
      onDrop={handleDrop}
      onDragOver={e => e.preventDefault()}
      className={`border-2 border-dashed border-panel-border rounded-sm p-8 text-center transition-colors ${
        isProcessing
          ? 'opacity-50 pointer-events-none'
          : 'hover:border-muted-foreground cursor-pointer'
      }`}
    >
      <label className="cursor-pointer flex flex-col items-center gap-3">
        <Upload className="w-8 h-8 text-muted-foreground" />
        <div className="font-mono text-sm text-foreground">
          Drop Excel file here
        </div>
        <div className="font-mono text-xs text-muted-foreground">
          .xlsx or .xls only
        </div>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleChange}
          className="hidden"
          disabled={isProcessing}
        />
      </label>
    </div>
  );
}
