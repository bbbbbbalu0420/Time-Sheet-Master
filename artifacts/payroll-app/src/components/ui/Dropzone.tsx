import React, { useCallback } from "react";
import { useDropzone, DropzoneOptions } from "react-dropzone";
import { UploadCloud, FileSpreadsheet, X } from "lucide-react";
import { cn } from "@/lib/utils";

interface DropzoneProps extends Omit<DropzoneOptions, 'onDrop'> {
  onDropFiles: (files: File[]) => void;
  title?: string;
  description?: string;
  className?: string;
  maxFiles?: number;
  accept?: Record<string, string[]>;
  value?: File[];
  onRemove?: (file: File) => void;
}

export function Dropzone({
  onDropFiles,
  title = "Перетащите файлы сюда",
  description = "или кликните для выбора (.xlsx, .xls)",
  className,
  value = [],
  onRemove,
  ...props
}: DropzoneProps) {
  const onDrop = useCallback((acceptedFiles: File[]) => {
    onDropFiles(acceptedFiles);
  }, [onDropFiles]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    ...props,
  });

  return (
    <div className="w-full">
      <div
        {...getRootProps()}
        className={cn(
          "relative overflow-hidden rounded-2xl border-2 border-dashed transition-all duration-300 ease-out cursor-pointer group",
          isDragActive 
            ? "border-primary bg-primary/5 scale-[1.02]" 
            : "border-white/10 bg-white/5 hover:bg-white/10 hover:border-white/20",
          className
        )}
      >
        <input {...getInputProps()} />
        <div className="flex flex-col items-center justify-center p-10 text-center">
          <div className={cn(
            "p-4 rounded-full mb-4 transition-colors duration-300",
            isDragActive ? "bg-primary/20 text-primary" : "bg-white/5 text-muted-foreground group-hover:bg-white/10 group-hover:text-foreground"
          )}>
            <UploadCloud className="w-8 h-8" />
          </div>
          <h3 className="text-lg font-semibold text-foreground mb-1">
            {title}
          </h3>
          <p className="text-sm text-muted-foreground max-w-xs">
            {description}
          </p>
        </div>
      </div>

      {value.length > 0 && (
        <div className="mt-4 space-y-2">
          {value.map((file, i) => (
            <div key={i} className="flex items-center justify-between p-3 rounded-xl bg-white/5 border border-white/10">
              <div className="flex items-center gap-3">
                <FileSpreadsheet className="w-5 h-5 text-primary" />
                <span className="text-sm font-medium">{file.name}</span>
                <span className="text-xs text-muted-foreground">
                  {(file.size / 1024 / 1024).toFixed(2)} MB
                </span>
              </div>
              {onRemove && (
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    onRemove(file);
                  }}
                  className="p-1 rounded-md hover:bg-white/10 text-muted-foreground hover:text-destructive transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
