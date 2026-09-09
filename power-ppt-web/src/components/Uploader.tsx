import * as React from "react";
import { UploadCloud, FileText, Loader2 } from "lucide-react";
import { cn } from "@/lib/utils";

interface UploaderProps {
  onFile: (file: File) => void;
  busy: boolean;
}

export function Uploader({ onFile, busy }: UploaderProps) {
  const [dragging, setDragging] = React.useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);

  function pick(file: File | undefined) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".pptx")) return; // guard mirrors the API
    onFile(file);
  }

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        if (!busy) pick(e.dataTransfer.files[0]);
      }}
      onClick={() => !busy && inputRef.current?.click()}
      className={cn(
        "flex cursor-pointer flex-col items-center justify-center gap-3 rounded-card border-2 border-dashed bg-white px-6 py-16 text-center transition-colors",
        dragging ? "border-brand-green bg-brand-green/5" : "border-gray-300",
        busy && "pointer-events-none opacity-70",
      )}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".pptx"
        className="hidden"
        onChange={(e) => pick(e.target.files?.[0])}
      />
      {busy ? (
        <Loader2 className="h-10 w-10 animate-spin text-brand-midblue" />
      ) : (
        <UploadCloud className="h-10 w-10 text-brand-midblue" />
      )}
      <div>
        <p className="font-medium text-brand-blue">
          {busy ? "Analysing your deck…" : "Drop a .pptx here, or click to browse"}
        </p>
        <p className="mt-1 flex items-center justify-center gap-1 text-sm text-brand-slate">
          <FileText className="h-4 w-4" /> PowerPoint decks only
        </p>
      </div>
    </div>
  );
}
