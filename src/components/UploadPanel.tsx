import { useCallback, useRef, useState } from "react";

interface UploadPanelProps {
  fileName: string | null;
  isProcessing: boolean;
  onFileSelected: (file: File) => void;
  onDownloadSample: () => void;
}

const ACCEPTED_EXT = [".xlsx", ".xlsm"];

export default function UploadPanel({ fileName, isProcessing, onFileSelected, onDownloadSample }: UploadPanelProps) {
  const [isDragOver, setIsDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const isValidFile = (file: File) => ACCEPTED_EXT.some((ext) => file.name.toLowerCase().endsWith(ext));

  const handleFiles = useCallback(
    (files: FileList | null) => {
      if (!files || files.length === 0) return;
      const file = files[0];
      if (!isValidFile(file)) {
        window.alert("Please upload a .xlsx or .xlsm file.\n請上傳 .xlsx 或 .xlsm 檔案。");
        return;
      }
      onFileSelected(file);
    },
    [onFileSelected],
  );

  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-900">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">Upload Excel File 上傳 Excel 檔案</h2>
          <p className="mt-0.5 text-xs text-slate-500 dark:text-slate-400">
            All processing happens locally in your browser. Files are never uploaded to any server.
            <br />
            所有處理均在瀏覽器本機完成，檔案不會上傳至任何伺服器。
          </p>
        </div>
        <a
        href="/sample.xlsx"
        download="sample.xlsx"
        className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-slate-50 px-3 py-2 text-xs font-medium text-slate-700 transition hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
        >
          <svg className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v12m0 0 4-4m-4 4-4-4M4 19h16" />
          </svg>
          Download Sample Template 下載範本
        </a>
      </div>

      <div
        onDragOver={(e) => {
          e.preventDefault();
          setIsDragOver(true);
        }}
        onDragLeave={() => setIsDragOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setIsDragOver(false);
          handleFiles(e.dataTransfer.files);
        }}
        onClick={() => inputRef.current?.click()}
        className={`flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed px-6 py-10 text-center transition ${
          isDragOver
            ? "border-indigo-500 bg-indigo-50 dark:border-indigo-400 dark:bg-indigo-500/10"
            : "border-slate-300 bg-slate-50/50 hover:border-indigo-400 hover:bg-indigo-50/40 dark:border-slate-700 dark:bg-slate-950/40 dark:hover:border-indigo-500 dark:hover:bg-indigo-500/5"
        }`}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xlsm"
          className="hidden"
          onChange={(e) => handleFiles(e.target.files)}
        />
        <div className="flex h-12 w-12 items-center justify-center rounded-full bg-indigo-100 text-indigo-600 dark:bg-indigo-500/15 dark:text-indigo-400">
          <svg className="h-6 w-6" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2M7 9l5-5 5 5M12 4v12" />
          </svg>
        </div>
        {isProcessing ? (
          <p className="text-sm font-medium text-indigo-600 dark:text-indigo-400">Processing… 處理中…</p>
        ) : fileName ? (
          <>
            <p className="text-sm font-medium text-slate-800 dark:text-slate-100">{fileName}</p>
            <p className="text-xs text-slate-500 dark:text-slate-400">Click or drop another file to replace 點擊或拖放以更換檔案</p>
          </>
        ) : (
          <>
            <p className="text-sm font-medium text-slate-700 dark:text-slate-200">
              Drag &amp; drop your Excel file here, or click to browse
            </p>
            <p className="text-xs text-slate-500 dark:text-slate-400">拖放 Excel 檔案至此，或點擊選擇檔案 (.xlsx / .xlsm)</p>
          </>
        )}
      </div>
    </section>
  );
}
