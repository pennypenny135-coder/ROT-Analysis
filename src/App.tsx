import { useMemo, useRef, useState } from "react";
import ExcelJS from "exceljs";
import ThemeToggle, { useTheme } from "./components/ThemeToggle";
import UploadPanel from "./components/UploadPanel";
import StatsPanel from "./components/StatsPanel";
import ErrorPanel from "./components/ErrorPanel";
import PivotBuilder from "./components/PivotBuilder";
import { readWorkbookData } from "./lib/excel/readWorkbook";
import { runLookup } from "./lib/excel/lookup";
import { applyAutoFillAndErrors, finalizeAndDownload, finalizeAndDownloadWithPivot } from "./lib/excel/exportWorkbook";
import { buildSummarySheet } from "./lib/excel/summary";
import { downloadSampleTemplate } from "./lib/excel/sampleTemplate";
import { buildJoinedRecords, type PivotConfig, type PivotResult } from "./lib/excel/pivot";
import type { ParsedIssue, PLRow, ROTRow, SARow, WorkbookStats } from "./lib/excel/types";

interface LoadedData {
  plRows: PLRow[];
  saRows: SARow[];
  rotRows: ROTRow[];
  issues: ParsedIssue[];
}

export default function App() {
  const { theme, setTheme } = useTheme();
  const workbookRef = useRef<ExcelJS.Workbook | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [data, setData] = useState<LoadedData | null>(null);
  const [summaryGenerated, setSummaryGenerated] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [fatalError, setFatalError] = useState<string | null>(null);

  const stats: WorkbookStats | null = useMemo(() => {
    if (!data) return null;
    return {
      plRowCount: data.plRows.length,
      saRowCount: data.saRows.length,
      rotRowCount: data.rotRows.length,
      plCount: data.saRows.filter((s) => s.isTender === false).length,
      tenderCount: data.saRows.filter((s) => s.isTender === true).length,
      soldCount: data.rotRows.filter((r) => r.dateOfPASP !== null).length,
      errorCount: data.issues.filter((i) => i.severity === "error").length,
      warningCount: data.issues.filter((i) => i.severity === "warning").length,
    };
  }, [data]);

  const joinedRecords = useMemo(() => {
    if (!data) return [];
    return buildJoinedRecords(data.plRows, data.saRows, data.rotRows);
  }, [data]);

  async function handleFileSelected(file: File) {
    setIsProcessing(true);
    setFatalError(null);
    setNotice(null);
    setSummaryGenerated(false);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const parsed = await readWorkbookData(workbook);
      if (parsed.missingSheets.length > 0) {
        setFatalError(
          `Missing required sheet(s): ${parsed.missingSheets.join(", ")}. The workbook must contain PL, SA and ROT sheets.\n缺少必要的工作表：${parsed.missingSheets.join(", ")}。活頁簿必須包含 PL、SA 及 ROT 三張工作表。`,
        );
        setIsProcessing(false);
        return;
      }

      const looked = runLookup(parsed.plRows, parsed.saRows, parsed.rotRows);
      const allIssues = [...parsed.issues, ...looked.issues];

      workbookRef.current = workbook;
      setFileName(file.name);
      setData({ plRows: looked.plRows, saRows: looked.saRows, rotRows: looked.rotRows, issues: allIssues });
    } catch (err) {
      console.error(err);
      setFatalError(`Failed to read the Excel file. Please check the file format.\n無法讀取 Excel 檔案，請檢查檔案格式。(${(err as Error).message})`);
    } finally {
      setIsProcessing(false);
    }
  }

  async function handleDownloadSample() {
    try {
      await downloadSampleTemplate();
    } catch (err) {
      console.error(err);
      window.alert("Failed to generate sample template. 無法生成範本檔案。");
    }
  }

  function handleGenerateSummary() {
    if (!workbookRef.current || !data) return;
    try {
      applyAutoFillAndErrors(workbookRef.current, data.plRows, data.saRows, data.rotRows, data.issues);
      buildSummarySheet(workbookRef.current, data.plRows, data.saRows, data.rotRows);
      setSummaryGenerated(true);
      setNotice("Summary sheet generated successfully. 已成功生成 Summary 工作表。");
    } catch (err) {
      console.error(err);
      window.alert(`Failed to generate Summary sheet: ${(err as Error).message}`);
    }
  }

  async function handleDownloadExcel() {
    if (!workbookRef.current || !data || !fileName) return;
    try {
      await finalizeAndDownload(workbookRef.current, data.plRows, data.saRows, data.rotRows, data.issues, fileName);
    } catch (err) {
      console.error(err);
      window.alert(`Failed to export Excel file: ${(err as Error).message}`);
    }
  }

  async function handlePivotExport(config: PivotConfig, result: PivotResult) {
    if (!workbookRef.current || !data || !fileName) return;
    try {
      await finalizeAndDownloadWithPivot(workbookRef.current, data.plRows, data.saRows, data.rotRows, data.issues, fileName, config, result);
    } catch (err) {
      console.error(err);
      window.alert(`Failed to export pivot: ${(err as Error).message}`);
    }
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 transition-colors dark:bg-slate-950 dark:text-slate-100">
      <header className="border-b border-slate-200 bg-white/80 backdrop-blur dark:border-slate-800 dark:bg-slate-950/80">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3 px-4 py-4 sm:px-6">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-500 to-sky-500 text-white shadow-sm">
              <svg className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M3 3v18h18M7 15l4-5 3 3 5-7" />
              </svg>
            </div>
            <div>
              <h1 className="text-lg font-bold leading-tight text-slate-900 dark:text-slate-50">
                Record of Transaction Summary
              </h1>
              <p className="text-xs text-slate-500 dark:text-slate-400">交易紀錄摘要工具 · 100% 本機處理，資料不上傳</p>
            </div>
          </div>
          <ThemeToggle theme={theme} setTheme={setTheme} />
        </div>
      </header>

      <main className="mx-auto max-w-6xl space-y-6 px-4 py-6 sm:px-6">
        <UploadPanel fileName={fileName} isProcessing={isProcessing} onFileSelected={handleFileSelected} onDownloadSample={handleDownloadSample} />

        {fatalError && (
          <section className="rounded-2xl border border-red-300 bg-red-50 p-4 text-sm font-medium text-red-700 shadow-sm dark:border-red-900/50 dark:bg-red-500/10 dark:text-red-300">
            <pre className="whitespace-pre-wrap font-sans">{fatalError}</pre>
          </section>
        )}

        {notice && (
          <section className="rounded-2xl border border-emerald-300 bg-emerald-50 p-4 text-sm font-medium text-emerald-700 shadow-sm dark:border-emerald-900/50 dark:bg-emerald-500/10 dark:text-emerald-300">
            {notice}
          </section>
        )}

        {stats && data && (
          <>
            <StatsPanel stats={stats} />

            <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-900">
              <h2 className="mb-1 text-base font-semibold text-slate-900 dark:text-slate-100">Generate &amp; Download 生成及下載</h2>
              <p className="mb-4 text-xs text-slate-500 dark:text-slate-400">
                Generate the 8 fixed Summary tables (Excel formulas), then download the completed workbook. The output keeps the
                original PL / SA / ROT sheets and rebuilds the Summary sheet.
                <br />
                生成 8 張固定 Summary 分析表（使用 Excel 公式），然後下載完整活頁簿。輸出檔案會保留原本 PL / SA / ROT 工作表，並重建 Summary 工作表。
              </p>
              <div className="flex flex-wrap gap-3">
                <button
                  type="button"
                  onClick={handleGenerateSummary}
                  className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-500"
                >
                  <svg className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 17v-6h6v6M9 7h6M4 21h16a1 1 0 0 0 1-1V4a1 1 0 0 0-1-1H4a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1Z" />
                  </svg>
                  Generate Summary 生成摘要
                </button>
                <button
                  type="button"
                  onClick={handleDownloadExcel}
                  className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
                >
                  <svg className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v12m0 0 4-4m-4 4-4-4M4 19h16" />
                  </svg>
                  Download Excel 下載 Excel
                </button>
                {!summaryGenerated && (
                  <p className="flex items-center text-xs text-amber-600 dark:text-amber-400">
                    Tip: Download will auto-generate Summary if you have not clicked "Generate Summary" yet.
                  </p>
                )}
              </div>
            </section>

            <ErrorPanel issues={data.issues} />

            <PivotBuilder records={joinedRecords} onExport={handlePivotExport} />
          </>
        )}
      </main>

      <footer className="mx-auto max-w-6xl px-4 py-8 text-center text-xs text-slate-400 dark:text-slate-600 sm:px-6">
        Record of Transaction Summary · Client-side only · No data leaves your browser
      </footer>
    </div>
  );
}
