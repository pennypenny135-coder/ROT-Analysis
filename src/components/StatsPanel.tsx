import type { WorkbookStats } from "../lib/excel/types";

interface StatCardDef {
  label: string;
  labelZh: string;
  value: number;
  accent: string;
}

export default function StatsPanel({ stats }: { stats: WorkbookStats }) {
  const cards: StatCardDef[] = [
    { label: "PL rows", labelZh: "PL 列數", value: stats.plRowCount, accent: "text-slate-100" },
    { label: "SA rows", labelZh: "SA 列數", value: stats.saRowCount, accent: "text-slate-100" },
    { label: "ROT rows", labelZh: "ROT 列數", value: stats.rotRowCount, accent: "text-slate-100" },
    { label: "PL count", labelZh: "PL 數目", value: stats.plCount, accent: "text-sky-400" },
    { label: "Tender count", labelZh: "Tender 數目", value: stats.tenderCount, accent: "text-violet-400" },
    { label: "Sold count", labelZh: "已售數目", value: stats.soldCount, accent: "text-emerald-400" },
    { label: "Errors", labelZh: "錯誤", value: stats.errorCount, accent: "text-red-400" },
    { label: "Warnings", labelZh: "警告", value: stats.warningCount, accent: "text-amber-400" },
  ];

  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-900">
      <h2 className="mb-4 text-base font-semibold text-slate-900 dark:text-slate-100">Summary Statistics 統計摘要</h2>
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
        {cards.map((c) => (
          <div
            key={c.label}
            className="rounded-xl border border-slate-200 bg-slate-50 p-4 text-center dark:border-slate-800 dark:bg-slate-950/60"
          >
            <p className={`text-2xl font-bold ${c.accent === "text-slate-100" ? "text-slate-900 dark:text-slate-100" : c.accent}`}>
              {c.value}
            </p>
            <p className="mt-1 text-xs font-medium text-slate-500 dark:text-slate-400">{c.label}</p>
            <p className="text-[11px] text-slate-400 dark:text-slate-500">{c.labelZh}</p>
          </div>
        ))}
      </div>
    </section>
  );
}
