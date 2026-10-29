import { useMemo, useState } from "react";
import type { ParsedIssue } from "../lib/excel/types";

export default function ErrorPanel({ issues }: { issues: ParsedIssue[] }) {
  const [filter, setFilter] = useState<"all" | "error" | "warning">("all");

  const filtered = useMemo(() => {
    if (filter === "all") return issues;
    return issues.filter((i) => i.severity === filter);
  }, [issues, filter]);

  const errorCount = issues.filter((i) => i.severity === "error").length;
  const warningCount = issues.filter((i) => i.severity === "warning").length;

  if (issues.length === 0) {
    return (
      <section className="rounded-2xl border border-emerald-200 bg-emerald-50 p-5 text-emerald-800 shadow-sm dark:border-emerald-900/50 dark:bg-emerald-500/10 dark:text-emerald-300">
        <div className="flex items-center gap-2">
          <svg className="h-5 w-5 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="m5 13 4 4L19 7" />
          </svg>
          <p className="text-sm font-medium">No errors or warnings found. 沒有發現錯誤或警告。</p>
        </div>
      </section>
    );
  }

  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-900">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">Errors &amp; Warnings 錯誤及警告</h2>
        <div className="flex gap-1.5 text-xs">
          {(
            [
              { key: "all", label: `All (${issues.length})` },
              { key: "error", label: `Errors (${errorCount})` },
              { key: "warning", label: `Warnings (${warningCount})` },
            ] as const
          ).map((tab) => (
            <button
              key={tab.key}
              onClick={() => setFilter(tab.key)}
              className={`rounded-full px-3 py-1 font-medium transition ${
                filter === tab.key
                  ? "bg-indigo-600 text-white"
                  : "bg-slate-100 text-slate-600 hover:bg-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:hover:bg-slate-700"
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      <div className="max-h-96 scroll-thin overflow-y-auto rounded-xl border border-slate-200 dark:border-slate-800">
        <table className="w-full border-collapse text-left text-xs">
          <thead className="sticky top-0 bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-300">
            <tr>
              <th className="px-3 py-2 font-semibold">Sheet</th>
              <th className="px-3 py-2 font-semibold">Row</th>
              <th className="px-3 py-2 font-semibold">Field</th>
              <th className="px-3 py-2 font-semibold">BFF</th>
              <th className="px-3 py-2 font-semibold">Message</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((issue, idx) => (
              <tr
                key={idx}
                className={`border-t border-slate-100 dark:border-slate-800 ${
                  issue.severity === "error"
                    ? "bg-red-50 dark:bg-red-500/10"
                    : "bg-amber-50 dark:bg-amber-500/10"
                }`}
              >
                <td className="px-3 py-2 font-medium text-slate-700 dark:text-slate-200">{issue.sheet}</td>
                <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{issue.row}</td>
                <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{issue.field}</td>
                <td className="px-3 py-2 font-mono text-slate-600 dark:text-slate-300">{issue.bff ?? "—"}</td>
                <td
                  className={`px-3 py-2 ${
                    issue.severity === "error"
                      ? "text-red-700 dark:text-red-300"
                      : "text-amber-700 dark:text-amber-300"
                  }`}
                >
                  {issue.severity === "error" ? "ERROR: " : "WARNING: "}
                  {issue.message}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
