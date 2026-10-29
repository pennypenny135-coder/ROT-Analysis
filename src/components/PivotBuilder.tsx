import { useMemo, useState } from "react";
import { AGGREGATIONS, JOINED_FIELDS, fieldByKey, type AggregationKey } from "../lib/excel/constants";
import { buildPivot, distinctValuesForField, type JoinedRecord, type PivotConfig, type PivotFilter, type PivotResult } from "../lib/excel/pivot";

interface PivotBuilderProps {
  records: JoinedRecord[];
  onExport: (config: PivotConfig, result: PivotResult) => void;
}

function FieldSelect({
  value,
  onChange,
  allowNone,
  id,
}: {
  value: string | null;
  onChange: (v: string | null) => void;
  allowNone?: boolean;
  id: string;
}) {
  const bySheet = useMemo(() => {
    const groups: Record<string, typeof JOINED_FIELDS> = { PL: [], SA: [], ROT: [] };
    JOINED_FIELDS.forEach((f) => groups[f.sheet].push(f));
    return groups;
  }, []);

  return (
    <select
      id={id}
      value={value ?? ""}
      onChange={(e) => onChange(e.target.value === "" ? null : e.target.value)}
      className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm text-slate-800 focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
    >
      {allowNone && <option value="">(none)</option>}
      {(["PL", "SA", "ROT"] as const).map((sheet) => (
        <optgroup key={sheet} label={sheet}>
          {bySheet[sheet].map((f) => (
            <option key={f.key} value={f.key}>
              {f.label}
            </option>
          ))}
        </optgroup>
      ))}
    </select>
  );
}

function FilterRow({
  filter,
  records,
  onChange,
  onRemove,
}: {
  filter: PivotFilter;
  records: JoinedRecord[];
  onChange: (f: PivotFilter) => void;
  onRemove: () => void;
}) {
  const field = fieldByKey(filter.field);
  const isDate = field?.type === "date";
  const options = useMemo(() => (isDate ? [] : distinctValuesForField(records, filter.field)), [records, filter.field, isDate]);

  return (
    <div className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 p-2.5 dark:border-slate-800 dark:bg-slate-950/50">
      <div className="w-44">
        <FieldSelect
          id={`filter-field-${filter.field}`}
          value={filter.field}
          onChange={(v) => v && onChange({ field: v })}
        />
      </div>
      {isDate ? (
        <div className="flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-300">
          <span>From</span>
          <input
            type="date"
            value={filter.dateFrom ?? ""}
            onChange={(e) => onChange({ ...filter, dateFrom: e.target.value || undefined })}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
          />
          <span>To</span>
          <input
            type="date"
            value={filter.dateTo ?? ""}
            onChange={(e) => onChange({ ...filter, dateTo: e.target.value || undefined })}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
          />
        </div>
      ) : (
        <select
          multiple
          value={filter.values ?? []}
          onChange={(e) => onChange({ ...filter, values: Array.from(e.target.selectedOptions).map((o) => o.value) })}
          className="h-20 min-w-[12rem] flex-1 rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
        >
          {options.map((opt) => (
            <option key={opt} value={opt}>
              {opt}
            </option>
          ))}
        </select>
      )}
      <button
        onClick={onRemove}
        type="button"
        className="ml-auto rounded-lg border border-red-200 bg-red-50 px-2 py-1 text-xs font-medium text-red-600 hover:bg-red-100 dark:border-red-900/50 dark:bg-red-500/10 dark:text-red-300 dark:hover:bg-red-500/20"
      >
        Remove 移除
      </button>
    </div>
  );
}

export default function PivotBuilder({ records, onExport }: PivotBuilderProps) {
  const [rowField, setRowField] = useState<string>("PL_PLNO");
  const [colField, setColField] = useState<string | null>("ROT_RMS");
  const [valueField, setValueField] = useState<string>("ROT_PSF");
  const [aggregation, setAggregation] = useState<AggregationKey>("average");
  const [filters, setFilters] = useState<PivotFilter[]>([]);
  const [result, setResult] = useState<PivotResult | null>(null);

  const config: PivotConfig = { rowField, colField, valueField, aggregation, filters };

  const handleGenerate = () => {
    setResult(buildPivot(records, config));
  };

  const addFilter = () => {
    const used = new Set(filters.map((f) => f.field));
    const next = JOINED_FIELDS.find((f) => !used.has(f.key)) ?? JOINED_FIELDS[0];
    setFilters([...filters, { field: next.key }]);
  };

  const rowLabel = fieldByKey(rowField)?.label ?? rowField;
  const colLabel = colField ? fieldByKey(colField)?.label ?? colField : null;
  const valueLabel = fieldByKey(valueField)?.label ?? valueField;

  const fmt = (v: number | null) => (v === null ? "—" : Number.isInteger(v) ? String(v) : v.toFixed(2));

  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-900">
      <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">Pivot Builder 樞紐分析產生器</h2>
      <p className="mt-0.5 mb-4 text-xs text-slate-500 dark:text-slate-400">
        Cross-sheet analysis joined by BFF across PL, SA and ROT. 以 BFF 跨 PL / SA / ROT 自動 join 作交叉分析。
      </p>

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300" htmlFor="row-field">
            Row field 列欄位
          </label>
          <FieldSelect id="row-field" value={rowField} onChange={(v) => v && setRowField(v)} />
        </div>
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300" htmlFor="col-field">
            Column field 欄欄位
          </label>
          <FieldSelect id="col-field" value={colField} onChange={setColField} allowNone />
        </div>
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300" htmlFor="value-field">
            Value field 數值欄位
          </label>
          <FieldSelect id="value-field" value={valueField} onChange={(v) => v && setValueField(v)} />
        </div>
        <div>
          <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300" htmlFor="agg-field">
            Aggregation 匯總方式
          </label>
          <select
            id="agg-field"
            value={aggregation}
            onChange={(e) => setAggregation(e.target.value as AggregationKey)}
            className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm text-slate-800 focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
          >
            {AGGREGATIONS.map((a) => (
              <option key={a.key} value={a.key}>
                {a.label}
              </option>
            ))}
          </select>
        </div>
      </div>

      <div className="mt-4">
        <div className="mb-2 flex items-center justify-between">
          <p className="text-xs font-medium text-slate-600 dark:text-slate-300">Filters 篩選條件</p>
          <button
            type="button"
            onClick={addFilter}
            className="rounded-lg border border-slate-300 bg-slate-50 px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            + Add filter 新增篩選
          </button>
        </div>
        {filters.length === 0 ? (
          <p className="text-xs text-slate-400 dark:text-slate-500">No filters applied. 未設定篩選條件。</p>
        ) : (
          <div className="space-y-2">
            {filters.map((f, idx) => (
              <FilterRow
                key={idx}
                filter={f}
                records={records}
                onChange={(nf) => setFilters(filters.map((old, i) => (i === idx ? nf : old)))}
                onRemove={() => setFilters(filters.filter((_, i) => i !== idx))}
              />
            ))}
          </div>
        )}
      </div>

      <div className="mt-5 flex flex-wrap gap-2">
        <button
          type="button"
          onClick={handleGenerate}
          className="inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-500"
        >
          Generate Pivot 產生樞紐表
        </button>
        {result && (
          <button
            type="button"
            onClick={() => onExport(config, result)}
            className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-300 bg-emerald-50 px-4 py-2 text-sm font-semibold text-emerald-700 transition hover:bg-emerald-100 dark:border-emerald-800 dark:bg-emerald-500/10 dark:text-emerald-300 dark:hover:bg-emerald-500/20"
          >
            Export Pivot to Excel 匯出至 Excel
          </button>
        )}
      </div>

      {result && (
        <div className="mt-5">
          <p className="mb-2 text-xs text-slate-500 dark:text-slate-400">
            Row: <span className="font-medium text-slate-700 dark:text-slate-200">{rowLabel}</span>
            {colLabel && (
              <>
                {" "}
                · Column: <span className="font-medium text-slate-700 dark:text-slate-200">{colLabel}</span>
              </>
            )}{" "}
            · Value: <span className="font-medium text-slate-700 dark:text-slate-200">{valueLabel}</span> ·{" "}
            <span className="font-medium text-slate-700 dark:text-slate-200">{aggregation}</span>
          </p>
          <div className="max-h-96 scroll-thin overflow-auto rounded-xl border border-slate-200 dark:border-slate-800">
            <table className="w-full min-w-max border-collapse text-left text-xs">
              <thead className="sticky top-0 bg-slate-800 text-white">
                <tr>
                  <th className="whitespace-nowrap px-3 py-2 font-semibold">{rowLabel}</th>
                  {result.colKeys.map((ck) => (
                    <th key={ck} className="whitespace-nowrap px-3 py-2 text-right font-semibold">
                      {ck}
                    </th>
                  ))}
                  <th className="whitespace-nowrap px-3 py-2 text-right font-semibold">Total</th>
                </tr>
              </thead>
              <tbody>
                {result.rowKeys.map((rk, rIdx) => (
                  <tr key={rk} className="border-t border-slate-100 odd:bg-white even:bg-slate-50 dark:border-slate-800 dark:odd:bg-slate-900 dark:even:bg-slate-950/50">
                    <td className="whitespace-nowrap px-3 py-2 font-medium text-slate-700 dark:text-slate-200">{rk}</td>
                    {result.matrix[rIdx].map((v, cIdx) => (
                      <td key={cIdx} className="whitespace-nowrap px-3 py-2 text-right text-slate-600 dark:text-slate-300">
                        {fmt(v)}
                      </td>
                    ))}
                    <td className="whitespace-nowrap px-3 py-2 text-right font-semibold text-slate-800 dark:text-slate-100">
                      {fmt(result.rowTotals[rIdx])}
                    </td>
                  </tr>
                ))}
                <tr className="border-t border-slate-200 bg-slate-100 font-semibold dark:border-slate-800 dark:bg-slate-800">
                  <td className="whitespace-nowrap px-3 py-2 text-slate-800 dark:text-slate-100">Total</td>
                  {result.colTotals.map((v, cIdx) => (
                    <td key={cIdx} className="whitespace-nowrap px-3 py-2 text-right text-slate-800 dark:text-slate-100">
                      {fmt(v)}
                    </td>
                  ))}
                  <td className="whitespace-nowrap px-3 py-2 text-right text-slate-900 dark:text-white">{fmt(result.grandTotal)}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}
    </section>
  );
}
