// Cross-sheet pivot builder logic. Pure, unit-testable functions.
import { normalizeBFFKey } from "./bff";
import type { AggregationKey } from "./constants";
import type { PLRow, ROTRow, SARow } from "./types";

export type JoinedRecord = Record<string, string | number | Date | null>;

/**
 * Joins PL, SA and ROT data by BFF into a flat list of records suitable for
 * pivoting. One record per SA row; if an SA BFF has one or more matching ROT
 * transactions, one record is emitted per ROT transaction (duplicated
 * PL/SA fields). SA rows without any ROT match still produce one record with
 * null ROT_* fields (useful for "released but not sold" analysis).
 */
export function buildJoinedRecords(plRows: PLRow[], saRows: SARow[], rotRows: ROTRow[]): JoinedRecord[] {
  const plIndex = new Map<string, PLRow>();
  for (const pl of plRows) {
    const key = normalizeBFFKey(pl.bff);
    if (key && !plIndex.has(key)) plIndex.set(key, pl);
  }

  const rotBySA = new Map<string, ROTRow[]>();
  for (const rot of rotRows) {
    const key = normalizeBFFKey(rot.bff);
    if (!key) continue;
    if (!rotBySA.has(key)) rotBySA.set(key, []);
    rotBySA.get(key)!.push(rot);
  }

  const records: JoinedRecord[] = [];

  for (const sa of saRows) {
    const key = normalizeBFFKey(sa.bff);
    const pl = sa.isTender === false ? plIndex.get(key) : undefined;
    const matches = rotBySA.get(key) ?? [];

    if (matches.length === 0) {
      records.push(makeRecord(pl, sa, undefined));
    } else {
      for (const rot of matches) {
        records.push(makeRecord(pl, sa, rot));
      }
    }
  }

  return records;
}

function makeRecord(pl: PLRow | undefined, sa: SARow | undefined, rot: ROTRow | undefined): JoinedRecord {
  return {
    PL_PLNO: pl?.plNo ?? null,
    PL_BLOCK: pl?.block ?? null,
    PL_FLOOR: pl?.floor ?? null,
    PL_FLAT: pl?.flat ?? null,
    PL_BFF: pl?.bff ?? sa?.bff ?? null,
    PL_PRICE: pl?.price ?? null,
    PL_DISCOUNTED_PSF: pl?.discountedPricePerSqFt ?? null,
    PL_MAX_DISCOUNT: pl?.maxDiscount ?? null,
    PL_RMS: pl?.rms ?? null,
    PL_SA_AREA: pl?.saleableAreaSqFt ?? sa?.saleableAreaSqFt ?? null,

    SA_NO: sa?.saNoNormalized ?? null,
    SA_BFF: sa?.bff ?? null,
    SA_AREA: sa?.saleableAreaSqFt ?? null,
    SA_ROOMS: sa?.noOfRooms ?? null,
    SA_TYPE: sa ? (sa.isTender ? "Tender" : "PL") : null,

    ROT_PASP: rot?.dateOfPASP ?? null,
    ROT_ASP: rot?.dateOfASP ?? null,
    ROT_TERMINATION: rot?.dateOfTermination ?? null,
    ROT_BFF: rot?.bff ?? null,
    ROT_PRICE: rot?.transactionPrice ?? null,
    ROT_TERMS: rot?.termsOfPayment ?? null,
    ROT_RMS: rot?.noOfRms ?? null,
    ROT_PSF: rot?.rotPricePerSqFt ?? null,
  };
}

export interface PivotFilter {
  field: string;
  values?: string[]; // "equals one of" (category match, stringified)
  dateFrom?: string; // yyyy-mm-dd
  dateTo?: string;
}

export function applyFilters(records: JoinedRecord[], filters: PivotFilter[]): JoinedRecord[] {
  if (filters.length === 0) return records;
  return records.filter((rec) =>
    filters.every((filter) => {
      const value = rec[filter.field];
      if (filter.dateFrom || filter.dateTo) {
        if (!(value instanceof Date)) return false;
        if (filter.dateFrom && value < new Date(filter.dateFrom)) return false;
        if (filter.dateTo && value > new Date(filter.dateTo + "T23:59:59")) return false;
        return true;
      }
      if (filter.values && filter.values.length > 0) {
        const str = value === null || value === undefined ? "" : String(value);
        return filter.values.includes(str);
      }
      return true;
    }),
  );
}

export function distinctValuesForField(records: JoinedRecord[], field: string): string[] {
  const set = new Set<string>();
  for (const rec of records) {
    const v = rec[field];
    if (v === null || v === undefined) continue;
    set.add(v instanceof Date ? v.toLocaleDateString() : String(v));
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

export interface PivotConfig {
  rowField: string;
  colField: string | null;
  valueField: string;
  aggregation: AggregationKey;
  filters: PivotFilter[];
}

export interface PivotResult {
  rowKeys: string[];
  colKeys: string[];
  matrix: (number | null)[][];
  rowTotals: (number | null)[];
  colTotals: (number | null)[];
  grandTotal: number | null;
}

function aggregate(values: number[], agg: AggregationKey): number | null {
  if (values.length === 0) return null;
  switch (agg) {
    case "sum":
      return values.reduce((a, b) => a + b, 0);
    case "average":
      return values.reduce((a, b) => a + b, 0) / values.length;
    case "count":
      return values.length;
    case "min":
      return Math.min(...values);
    case "max":
      return Math.max(...values);
    default:
      return null;
  }
}

function keyOf(value: unknown): string {
  if (value === null || value === undefined || value === "") return "(blank)";
  if (value instanceof Date) return value.toLocaleDateString();
  return String(value);
}

export function buildPivot(records: JoinedRecord[], config: PivotConfig): PivotResult {
  const filtered = applyFilters(records, config.filters);

  const rowKeySet = new Set<string>();
  const colKeySet = new Set<string>();
  const buckets = new Map<string, number[]>();

  for (const rec of filtered) {
    const rowKey = keyOf(rec[config.rowField]);
    const colKey = config.colField ? keyOf(rec[config.colField]) : "Value";
    rowKeySet.add(rowKey);
    colKeySet.add(colKey);

    const rawVal = rec[config.valueField];
    const bucketKey = `${rowKey}\u0001${colKey}`;
    if (!buckets.has(bucketKey)) buckets.set(bucketKey, []);
    if (config.aggregation === "count") {
      if (rawVal !== null && rawVal !== undefined && rawVal !== "") {
        buckets.get(bucketKey)!.push(1);
      }
    } else if (typeof rawVal === "number" && Number.isFinite(rawVal)) {
      buckets.get(bucketKey)!.push(rawVal);
    }
  }

  const rowKeys = Array.from(rowKeySet).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const colKeys = Array.from(colKeySet).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const matrix: (number | null)[][] = rowKeys.map((rowKey) =>
    colKeys.map((colKey) => aggregate(buckets.get(`${rowKey}\u0001${colKey}`) ?? [], config.aggregation)),
  );

  const rowTotals = rowKeys.map((rowKey) => {
    const vals: number[] = [];
    colKeys.forEach((colKey) => {
      const b = buckets.get(`${rowKey}\u0001${colKey}`);
      if (b) vals.push(...b);
    });
    return aggregate(vals, config.aggregation);
  });

  const colTotals = colKeys.map((colKey) => {
    const vals: number[] = [];
    rowKeys.forEach((rowKey) => {
      const b = buckets.get(`${rowKey}\u0001${colKey}`);
      if (b) vals.push(...b);
    });
    return aggregate(vals, config.aggregation);
  });

  const allVals: number[] = [];
  buckets.forEach((v) => allVals.push(...v));
  const grandTotal = aggregate(allVals, config.aggregation);

  return { rowKeys, colKeys, matrix, rowTotals, colTotals, grandTotal };
}
