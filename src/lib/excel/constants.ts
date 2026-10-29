// Shared constants: Excel styling palette and pivot field metadata.

export const XLSX_COLORS = {
  headerFill: "FF1E293B", // slate-800
  headerFont: "FFFFFFFF",
  titleFont: "FF0F172A",
  titleFill: "FFCBD5E1", // slate-300
  totalFill: "FFE2E8F0", // slate-200
  errorFill: "FFFFC7CE",
  errorFont: "FF9C0006",
  helperFill: "FFF1F5F9",
  helperFont: "FF64748B",
  borderColor: "FF94A3B8",
};

export type FieldType = "string" | "number" | "date";

export interface FieldDef {
  key: string;
  label: string;
  sheet: "PL" | "SA" | "ROT";
  type: FieldType;
}

export const JOINED_FIELDS: FieldDef[] = [
  { key: "PL_PLNO", label: "PL No.", sheet: "PL", type: "string" },
  { key: "PL_BLOCK", label: "Block", sheet: "PL", type: "string" },
  { key: "PL_FLOOR", label: "Floor", sheet: "PL", type: "string" },
  { key: "PL_FLAT", label: "Flat", sheet: "PL", type: "string" },
  { key: "PL_BFF", label: "BFF", sheet: "PL", type: "string" },
  { key: "PL_PRICE", label: "Price ($)", sheet: "PL", type: "number" },
  { key: "PL_DISCOUNTED_PSF", label: "Discounted 呎價 Price per sq.ft.", sheet: "PL", type: "number" },
  { key: "PL_MAX_DISCOUNT", label: "Max Discount (%)", sheet: "PL", type: "number" },
  { key: "PL_RMS", label: "Rms", sheet: "PL", type: "number" },
  { key: "PL_SA_AREA", label: "Saleable Area (sq. ft.)", sheet: "PL", type: "number" },

  { key: "SA_NO", label: "SA No.", sheet: "SA", type: "string" },
  { key: "SA_BFF", label: "BFF", sheet: "SA", type: "string" },
  { key: "SA_AREA", label: "Saleable Area (sq. ft.)", sheet: "SA", type: "number" },
  { key: "SA_ROOMS", label: "No. of Rooms", sheet: "SA", type: "number" },
  { key: "SA_TYPE", label: "Type (PL / Tender)", sheet: "SA", type: "string" },

  { key: "ROT_PASP", label: "Date of PASP", sheet: "ROT", type: "date" },
  { key: "ROT_ASP", label: "Date of ASP", sheet: "ROT", type: "date" },
  { key: "ROT_TERMINATION", label: "Date of termination", sheet: "ROT", type: "date" },
  { key: "ROT_BFF", label: "BFF", sheet: "ROT", type: "string" },
  { key: "ROT_PRICE", label: "Transaction Price", sheet: "ROT", type: "number" },
  { key: "ROT_TERMS", label: "Terms of Payment", sheet: "ROT", type: "string" },
  { key: "ROT_RMS", label: "No. of Rms", sheet: "ROT", type: "number" },
  { key: "ROT_PSF", label: "ROT 呎價", sheet: "ROT", type: "number" },
];

export function fieldByKey(key: string): FieldDef | undefined {
  return JOINED_FIELDS.find((f) => f.key === key);
}

export const AGGREGATIONS = [
  { key: "sum", label: "Sum" },
  { key: "average", label: "Average" },
  { key: "count", label: "Count" },
  { key: "min", label: "Min" },
  { key: "max", label: "Max" },
] as const;

export type AggregationKey = (typeof AGGREGATIONS)[number]["key"];
