// Shared type definitions for the Record of Transaction Summary app.

export type Severity = "error" | "warning";

export interface ParsedIssue {
  sheet: "PL" | "SA" | "ROT";
  row: number;
  field: string;
  bff?: string;
  message: string;
  severity: Severity;
}

export interface PLRow {
  rowNumber: number;
  block: string;
  floor: string;
  flat: string;
  bff: string;
  saleableAreaRaw: string;
  saleableAreaSqFt: number | null;
  price: number | null;
  pricePerSqFt: number | null;
  maxDiscount: number | null;
  discountedPricePerSqFt: number | null;
  rms: number | null;
  plNo: string | number | null;
  isDuplicateBff?: boolean;
}

export interface SARow {
  rowNumber: number;
  bff: string;
  saNoRaw: string;
  saNoNormalized: string | null;
  isTender: boolean | null;
  saleableAreaSqFt: number | null;
  noOfRooms: number | null;
  filledArea: boolean;
  filledRooms: boolean;
  plMatchFound?: boolean;
  isDuplicateBff?: boolean;
}

export interface ROTRow {
  rowNumber: number;
  dateOfPASP: Date | null;
  dateOfASP: Date | null;
  dateOfTermination: Date | null;
  block: string;
  floor: string;
  flat: string;
  bff: string;
  transactionPrice: number | null;
  termsOfPayment: string;
  noOfRms: number | null;
  rotPricePerSqFt: number | null;
  isTender: boolean | null;
  saMatchFound?: boolean;
  invalidDateFields?: string[];
}

export interface ParsedWorkbookData {
  plRows: PLRow[];
  saRows: SARow[];
  rotRows: ROTRow[];
  issues: ParsedIssue[];
}

export interface WorkbookStats {
  plRowCount: number;
  saRowCount: number;
  rotRowCount: number;
  plCount: number;
  tenderCount: number;
  soldCount: number;
  errorCount: number;
  warningCount: number;
}
