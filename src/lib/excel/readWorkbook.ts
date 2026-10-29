import ExcelJS from "exceljs";
import { buildBFF } from "./bff";
import {
  extractSqFtFromAreaText,
  normalizeMaxDiscount,
  normalizeSANo,
  parseFlexibleDate,
  toNumber,
  toTrimmedString,
} from "./normalize";
import type { ParsedIssue, ParsedWorkbookData, PLRow, ROTRow, SARow } from "./types";

/** Reads the resolved value of a cell, unwrapping formula results, rich text, and hyperlinks. */
export function getCellValue(cell: ExcelJS.Cell | undefined): unknown {
  if (!cell) return null;
  const v = cell.value as unknown;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    const obj = v as Record<string, unknown>;
    if ("result" in obj) {
      const result = obj.result;
      if (result && typeof result === "object" && "error" in (result as Record<string, unknown>)) {
        return null;
      }
      return result ?? null;
    }
    if ("richText" in obj && Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text: string }>).map((rt) => rt.text).join("");
    }
    if ("text" in obj) {
      return (obj as { text: unknown }).text;
    }
    return null;
  }
  return v;
}

/** Finds a worksheet case-insensitively & whitespace-insensitively by name. */
export function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
  const target = name.trim().toLowerCase();
  return workbook.worksheets.find((ws) => ws.name.trim().toLowerCase() === target);
}

function rowAllBlank(values: unknown[]): boolean {
  return values.every((v) => v === null || v === undefined || v === "");
}

export function parsePLSheet(ws: ExcelJS.Worksheet): { rows: PLRow[]; issues: ParsedIssue[] } {
  const rows: PLRow[] = [];
  const issues: ParsedIssue[] = [];
  const rowCount = ws.rowCount;

  for (let r = 2; r <= rowCount; r++) {
    const row = ws.getRow(r);
    const block = getCellValue(row.getCell(1));
    const floor = getCellValue(row.getCell(2));
    const flat = getCellValue(row.getCell(3));
    const bffRaw = getCellValue(row.getCell(4));
    const areaText = getCellValue(row.getCell(5));
    const price = toNumber(getCellValue(row.getCell(6)));
    const pricePerSqFt = toNumber(getCellValue(row.getCell(7)));
    const maxDiscountRaw = getCellValue(row.getCell(8));
    const discountedPricePerSqFt = toNumber(getCellValue(row.getCell(9)));
    const rmsRaw = getCellValue(row.getCell(10));
    const plNoRaw = getCellValue(row.getCell(11));

    if (
      rowAllBlank([block, floor, flat, bffRaw, areaText, price, pricePerSqFt, maxDiscountRaw, discountedPricePerSqFt, rmsRaw, plNoRaw])
    ) {
      continue;
    }

    const bff = toTrimmedString(bffRaw) || buildBFF(block, flat, floor);
    const saleableAreaSqFt = extractSqFtFromAreaText(areaText);
    const maxDiscount = normalizeMaxDiscount(maxDiscountRaw);
    const rmsNum = rmsRaw === null || rmsRaw === undefined || rmsRaw === "" ? null : Number(rmsRaw);
    const plNo = plNoRaw === null || plNoRaw === undefined || plNoRaw === "" ? null : (plNoRaw as string | number);

    const plRow: PLRow = {
      rowNumber: r,
      block: toTrimmedString(block),
      floor: toTrimmedString(floor),
      flat: toTrimmedString(flat),
      bff,
      saleableAreaRaw: areaText === null || areaText === undefined ? "" : String(areaText),
      saleableAreaSqFt,
      price,
      pricePerSqFt,
      maxDiscount,
      discountedPricePerSqFt,
      rms: rmsNum !== null && Number.isFinite(rmsNum) ? rmsNum : null,
      plNo,
    };

    if (plNo === null) {
      issues.push({ sheet: "PL", row: r, field: "PL No. (Column K)", bff, message: "PL No. is empty.", severity: "error" });
    }
    if (saleableAreaSqFt === null) {
      issues.push({
        sheet: "PL",
        row: r,
        field: "Saleable Area (Column E)",
        bff,
        message: "Unable to extract saleable area in sq. ft. from Column E.",
        severity: "error",
      });
    }

    rows.push(plRow);
  }

  // Duplicate BFF detection
  const seen = new Map<string, PLRow[]>();
  for (const row of rows) {
    const key = row.bff.toUpperCase().replace(/\s+/g, "");
    if (!key) continue;
    if (!seen.has(key)) seen.set(key, []);
    seen.get(key)!.push(row);
  }
  for (const group of seen.values()) {
    if (group.length > 1) {
      for (const row of group) {
        row.isDuplicateBff = true;
        issues.push({
          sheet: "PL",
          row: row.rowNumber,
          field: "BFF (Column D)",
          bff: row.bff,
          message: `Duplicate BFF "${row.bff}" found in PL.`,
          severity: "error",
        });
      }
    }
  }

  return { rows, issues };
}

export function parseSASheet(ws: ExcelJS.Worksheet): { rows: SARow[]; issues: ParsedIssue[] } {
  const rows: SARow[] = [];
  const issues: ParsedIssue[] = [];
  const rowCount = ws.rowCount;

  for (let r = 2; r <= rowCount; r++) {
    const row = ws.getRow(r);
    const bffRaw = getCellValue(row.getCell(1));
    const saNoRaw = getCellValue(row.getCell(2));
    const areaRaw = getCellValue(row.getCell(3));
    const roomsRaw = getCellValue(row.getCell(4));

    if (rowAllBlank([bffRaw, saNoRaw, areaRaw, roomsRaw])) continue;

    const bff = toTrimmedString(bffRaw);
    const { normalized, isTender } = normalizeSANo(saNoRaw);

    const saRow: SARow = {
      rowNumber: r,
      bff,
      saNoRaw: saNoRaw === null || saNoRaw === undefined ? "" : String(saNoRaw),
      saNoNormalized: normalized,
      isTender,
      saleableAreaSqFt: toNumber(areaRaw),
      noOfRooms: roomsRaw === null || roomsRaw === undefined || roomsRaw === "" ? null : Number(roomsRaw),
      filledArea: false,
      filledRooms: false,
    };

    if (!bff) {
      issues.push({ sheet: "SA", row: r, field: "BFF (Column A)", message: "BFF is empty.", severity: "error" });
    }
    if (normalized === null || !/^\d+T?$/.test(normalized)) {
      issues.push({
        sheet: "SA",
        row: r,
        field: "SA No. (Column B)",
        bff,
        message: `Unable to normalize SA No. "${saRow.saNoRaw}" into PL / Tender.`,
        severity: "error",
      });
    }

    rows.push(saRow);
  }

  const seen = new Map<string, SARow[]>();
  for (const row of rows) {
    const key = row.bff.toUpperCase().replace(/\s+/g, "");
    if (!key) continue;
    if (!seen.has(key)) seen.set(key, []);
    seen.get(key)!.push(row);
  }
  for (const group of seen.values()) {
    if (group.length > 1) {
      for (const row of group) {
        row.isDuplicateBff = true;
        issues.push({
          sheet: "SA",
          row: row.rowNumber,
          field: "BFF (Column A)",
          bff: row.bff,
          message: `Duplicate BFF "${row.bff}" found in SA.`,
          severity: "error",
        });
      }
    }
  }

  return { rows, issues };
}

export function parseROTSheet(ws: ExcelJS.Worksheet): { rows: ROTRow[]; issues: ParsedIssue[] } {
  const rows: ROTRow[] = [];
  const issues: ParsedIssue[] = [];
  const rowCount = ws.rowCount;

  for (let r = 2; r <= rowCount; r++) {
    const row = ws.getRow(r);
    const paspRaw = getCellValue(row.getCell(1));
    const aspRaw = getCellValue(row.getCell(2));
    const termRaw = getCellValue(row.getCell(3));
    const block = getCellValue(row.getCell(4));
    const floor = getCellValue(row.getCell(5));
    const flat = getCellValue(row.getCell(6));
    const bffRaw = getCellValue(row.getCell(7));
    const priceRaw = getCellValue(row.getCell(8));
    const termsRaw = getCellValue(row.getCell(9));
    const rmsRaw = getCellValue(row.getCell(11));
    const psfRaw = getCellValue(row.getCell(12));

    if (
      rowAllBlank([paspRaw, aspRaw, termRaw, block, floor, flat, bffRaw, priceRaw, termsRaw, rmsRaw, psfRaw])
    ) {
      continue;
    }

    const bff = toTrimmedString(bffRaw) || buildBFF(block, flat, floor);
    const invalidDateFields: string[] = [];

    const dateOfPASP = parseFlexibleDate(paspRaw);
    if (paspRaw && !dateOfPASP) invalidDateFields.push("Date of PASP (Column A)");
    const dateOfASP = parseFlexibleDate(aspRaw);
    if (aspRaw && !dateOfASP) invalidDateFields.push("Date of ASP (Column B)");
    const dateOfTermination = parseFlexibleDate(termRaw);
    if (termRaw && !dateOfTermination) invalidDateFields.push("Date of termination (Column C)");

    const transactionPrice = toNumber(priceRaw);

    const rotRow: ROTRow = {
      rowNumber: r,
      dateOfPASP,
      dateOfASP,
      dateOfTermination,
      block: toTrimmedString(block),
      floor: toTrimmedString(floor),
      flat: toTrimmedString(flat),
      bff,
      transactionPrice,
      termsOfPayment: termsRaw === null || termsRaw === undefined ? "" : String(termsRaw),
      noOfRms: rmsRaw === null || rmsRaw === undefined || rmsRaw === "" ? null : Number(rmsRaw),
      rotPricePerSqFt: toNumber(psfRaw),
      isTender: null,
      invalidDateFields,
    };

    if (transactionPrice === null) {
      issues.push({
        sheet: "ROT",
        row: r,
        field: "Transaction Price (Column H)",
        bff,
        message: "Transaction Price is blank or not a valid number.",
        severity: "error",
      });
    }

    for (const f of invalidDateFields) {
      issues.push({ sheet: "ROT", row: r, field: f, bff, message: `Invalid date format in ${f}.`, severity: "warning" });
    }

    rows.push(rotRow);
  }

  return { rows, issues };
}

export async function readWorkbookData(workbook: ExcelJS.Workbook): Promise<
  ParsedWorkbookData & { missingSheets: string[] }
> {
  const missingSheets: string[] = [];
  const plWs = findSheet(workbook, "PL");
  const saWs = findSheet(workbook, "SA");
  const rotWs = findSheet(workbook, "ROT");

  if (!plWs) missingSheets.push("PL");
  if (!saWs) missingSheets.push("SA");
  if (!rotWs) missingSheets.push("ROT");

  const plResult = plWs ? parsePLSheet(plWs) : { rows: [], issues: [] };
  const saResult = saWs ? parseSASheet(saWs) : { rows: [], issues: [] };
  const rotResult = rotWs ? parseROTSheet(rotWs) : { rows: [], issues: [] };

  return {
    plRows: plResult.rows,
    saRows: saResult.rows,
    rotRows: rotResult.rows,
    issues: [...plResult.issues, ...saResult.issues, ...rotResult.issues],
    missingSheets,
  };
}
