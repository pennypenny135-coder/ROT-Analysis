// Applies auto-fill + error marking to PL/SA/ROT sheets, (re)builds the
// Summary sheet, and triggers a client-side download of the resulting file.
import ExcelJS from "exceljs";
import { XLSX_COLORS } from "./constants";
import { findSheet } from "./readWorkbook";
import { buildSummarySheet } from "./summary";
import type { ParsedIssue, PLRow, ROTRow, SARow } from "./types";
import type { PivotConfig, PivotResult } from "./pivot";
import { fieldByKey } from "./constants";

function markErrorCell(cell: ExcelJS.Cell, text?: string) {
  if (text !== undefined) cell.value = text;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.errorFill } };
  cell.font = { color: { argb: XLSX_COLORS.errorFont }, bold: true };
}

/**
 * Writes auto-filled values (SA area/rooms, ROT rooms/呎價) back into the
 * live PL/SA/ROT worksheets, and highlights any erroring cells in red with
 * an "ERROR" marker, mirroring the issues already surfaced in the web UI.
 */
export function applyAutoFillAndErrors(
  workbook: ExcelJS.Workbook,
  plRows: PLRow[],
  saRows: SARow[],
  rotRows: ROTRow[],
  issues: ParsedIssue[],
) {
  const plWs = findSheet(workbook, "PL");
  const saWs = findSheet(workbook, "SA");
  const rotWs = findSheet(workbook, "ROT");

  const plIssueRows = new Set(issues.filter((i) => i.sheet === "PL").map((i) => i.row));
  const saIssueRows = new Set(issues.filter((i) => i.sheet === "SA").map((i) => i.row));
  const rotIssueRows = new Set(issues.filter((i) => i.sheet === "ROT").map((i) => i.row));

  if (plWs) {
    for (const pl of plRows) {
      if (!plIssueRows.has(pl.rowNumber)) continue;
      const row = plWs.getRow(pl.rowNumber);
      if (pl.plNo === null) markErrorCell(row.getCell(11), "ERROR");
      if (pl.saleableAreaSqFt === null) markErrorCell(row.getCell(5));
      if (pl.isDuplicateBff) markErrorCell(row.getCell(4));
    }
  }

  if (saWs) {
    for (const sa of saRows) {
      const row = saWs.getRow(sa.rowNumber);
      if (sa.isTender === false) {
        if (sa.plMatchFound === false) {
          markErrorCell(row.getCell(1));
          markErrorCell(row.getCell(3), "ERROR");
          markErrorCell(row.getCell(4), "ERROR");
          continue;
        }
        if (sa.saleableAreaSqFt !== null) row.getCell(3).value = sa.saleableAreaSqFt;
        if (sa.noOfRooms !== null) row.getCell(4).value = sa.noOfRooms;
      }
      if (!saIssueRows.has(sa.rowNumber)) continue;
      if (sa.saNoNormalized === null || !/^\d+T?$/.test(sa.saNoNormalized)) {
        markErrorCell(row.getCell(2));
      }
      if (sa.isDuplicateBff) markErrorCell(row.getCell(1));
    }
  }

  if (rotWs) {
    for (const rot of rotRows) {
      const row = rotWs.getRow(rot.rowNumber);
      if (rot.saMatchFound === false) {
        markErrorCell(row.getCell(7));
        markErrorCell(row.getCell(11), "ERROR");
        markErrorCell(row.getCell(12), "ERROR");
        continue;
      }
      if (rot.noOfRms !== null) row.getCell(11).value = rot.noOfRms;
      if (rot.rotPricePerSqFt !== null) {
        row.getCell(12).value = rot.rotPricePerSqFt;
        row.getCell(12).numFmt = "0.00";
      } else {
        markErrorCell(row.getCell(12), "ERROR");
      }
      if (!rotIssueRows.has(rot.rowNumber)) continue;
      if (rot.transactionPrice === null) markErrorCell(row.getCell(8), undefined);
      for (const f of rot.invalidDateFields ?? []) {
        if (f.includes("Column A")) markErrorCell(row.getCell(1));
        if (f.includes("Column B")) markErrorCell(row.getCell(2));
        if (f.includes("Column C")) markErrorCell(row.getCell(3));
      }
    }
  }
}

export function deriveOutputFileName(originalName: string): string {
  const base = originalName.replace(/\.(xlsx|xlsm)$/i, "");
  return `${base}_Summary.xlsx`;
}

export function downloadWorkbookBuffer(buffer: ExcelJS.Buffer | ArrayBuffer, filename: string) {
  const blob = new Blob([buffer as ArrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

/** Runs the full finalize pipeline (autofill, error marking, summary rebuild) and downloads the result. */
export async function finalizeAndDownload(
  workbook: ExcelJS.Workbook,
  plRows: PLRow[],
  saRows: SARow[],
  rotRows: ROTRow[],
  issues: ParsedIssue[],
  originalFileName: string,
): Promise<void> {
  applyAutoFillAndErrors(workbook, plRows, saRows, rotRows, issues);
  buildSummarySheet(workbook, plRows, saRows, rotRows);
  const buffer = await workbook.xlsx.writeBuffer();
  downloadWorkbookBuffer(buffer, deriveOutputFileName(originalFileName));
}

/** Adds (or replaces) a "Pivot" worksheet containing the current pivot result as static values. */
export function addPivotResultSheet(workbook: ExcelJS.Workbook, config: PivotConfig, result: PivotResult) {
  const existing = workbook.getWorksheet("Pivot");
  if (existing) workbook.removeWorksheet(existing.id);

  const ws = workbook.addWorksheet("Pivot");
  const rowFieldLabel = fieldByKey(config.rowField)?.label ?? config.rowField;
  const colFieldLabel = config.colField ? fieldByKey(config.colField)?.label ?? config.colField : null;
  const valueFieldLabel = fieldByKey(config.valueField)?.label ?? config.valueField;

  ws.getCell(1, 1).value = `Pivot: Row=${rowFieldLabel}${colFieldLabel ? `, Column=${colFieldLabel}` : ""}, Value=${valueFieldLabel} (${config.aggregation})`;
  ws.getCell(1, 1).font = { bold: true, size: 12 };

  const headerRowIdx = 3;
  ws.getCell(headerRowIdx, 1).value = rowFieldLabel;
  ws.getCell(headerRowIdx, 1).font = { bold: true };
  ws.getCell(headerRowIdx, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.headerFill } };
  ws.getCell(headerRowIdx, 1).font = { bold: true, color: { argb: XLSX_COLORS.headerFont } };

  result.colKeys.forEach((colKey, i) => {
    const cell = ws.getCell(headerRowIdx, i + 2);
    cell.value = colKey;
    cell.font = { bold: true, color: { argb: XLSX_COLORS.headerFont } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.headerFill } };
  });
  const totalColIdx = result.colKeys.length + 2;
  ws.getCell(headerRowIdx, totalColIdx).value = "Total";
  ws.getCell(headerRowIdx, totalColIdx).font = { bold: true, color: { argb: XLSX_COLORS.headerFont } };
  ws.getCell(headerRowIdx, totalColIdx).fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.headerFill } };

  result.rowKeys.forEach((rowKey, rIdx) => {
    const excelRow = headerRowIdx + 1 + rIdx;
    ws.getCell(excelRow, 1).value = rowKey;
    ws.getCell(excelRow, 1).font = { bold: true };
    result.matrix[rIdx].forEach((val, cIdx) => {
      const cell = ws.getCell(excelRow, cIdx + 2);
      cell.value = val === null ? "" : Math.round(val * 100) / 100;
      cell.numFmt = "0.00";
    });
    const totalCell = ws.getCell(excelRow, totalColIdx);
    const rt = result.rowTotals[rIdx];
    totalCell.value = rt === null ? "" : Math.round(rt * 100) / 100;
    totalCell.numFmt = "0.00";
    totalCell.font = { bold: true };
  });

  const totalRowIdx = headerRowIdx + 1 + result.rowKeys.length;
  ws.getCell(totalRowIdx, 1).value = "Total";
  ws.getCell(totalRowIdx, 1).font = { bold: true };
  result.colTotals.forEach((val, cIdx) => {
    const cell = ws.getCell(totalRowIdx, cIdx + 2);
    cell.value = val === null ? "" : Math.round(val * 100) / 100;
    cell.numFmt = "0.00";
    cell.font = { bold: true };
  });
  const grandCell = ws.getCell(totalRowIdx, totalColIdx);
  grandCell.value = result.grandTotal === null ? "" : Math.round(result.grandTotal * 100) / 100;
  grandCell.numFmt = "0.00";
  grandCell.font = { bold: true };

  ws.getColumn(1).width = 24;
  for (let i = 2; i <= totalColIdx; i++) ws.getColumn(i).width = 18;
}

/** Finalizes the workbook, adds the current pivot result as a sheet, and downloads. */
export async function finalizeAndDownloadWithPivot(
  workbook: ExcelJS.Workbook,
  plRows: PLRow[],
  saRows: SARow[],
  rotRows: ROTRow[],
  issues: ParsedIssue[],
  originalFileName: string,
  pivotConfig: PivotConfig,
  pivotResult: PivotResult,
): Promise<void> {
  applyAutoFillAndErrors(workbook, plRows, saRows, rotRows, issues);
  buildSummarySheet(workbook, plRows, saRows, rotRows);
  addPivotResultSheet(workbook, pivotConfig, pivotResult);
  const buffer = await workbook.xlsx.writeBuffer();
  downloadWorkbookBuffer(buffer, deriveOutputFileName(originalFileName));
}
