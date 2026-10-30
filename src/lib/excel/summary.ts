// Builds nine Summary tables with Excel formulas.
import ExcelJS from "exceljs";
import { XLSX_COLORS } from "./constants";
import { findSheet } from "./readWorkbook";
import type { PLRow, ROTRow, SARow } from "./types";

const BORDER: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: XLSX_COLORS.borderColor } },
  left: { style: "thin", color: { argb: XLSX_COLORS.borderColor } },
  bottom: { style: "thin", color: { argb: XLSX_COLORS.borderColor } },
  right: { style: "thin", color: { argb: XLSX_COLORS.borderColor } },
};
function styleTitle(cell: ExcelJS.Cell) {
  cell.font = { bold: true, size: 13, color: { argb: XLSX_COLORS.titleFont } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.titleFill } };
  cell.alignment = { vertical: "middle", horizontal: "left" };
}
function styleHeader(cell: ExcelJS.Cell) {
  cell.font = { bold: true, color: { argb: XLSX_COLORS.headerFont } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.headerFill } };
  cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  cell.border = BORDER;
}
function styleData(cell: ExcelJS.Cell, align: "left" | "center" | "right" = "center") {
  cell.border = BORDER;
  cell.alignment = { vertical: "middle", horizontal: align };
}
function styleTotal(cell: ExcelJS.Cell) {
  cell.font = { bold: true };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.totalFill } };
  cell.border = BORDER;
  cell.alignment = { vertical: "middle", horizontal: "center" };
}
function roomLabel(n: number): string { return n === 0 ? "0 (Studio)" : String(n); }
function uniqueSorted<T extends string | number>(values: T[]): T[] {
  return Array.from(new Set(values)).sort((a, b) => {
    if (typeof a === "number" && typeof b === "number") return a - b;
    return String(a).localeCompare(String(b), undefined, { numeric: true });
  });
}
function addRotHelperColumn(rotWs: ExcelJS.Worksheet, rotLastRow: number) {
  const header = rotWs.getRow(1).getCell(13);
  header.value = "SA No. (Join Helper)";
  header.font = { bold: true, italic: true, color: { argb: XLSX_COLORS.helperFont } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.helperFill } };
  header.border = BORDER;
  rotWs.getColumn(13).width = 22;
  for (let r = 2; r <= rotLastRow; r++) {
    const cell = rotWs.getRow(r).getCell(13);
    cell.value = { formula: `IFERROR(SUBSTITUTE(UPPER(VLOOKUP($G${r},SA!$A:$B,2,FALSE))," ",""),"")` } as ExcelJS.CellFormulaValue;
    cell.font = { italic: true, color: { argb: XLSX_COLORS.helperFont } };
  }
}
function priceFormula(cell: ExcelJS.Cell, formula: string, total = false) {
  cell.value = { formula } as ExcelJS.CellFormulaValue;
  cell.numFmt = "0.00";
  if (total) styleTotal(cell); else styleData(cell, "right");
}
function tableHeader(ws: ExcelJS.Worksheet, r: number, title: string, headers: string[]): number {
  ws.mergeCells(r, 1, r, headers.length);
  ws.getCell(r, 1).value = title;
  styleTitle(ws.getCell(r, 1));
  headers.forEach((h, i) => {
    const cell = ws.getCell(r + 1, i + 1);
    cell.value = h;
    styleHeader(cell);
  });
  return r + 2;
}
export interface SummaryBuildResult { sheetName: string; }
export function buildSummarySheet(workbook: ExcelJS.Workbook, plRows: PLRow[], saRows: SARow[], rotRows: ROTRow[]): SummaryBuildResult {
  const plWs = findSheet(workbook, "PL");
  const saWs = findSheet(workbook, "SA");
  const rotWs = findSheet(workbook, "ROT");
  if (!plWs || !saWs || !rotWs) throw new Error("PL, SA, and ROT sheets are required to build the Summary sheet.");
  const plLast = Math.max(2, plWs.rowCount);
  const saLast = Math.max(2, saWs.rowCount);
  const rotLast = Math.max(2, rotWs.rowCount);
  addRotHelperColumn(rotWs, rotLast);
  const existing = workbook.getWorksheet("Summary");
  if (existing) workbook.removeWorksheet(existing.id);
  const ws = workbook.addWorksheet("Summary", { properties: { defaultColWidth: 20 } });
  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 24;
  ws.getColumn(3).width = 24;
  ws.getColumn(4).width = 24;
  const PL_K = `PL!$K$2:$K$${plLast}`;
  const PL_I = `PL!$I$2:$I$${plLast}`;
  const SA_B = `SA!$B$2:$B$${saLast}`;
  const SA_D = `SA!$D$2:$D$${saLast}`;
  const ROT_A = `ROT!$A$2:$A$${rotLast}`;
  const ROT_K = `ROT!$K$2:$K$${rotLast}`;
  const ROT_L = `ROT!$L$2:$L$${rotLast}`;
  const ROT_M = `ROT!$M$2:$M$${rotLast}`;
  const rotAvgHeader = "Avg. ROT 呎價 (per sq.ft.)";
  let r = 1;
  const plNos = uniqueSorted(plRows.map(p => p.plNo).filter((v): v is string | number => v !== null));
  r = tableHeader(ws, r, "Table 1 · PL: Avg Discounted Price 各 PL 平均折實呎價", ["PL No.", "Avg. Discounted 呎價 (per sq.ft.)", rotAvgHeader]);
  for (const plNo of plNos) {
    ws.getCell(r, 1).value = plNo;
    styleData(ws.getCell(r, 1));
    priceFormula(ws.getCell(r, 2), `IFERROR(AVERAGEIF(${PL_K},A${r},${PL_I}),"N/A")`);
    // ROT M is normalized text, matching the SA label rather than ROT room count.
    priceFormula(ws.getCell(r, 3), `IFERROR(AVERAGEIF(${ROT_M},UPPER(SUBSTITUTE(A${r}&""," ","")),${ROT_L}),"N/A")`);
    r++;
  }
  if (plNos.length === 0) { ws.getCell(r, 1).value = "No PL data"; styleData(ws.getCell(r++, 1), "left"); }
  r += 2;
  const tenderLabels = uniqueSorted(saRows.filter(s => s.isTender === true && s.saNoNormalized).map(s => s.saNoNormalized as string));
  r = tableHeader(ws, r, "Table 2 · Tender: Avg Discounted Price 各 Tender 平均折實呎價", ["Tender No.", "Avg. Price (per sq.ft.)", rotAvgHeader]);
  for (const label of tenderLabels) {
    ws.getCell(r, 1).value = label;
    styleData(ws.getCell(r, 1));
    const formula = `IFERROR(AVERAGEIF(${ROT_M},A${r},${ROT_L}),"N/A")`;
    // Preserve the existing Tender average; both columns use ROT L by design.
    priceFormula(ws.getCell(r, 2), formula);
    priceFormula(ws.getCell(r, 3), formula);
    r++;
  }
  if (tenderLabels.length === 0) { ws.getCell(r, 1).value = "No Tender data"; styleData(ws.getCell(r++, 1), "left"); }
  r += 2;
  r = tableHeader(ws, r, "Table 3 · Overall: Avg Discounted Price 整體平均折實呎價", ["Overall", "Avg. Price (per sq.ft.)", rotAvgHeader]);
  ws.getCell(r, 1).value = "All PL + Tender";
  styleData(ws.getCell(r, 1), "left");
  priceFormula(ws.getCell(r, 2), `IFERROR((SUM(${PL_I})+SUMIF(${ROT_M},"*T*",${ROT_L}))/(COUNT(${PL_I})+COUNTIF(${ROT_M},"*T*")),"N/A")`);
  priceFormula(ws.getCell(r, 3), `IFERROR(AVERAGE(${ROT_L}),"N/A")`);
  r += 3;
  const plRoomsSet = uniqueSorted([
    ...saRows.filter(s => s.isTender === false && s.noOfRooms !== null).map(s => s.noOfRooms as number),
    ...rotRows.filter(rr => rr.isTender === false && rr.noOfRms !== null).map(rr => rr.noOfRms as number),
  ]);
  r = writeUnitsSoldTable(ws, r, "Table 4 · PL-Units Sold PL 去貨情況", plRoomsSet, SA_D, SA_B, ROT_K, ROT_M, ROT_A, "<>*T*") + 2;
  const tenderRoomsSet = uniqueSorted([
    ...saRows.filter(s => s.isTender === true && s.noOfRooms !== null).map(s => s.noOfRooms as number),
    ...rotRows.filter(rr => rr.isTender === true && rr.noOfRms !== null).map(rr => rr.noOfRms as number),
  ]);
  r = writeUnitsSoldTable(ws, r, "Table 5 · Tender-Units Sold Tender 去貨情況", tenderRoomsSet, SA_D, SA_B, ROT_K, ROT_M, ROT_A, "*T*") + 2;
  const plPriceRooms = uniqueSorted(rotRows.filter(rr => rr.isTender === false && rr.noOfRms !== null).map(rr => rr.noOfRms as number));
  const tenderPriceRooms = uniqueSorted(rotRows.filter(rr => rr.isTender === true && rr.noOfRms !== null).map(rr => rr.noOfRms as number));
  const overallPriceRooms = uniqueSorted([...plPriceRooms, ...tenderPriceRooms]);
  r = writePriceTable(ws, r, "Table 6 · Price PL + Tender 呎價分析", overallPriceRooms, ROT_K, ROT_M, ROT_L, null) + 2;
  r = writePriceTable(ws, r, "Table 7 · PL-Price PL 呎價分析", plPriceRooms, ROT_K, ROT_M, ROT_L, "<>*T*") + 2;
  r = writePriceTable(ws, r, "Table 8 · Tender-Price Tender 呎價分析", tenderPriceRooms, ROT_K, ROT_M, ROT_L, "*T*") + 2;
  r = tableHeader(ws, r, "Table 9 · Overall-Price 整體呎價分析", ["Metric", "Price (per sq.ft.)"]);
  const overallMetrics: Array<[string, string]> = [
    ["Min. Price (per sq. ft.)", `IFERROR(MIN(${ROT_L}),"N/A")`],
    ["Max. Price (per sq. ft.)", `IFERROR(MAX(${ROT_L}),"N/A")`],
    ["Avg. Price (per sq. ft.)", `IFERROR(AVERAGE(${ROT_L}),"N/A")`],
  ];
  for (const [label, formula] of overallMetrics) {
    ws.getCell(r, 1).value = label;
    styleData(ws.getCell(r, 1), "left");
    priceFormula(ws.getCell(r, 2), formula);
    r++;
  }
  ws.views = [{ state: "frozen", ySplit: 0 }];
  return { sheetName: "Summary" };
}
function writeUnitsSoldTable(ws: ExcelJS.Worksheet, startRow: number, title: string, rooms: number[], SA_D: string, SA_B: string, ROT_K: string, ROT_M: string, ROT_A: string, tenderCriteria: string): number {
  let r = tableHeader(ws, startRow, title, ["No. of Rms", "Released", "Sold", "Take-up rate (%)"]);
  const firstDataRow = r;
  for (const room of rooms) {
    ws.getCell(r, 1).value = roomLabel(room);
    styleData(ws.getCell(r, 1));
    ws.getCell(r, 2).value = { formula: `COUNTIFS(${SA_D},${room},${SA_B},"${tenderCriteria}")` } as ExcelJS.CellFormulaValue;
    ws.getCell(r, 3).value = { formula: `COUNTIFS(${ROT_K},${room},${ROT_M},"${tenderCriteria}",${ROT_A},"<>")` } as ExcelJS.CellFormulaValue;
    ws.getCell(r, 4).value = { formula: `IFERROR(C${r}/B${r},0)` } as ExcelJS.CellFormulaValue;
    ws.getCell(r, 4).numFmt = "0.0%";
    [2, 3, 4].forEach(c => styleData(ws.getCell(r, c), "right"));
    r++;
  }
  const lastDataRow = r - 1;
  ws.getCell(r, 1).value = "Total";
  if (rooms.length > 0) {
    ws.getCell(r, 2).value = { formula: `SUM(B${firstDataRow}:B${lastDataRow})` } as ExcelJS.CellFormulaValue;
    ws.getCell(r, 3).value = { formula: `SUM(C${firstDataRow}:C${lastDataRow})` } as ExcelJS.CellFormulaValue;
    ws.getCell(r, 4).value = { formula: `IFERROR(C${r}/B${r},0)` } as ExcelJS.CellFormulaValue;
    ws.getCell(r, 4).numFmt = "0.0%";
  }
  [1, 2, 3, 4].forEach(c => styleTotal(ws.getCell(r, c)));
  return r + 1;
}
function writePriceTable(ws: ExcelJS.Worksheet, startRow: number, title: string, rooms: number[], ROT_K: string, ROT_M: string, ROT_L: string, tenderCriteria: string | null): number {
  let r = tableHeader(ws, startRow, title, ["No. of Rms", "Min. Price (per sq.ft.)", "Max. Price (per sq.ft.)", "Avg. Price (per sq.ft.)"]);
  const firstDataRow = r;
  // null means all ROT records: do not filter on the join-helper column.
  const tenderFilter = tenderCriteria === null ? "" : `,${ROT_M},"${tenderCriteria}"`;
  for (const room of rooms) {
    ws.getCell(r, 1).value = roomLabel(room);
    styleData(ws.getCell(r, 1));
    priceFormula(ws.getCell(r, 2), `IFERROR(_xlfn.MINIFS(${ROT_L},${ROT_K},${room}${tenderFilter}),"N/A")`);
    priceFormula(ws.getCell(r, 3), `IFERROR(_xlfn.MAXIFS(${ROT_L},${ROT_K},${room}${tenderFilter}),"N/A")`);
    const avgFunction = tenderCriteria === null ? "AVERAGEIFS" : "_xlfn.AVERAGEIFS";
    priceFormula(ws.getCell(r, 4), `IFERROR(${avgFunction}(${ROT_L},${ROT_K},${room}${tenderFilter}),"N/A")`);
    r++;
  }
  const lastDataRow = r - 1;
  ws.getCell(r, 1).value = "Total";
  [1, 2, 3, 4].forEach(c => styleTotal(ws.getCell(r, c)));
  if (rooms.length > 0) {
    priceFormula(ws.getCell(r, 2), `IFERROR(MIN(B${firstDataRow}:B${lastDataRow}),"N/A")`, true);
    priceFormula(ws.getCell(r, 3), `IFERROR(MAX(C${firstDataRow}:C${lastDataRow}),"N/A")`, true);
    const avgFormula = tenderCriteria === null ? `IFERROR(AVERAGE(${ROT_L}),"N/A")` : `IFERROR(AVERAGEIFS(${ROT_L},${ROT_M},"${tenderCriteria}"),"N/A")`;
    priceFormula(ws.getCell(r, 4), avgFormula, true);
  }
  return r + 1;
}
