// Builds the 8 fixed Summary tables into the workbook, using Excel formulas
// (AVERAGEIF / COUNTIFS / MINIFS / MAXIFS / AVERAGEIFS) wherever possible.
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

function roomLabel(n: number): string {
  return n === 0 ? "0 (Studio)" : String(n);
}

function uniqueSorted<T extends string | number>(values: T[]): T[] {
  return Array.from(new Set(values)).sort((a, b) => {
    if (typeof a === "number" && typeof b === "number") return a - b;
    return String(a).localeCompare(String(b), undefined, { numeric: true });
  });
}

/** Adds a hidden-ish helper column to ROT that resolves each row's normalized SA No. (PL or Tender label) via formula, used by Summary aggregation formulas. */
function addRotHelperColumn(rotWs: ExcelJS.Worksheet, rotLastRow: number) {
  const header = rotWs.getRow(1).getCell(13);
  header.value = "SA No. (Join Helper)";
  header.font = { bold: true, italic: true, color: { argb: XLSX_COLORS.helperFont } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XLSX_COLORS.helperFill } };
  header.border = BORDER;
  rotWs.getColumn(13).width = 22;

  for (let r = 2; r <= rotLastRow; r++) {
    const cell = rotWs.getRow(r).getCell(13);
    cell.value = {
      formula: `IFERROR(SUBSTITUTE(UPPER(VLOOKUP($G${r},SA!$A:$B,2,FALSE))," ",""),"")`,
    } as ExcelJS.CellFormulaValue;
    cell.font = { italic: true, color: { argb: XLSX_COLORS.helperFont } };
  }
}

export interface SummaryBuildResult {
  sheetName: string;
}

export function buildSummarySheet(workbook: ExcelJS.Workbook, plRows: PLRow[], saRows: SARow[], rotRows: ROTRow[]): SummaryBuildResult {
  const plWs = findSheet(workbook, "PL");
  const saWs = findSheet(workbook, "SA");
  const rotWs = findSheet(workbook, "ROT");
  if (!plWs || !saWs || !rotWs) {
    throw new Error("PL, SA, and ROT sheets are required to build the Summary sheet.");
  }

  const plLast = Math.max(2, plWs.rowCount);
  const saLast = Math.max(2, saWs.rowCount);
  const rotLast = Math.max(2, rotWs.rowCount);

  addRotHelperColumn(rotWs, rotLast);

  const existing = workbook.getWorksheet("Summary");
  if (existing) workbook.removeWorksheet(existing.id);

  const ws = workbook.addWorksheet("Summary", { properties: { defaultColWidth: 20 } });
  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 20;
  ws.getColumn(3).width = 20;
  ws.getColumn(4).width = 22;

  const PL_K = `PL!$K$2:$K$${plLast}`;
  const PL_I = `PL!$I$2:$I$${plLast}`;
  const SA_B = `SA!$B$2:$B$${saLast}`;
  const SA_D = `SA!$D$2:$D$${saLast}`;
  const ROT_A = `ROT!$A$2:$A$${rotLast}`;
  const ROT_K = `ROT!$K$2:$K$${rotLast}`;
  const ROT_L = `ROT!$L$2:$L$${rotLast}`;
  const ROT_M = `ROT!$M$2:$M$${rotLast}`;

  let r = 1;

  // ---------------- Table 1: PL - Avg Discounted Price ----------------
  const plNos = uniqueSorted(plRows.map((p) => p.plNo).filter((v): v is string | number => v !== null));
  ws.mergeCells(r, 1, r, 2);
  styleTitle(ws.getCell(r, 1));
  ws.getCell(r, 1).value = "Table 1 · PL: Avg Discounted Price 各 PL 平均折實呎價";
  r++;
  [1, 2].forEach((i) => styleHeader(ws.getCell(r, i)));
  ws.getRow(r).values = ["PL No.", "Avg. Discounted 呎價 (per sq.ft.)"];
  r++;
  const table1Start = r;
  for (const plNo of plNos) {
    ws.getCell(r, 1).value = plNo;
    styleData(ws.getCell(r, 1));
    const cell = ws.getCell(r, 2);
    cell.value = { formula: `IFERROR(AVERAGEIF(${PL_K},A${r},${PL_I}),"N/A")` } as ExcelJS.CellFormulaValue;
    cell.numFmt = "0.00";
    styleData(cell, "right");
    r++;
  }
  if (plNos.length === 0) {
    ws.getCell(r, 1).value = "No PL data";
    styleData(ws.getCell(r, 1), "left");
    r++;
  }
  void table1Start;
  r += 2;

  // ---------------- Table 2: Tender - Avg Discounted Price ----------------
  const tenderLabels = uniqueSorted(
    saRows.filter((s) => s.isTender === true && s.saNoNormalized).map((s) => s.saNoNormalized as string),
  );
  ws.mergeCells(r, 1, r, 2);
  styleTitle(ws.getCell(r, 1));
  ws.getCell(r, 1).value = "Table 2 · Tender: Avg Discounted Price 各 Tender 平均折實呎價";
  r++;
  ws.getRow(r).values = ["Tender No.", "Avg. Price (per sq.ft.)"];
  [1, 2].forEach((c) => styleHeader(ws.getCell(r, c)));
  r++;
  for (const label of tenderLabels) {
    ws.getCell(r, 1).value = label;
    styleData(ws.getCell(r, 1));
    const cell = ws.getCell(r, 2);
    cell.value = { formula: `IFERROR(AVERAGEIF(${ROT_M},"${label}",${ROT_L}),"N/A")` } as ExcelJS.CellFormulaValue;
    cell.numFmt = "0.00";
    styleData(cell, "right");
    r++;
  }
  if (tenderLabels.length === 0) {
    ws.getCell(r, 1).value = "No Tender data";
    styleData(ws.getCell(r, 1), "left");
    r++;
  }
  r += 2;

  // ---------------- Table 3: Overall - Avg Discounted Price ----------------
  ws.mergeCells(r, 1, r, 2);
  styleTitle(ws.getCell(r, 1));
  ws.getCell(r, 1).value = "Table 3 · Overall: Avg Discounted Price 整體平均折實呎價";
  r++;
  ws.getRow(r).values = ["Overall", "Avg. Price (per sq.ft.)"];
  [1, 2].forEach((c) => styleHeader(ws.getCell(r, c)));
  r++;
  ws.getCell(r, 1).value = "All PL + Tender";
  styleData(ws.getCell(r, 1), "left");
  {
    const cell = ws.getCell(r, 2);
    cell.value = {
      formula: `IFERROR((SUM(${PL_I})+SUMIF(${ROT_M},"*T*",${ROT_L}))/(COUNT(${PL_I})+COUNTIF(${ROT_M},"*T*")),"N/A")`,
    } as ExcelJS.CellFormulaValue;
    cell.numFmt = "0.00";
    styleData(cell, "right");
  }
  r += 3;

  // ---------------- Table 4: PL - Units Sold ----------------
  const plRoomsSet = uniqueSorted([
    ...saRows.filter((s) => s.isTender === false && s.noOfRooms !== null).map((s) => s.noOfRooms as number),
    ...rotRows.filter((rr) => rr.isTender === false && rr.noOfRms !== null).map((rr) => rr.noOfRms as number),
  ]);
  r = writeUnitsSoldTable(ws, r, "Table 4 · PL-Units Sold PL 去貨情況", plRoomsSet, SA_D, SA_B, ROT_K, ROT_M, ROT_A, "<>*T*");
  r += 2;

  // ---------------- Table 5: Tender - Units Sold ----------------
  const tenderRoomsSet = uniqueSorted([
    ...saRows.filter((s) => s.isTender === true && s.noOfRooms !== null).map((s) => s.noOfRooms as number),
    ...rotRows.filter((rr) => rr.isTender === true && rr.noOfRms !== null).map((rr) => rr.noOfRms as number),
  ]);
  r = writeUnitsSoldTable(ws, r, "Table 5 · Tender-Units Sold Tender 去貨情況", tenderRoomsSet, SA_D, SA_B, ROT_K, ROT_M, ROT_A, "*T*");
  r += 2;

  // ---------------- Table 6: PL - Price ----------------
  const plPriceRooms = uniqueSorted(rotRows.filter((rr) => rr.isTender === false && rr.noOfRms !== null).map((rr) => rr.noOfRms as number));
  r = writePriceTable(ws, r, "Table 6 · PL-Price PL 呎價分析", plPriceRooms, ROT_K, ROT_M, ROT_L, "<>*T*");
  r += 2;

  // ---------------- Table 7: Tender - Price ----------------
  const tenderPriceRooms = uniqueSorted(rotRows.filter((rr) => rr.isTender === true && rr.noOfRms !== null).map((rr) => rr.noOfRms as number));
  r = writePriceTable(ws, r, "Table 7 · Tender-Price Tender 呎價分析", tenderPriceRooms, ROT_K, ROT_M, ROT_L, "*T*");
  r += 2;

  // ---------------- Table 8: Overall - Price ----------------
  ws.mergeCells(r, 1, r, 2);
  styleTitle(ws.getCell(r, 1));
  ws.getCell(r, 1).value = "Table 8 · Overall-Price 整體呎價分析";
  r++;
  ws.getRow(r).values = ["Metric", "Price (per sq.ft.)"];
  [1, 2].forEach((c) => styleHeader(ws.getCell(r, c)));
  r++;
  const overallMetrics: Array<[string, string]> = [
    ["Min. Price (per sq. ft.)", `IFERROR(MIN(${ROT_L}),"N/A")`],
    ["Max. Price (per sq. ft.)", `IFERROR(MAX(${ROT_L}),"N/A")`],
    ["Avg. Price (per sq. ft.)", `IFERROR(AVERAGE(${ROT_L}),"N/A")`],
  ];
  for (const [label, formula] of overallMetrics) {
    ws.getCell(r, 1).value = label;
    styleData(ws.getCell(r, 1), "left");
    const cell = ws.getCell(r, 2);
    cell.value = { formula } as ExcelJS.CellFormulaValue;
    cell.numFmt = "0.00";
    styleData(cell, "right");
    r++;
  }

  ws.views = [{ state: "frozen", ySplit: 0 }];

  return { sheetName: "Summary" };
}

function writeUnitsSoldTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  title: string,
  rooms: number[],
  SA_D: string,
  SA_B: string,
  ROT_K: string,
  ROT_M: string,
  ROT_A: string,
  tenderCriteria: string,
): number {
  let r = startRow;
  ws.mergeCells(r, 1, r, 4);
  styleTitle(ws.getCell(r, 1));
  ws.getCell(r, 1).value = title;
  r++;
  const headers = ["No. of Rms", "Released", "Sold", "Take-up rate (%)"];
  headers.forEach((h, i) => {
    const cell = ws.getCell(r, i + 1);
    cell.value = h;
    styleHeader(cell);
  });
  r++;
  const firstDataRow = r;
  for (const room of rooms) {
    ws.getCell(r, 1).value = roomLabel(room);
    styleData(ws.getCell(r, 1));

    const releasedCell = ws.getCell(r, 2);
    releasedCell.value = {
      formula: `COUNTIFS(${SA_D},${room},${SA_B},"${tenderCriteria}")`,
    } as ExcelJS.CellFormulaValue;
    styleData(releasedCell, "right");

    const soldCell = ws.getCell(r, 3);
    soldCell.value = {
      formula: `COUNTIFS(${ROT_K},${room},${ROT_M},"${tenderCriteria}",${ROT_A},"<>")`,
    } as ExcelJS.CellFormulaValue;
    styleData(soldCell, "right");

    const rateCell = ws.getCell(r, 4);
    rateCell.value = { formula: `IFERROR(C${r}/B${r},0)` } as ExcelJS.CellFormulaValue;
    rateCell.numFmt = "0.0%";
    styleData(rateCell, "right");
    r++;
  }
  const lastDataRow = r - 1;
  // Total row
  ws.getCell(r, 1).value = "Total";
  styleTotal(ws.getCell(r, 1));
  if (rooms.length > 0) {
    const totalReleased = ws.getCell(r, 2);
    totalReleased.value = { formula: `SUM(B${firstDataRow}:B${lastDataRow})` } as ExcelJS.CellFormulaValue;
    styleTotal(totalReleased);
    const totalSold = ws.getCell(r, 3);
    totalSold.value = { formula: `SUM(C${firstDataRow}:C${lastDataRow})` } as ExcelJS.CellFormulaValue;
    styleTotal(totalSold);
    const totalRate = ws.getCell(r, 4);
    totalRate.value = { formula: `IFERROR(C${r}/B${r},0)` } as ExcelJS.CellFormulaValue;
    totalRate.numFmt = "0.0%";
    styleTotal(totalRate);
  } else {
    styleTotal(ws.getCell(r, 2));
    styleTotal(ws.getCell(r, 3));
    styleTotal(ws.getCell(r, 4));
  }
  r++;
  return r;
}

function writePriceTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  title: string,
  rooms: number[],
  ROT_K: string,
  ROT_M: string,
  ROT_L: string,
  tenderCriteria: string,
): number {
  let r = startRow;
  ws.mergeCells(r, 1, r, 4);
  styleTitle(ws.getCell(r, 1));
  ws.getCell(r, 1).value = title;
  r++;
  const headers = ["No. of Rms", "Min. Price (per sq.ft.)", "Max. Price (per sq.ft.)", "Avg. Price (per sq.ft.)"];
  headers.forEach((h, i) => {
    const cell = ws.getCell(r, i + 1);
    cell.value = h;
    styleHeader(cell);
  });
  r++;
  const firstDataRow = r;
  for (const room of rooms) {
    ws.getCell(r, 1).value = roomLabel(room);
    styleData(ws.getCell(r, 1));

    const minCell = ws.getCell(r, 2);
    minCell.value = { formula: `IFERROR(_xlfn.MINIFS(${ROT_L},${ROT_K},${room},${ROT_M},"${tenderCriteria}"),"N/A")` } as ExcelJS.CellFormulaValue;
    minCell.numFmt = "0.00";
    styleData(minCell, "right");

    const maxCell = ws.getCell(r, 3);
    maxCell.value = { formula: `IFERROR(_xlfn.MAXIFS(${ROT_L},${ROT_K},${room},${ROT_M},"${tenderCriteria}"),"N/A")` } as ExcelJS.CellFormulaValue;
    maxCell.numFmt = "0.00";
    styleData(maxCell, "right");

    const avgCell = ws.getCell(r, 4);
    avgCell.value = {
      formula: `IFERROR(_xlfn.AVERAGEIFS(${ROT_L},${ROT_K},${room},${ROT_M},"${tenderCriteria}"),"N/A")`,
    } as ExcelJS.CellFormulaValue;
    avgCell.numFmt = "0.00";
    styleData(avgCell, "right");
    r++;
  }
  const lastDataRow = r - 1;
  ws.getCell(r, 1).value = "Total";
  styleTotal(ws.getCell(r, 1));
  if (rooms.length > 0) {
    const minCell = ws.getCell(r, 2);
    minCell.value = { formula: `IFERROR(MIN(B${firstDataRow}:B${lastDataRow}),"N/A")` } as ExcelJS.CellFormulaValue;
    minCell.numFmt = "0.00";
    styleTotal(minCell);
    const maxCell = ws.getCell(r, 3);
    maxCell.value = { formula: `IFERROR(MAX(C${firstDataRow}:C${lastDataRow}),"N/A")` } as ExcelJS.CellFormulaValue;
    maxCell.numFmt = "0.00";
    styleTotal(maxCell);
    const avgCell = ws.getCell(r, 4);
    avgCell.value = { formula: `IFERROR(AVERAGEIFS(${ROT_L},${ROT_M},"${tenderCriteria}"),"N/A")` } as ExcelJS.CellFormulaValue;
    avgCell.numFmt = "0.00";
    styleTotal(avgCell);
  } else {
    styleTotal(ws.getCell(r, 2));
    styleTotal(ws.getCell(r, 3));
    styleTotal(ws.getCell(r, 4));
  }
  r++;
  return r;
}
