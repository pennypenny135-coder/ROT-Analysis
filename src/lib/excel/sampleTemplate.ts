// Generates a fictional sample workbook (PL / SA / ROT / Summary) that
// exercises every parsing rule described in the spec, and triggers a
// client-side download.
import ExcelJS from "exceljs";
import { readWorkbookData } from "./readWorkbook";
import { runLookup } from "./lookup";
import { applyAutoFillAndErrors, downloadWorkbookBuffer } from "./exportWorkbook";
import { buildSummarySheet } from "./summary";

interface SamplePLUnit {
  block: string;
  floor: string;
  flat: string;
  rms: number;
  plNo: number;
  sqft: number;
  pricePerSqFt: number;
  maxDiscount: number | string;
}

const HEADER_FONT: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FFFFFFFF" } };
const HEADER_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };

function styleHeaderRow(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.font = HEADER_FONT;
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });
}

export async function buildSampleWorkbook(): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Record of Transaction Summary App";
  workbook.created = new Date();

  const units: SamplePLUnit[] = [
    { block: "1", floor: "5", flat: "A", rms: 0, plNo: 1, sqft: 320, pricePerSqFt: 18500, maxDiscount: 0.05 },
    { block: "1", floor: "5", flat: "B", rms: 1, plNo: 1, sqft: 492, pricePerSqFt: 17800, maxDiscount: 0.05 },
    { block: "1", floor: "6", flat: "A", rms: 1, plNo: 1, sqft: 505, pricePerSqFt: 17950, maxDiscount: "5%" },
    { block: "1", floor: "6", flat: "B", rms: 2, plNo: 1, sqft: 650, pricePerSqFt: 17200, maxDiscount: 0.05 },
    { block: "2", floor: "7", flat: "A", rms: 2, plNo: 2, sqft: 668, pricePerSqFt: 17600, maxDiscount: 0.08 },
    { block: "2", floor: "7", flat: "B", rms: 2, plNo: 2, sqft: 672, pricePerSqFt: 17550, maxDiscount: 0.08 },
    { block: "2", floor: "8", flat: "A", rms: 3, plNo: 2, sqft: 820, pricePerSqFt: 16900, maxDiscount: "8%" },
    { block: "2", floor: "9", flat: "B", rms: 3, plNo: 2, sqft: 835, pricePerSqFt: 16850, maxDiscount: 0.08 },
    { block: "3", floor: "10", flat: "A", rms: 3, plNo: 3, sqft: 845, pricePerSqFt: 16700, maxDiscount: 0.1 },
    { block: "3", floor: "10", flat: "B", rms: 2, plNo: 3, sqft: 680, pricePerSqFt: 17400, maxDiscount: 0.1 },
    { block: "3", floor: "11", flat: "A", rms: 1, plNo: 3, sqft: 498, pricePerSqFt: 17850, maxDiscount: "10%" },
    { block: "3", floor: "11", flat: "B", rms: 0, plNo: 3, sqft: 318, pricePerSqFt: 18600, maxDiscount: 0.1 },
  ];

  // ---------------- PL sheet ----------------
  const plWs = workbook.addWorksheet("PL");
  plWs.columns = [
    { header: "Block", key: "block", width: 10 },
    { header: "Floor", key: "floor", width: 10 },
    { header: "Flat", key: "flat", width: 10 },
    { header: "BFF", key: "bff", width: 14 },
    { header: "實用面積 Saleable Area", key: "area", width: 34 },
    { header: "Price ($)", key: "price", width: 14 },
    { header: "呎價 Price per sq. ft.", key: "psf", width: 16 },
    { header: "Max Discount (%)", key: "maxDiscount", width: 16 },
    { header: "Discounted 呎價 Discounted Price per sq. ft.", key: "discPsf", width: 20 },
    { header: "Rms", key: "rms", width: 8 },
    { header: "PL No.", key: "plNo", width: 10 },
  ];
  styleHeaderRow(plWs.getRow(1));

  units.forEach((u, idx) => {
    const r = idx + 2;
    const bff = `${u.block}${u.flat}${u.floor}`;
    const price = Math.round(u.pricePerSqFt * u.sqft);
    const discount = typeof u.maxDiscount === "string" ? parseFloat(u.maxDiscount) / 100 : u.maxDiscount;
    const discountedPsf = Math.round(u.pricePerSqFt * (1 - discount));
    const sqm = Math.round(u.sqft * 0.0929 * 1000) / 1000;
    const balconySqm = 2.02;
    const balconySqft = 22;

    const row = plWs.getRow(r);
    row.getCell(1).value = u.block;
    row.getCell(2).value = u.floor;
    row.getCell(3).value = u.flat;
    row.getCell(4).value = { formula: "A" + r + "&C" + r + "&B" + r, result: bff } as ExcelJS.CellFormulaValue;
    row.getCell(5).value = `${sqm.toFixed(3)} ( ${u.sqft} )\n露台 Balcony: ${balconySqm.toFixed(3)} ( ${balconySqft} )\n工作平台 Utility Platform: - (-)`;
    row.getCell(5).alignment = { wrapText: true, vertical: "top" };
    row.getCell(6).value = price;
    row.getCell(6).numFmt = "#,##0";
    row.getCell(7).value = u.pricePerSqFt;
    row.getCell(7).numFmt = "#,##0";
    row.getCell(8).value = u.maxDiscount;
    if (typeof u.maxDiscount === "number") row.getCell(8).numFmt = "0%";
    row.getCell(9).value = discountedPsf;
    row.getCell(9).numFmt = "#,##0";
    row.getCell(10).value = u.rms;
    row.getCell(11).value = u.plNo;
    row.commit();
  });

  // ---------------- SA sheet ----------------
  const saWs = workbook.addWorksheet("SA");
  saWs.columns = [
    { header: "BFF", key: "bff", width: 14 },
    { header: "SA", key: "sa", width: 10 },
    { header: "實用面積 (sq. ft.) Saleable Area", key: "area", width: 24 },
    { header: "No. of Rooms", key: "rooms", width: 14 },
  ];
  styleHeaderRow(saWs.getRow(1));

  let saRowIdx = 2;
  units.forEach((u) => {
    const bff = `${u.block}${u.flat}${u.floor}`;
    const row = saWs.getRow(saRowIdx++);
    row.getCell(1).value = bff;
    row.getCell(2).value = u.plNo; // 1 / 2 / 3 = PL
    // Area / Rooms intentionally left blank - the app fills these from PL.
    row.commit();
  });
  // Tender units (not present in PL) - user supplies area & rooms manually.
  const tenderUnits = [
    { bff: "4A12", saNo: "1T", area: 560, rooms: 2 },
    { bff: "4B12", saNo: "2T", area: 905, rooms: 3 },
  ];
  tenderUnits.forEach((t) => {
    const row = saWs.getRow(saRowIdx++);
    row.getCell(1).value = t.bff;
    row.getCell(2).value = t.saNo;
    row.getCell(3).value = t.area;
    row.getCell(4).value = t.rooms;
    row.commit();
  });

  // ---------------- ROT sheet ----------------
  const rotWs = workbook.addWorksheet("ROT");
  rotWs.columns = [
    { header: "Date of PASP", key: "pasp", width: 16 },
    { header: "Date of ASP", key: "asp", width: 16 },
    { header: "Date of termination of ASP", key: "term", width: 20 },
    { header: "Block", key: "block", width: 10 },
    { header: "Floor", key: "floor", width: 10 },
    { header: "Flat", key: "flat", width: 10 },
    { header: "BFF", key: "bff", width: 14 },
    { header: "Transaction Price", key: "price", width: 16 },
    { header: "Terms of Payment", key: "terms", width: 28 },
    { header: "", key: "empty", width: 4 },
    { header: "No. of Rms", key: "rms", width: 12 },
    { header: "ROT 呎價", key: "psf", width: 14 },
  ];
  styleHeaderRow(rotWs.getRow(1));

  interface SoldRecord {
    unit: SamplePLUnit | { bff: string; discountedPricePerSqFt: number };
    sqft: number;
    pasp: Date | string;
    asp: Date | string | null;
    terms: string;
  }

  const soldPLUnits = units.slice(0, 6); // first 6 PL units sold
  const sold: SoldRecord[] = soldPLUnits.map((u, idx) => {
    const usesTextDate = idx % 2 === 1;
    const day = 5 + idx;
    const paspDay = day > 31 ? day - 31 : day;
    const aspDay = day + 14 > 31 ? day + 14 - 31 : day + 14;
    return {
      unit: u,
      sqft: u.sqft,
      pasp: usesTextDate ? `${String(paspDay).padStart(2, "0")}-03-2024` : new Date(2024, 2, day),
      asp: usesTextDate ? `${String(aspDay).padStart(2, "0")}-03-2024` : new Date(2024, 2, day + 14),
      terms: idx % 2 === 0 ? "120-day cash payment plan" : "360-day standard payment plan",
    };
  });

  let rotRowIdx = 2;
  sold.forEach((s) => {
    const u = s.unit as SamplePLUnit;
    const bff = `${u.block}${u.flat}${u.floor}`;
    const discount = typeof u.maxDiscount === "string" ? parseFloat(u.maxDiscount) / 100 : u.maxDiscount;
    const discountedPsf = Math.round(u.pricePerSqFt * (1 - discount));
    const txPrice = Math.round(discountedPsf * u.sqft);
    const r = rotRowIdx++;
    const row = rotWs.getRow(r);
    row.getCell(1).value = s.pasp;
    if (typeof s.pasp === "string") row.getCell(1).numFmt = "@";
    row.getCell(2).value = s.asp ?? "";
    row.getCell(3).value = "";
    row.getCell(4).value = u.block;
    row.getCell(5).value = u.floor;
    row.getCell(6).value = u.flat;
    row.getCell(7).value = { formula: "D" + r + "&F" + r + "&E" + r, result: bff } as ExcelJS.CellFormulaValue;
    row.getCell(8).value = txPrice;
    row.getCell(8).numFmt = "#,##0";
    row.getCell(9).value = s.terms;
    row.commit();
  });

  // Sold tenders
  const tenderSold = [
    { bff: "4A12", pasp: new Date(2024, 3, 2), asp: new Date(2024, 3, 16), price: 560 * 17000, terms: "150-day cash payment plan" },
    { bff: "4B12", pasp: "10-04-2024", asp: "24-04-2024", price: 905 * 16500, terms: "270-day standard payment plan" },
  ];
  tenderSold.forEach((t) => {
    const r = rotRowIdx++;
    const row = rotWs.getRow(r);
    row.getCell(1).value = t.pasp;
    row.getCell(2).value = t.asp;
    row.getCell(3).value = "";
    // Tender units are not part of PL's Block/Floor/Flat grid; store BFF directly.
    row.getCell(4).value = "";
    row.getCell(5).value = "";
    row.getCell(6).value = "";
    row.getCell(7).value = t.bff;
    row.getCell(8).value = Math.round(t.price);
    row.getCell(8).numFmt = "#,##0";
    row.getCell(9).value = t.terms;
    row.commit();
  });

  // One terminated (and re-sold) example to show the date-of-termination column in use.
  {
    const u = units[6];
    const bff = `${u.block}${u.flat}${u.floor}`;
    const r = rotRowIdx++;
    const row = rotWs.getRow(r);
    row.getCell(1).value = "02-01-2024";
    row.getCell(2).value = "16-01-2024";
    row.getCell(3).value = "20-02-2024";
    row.getCell(4).value = u.block;
    row.getCell(5).value = u.floor;
    row.getCell(6).value = u.flat;
    row.getCell(7).value = { formula: "D" + r + "&F" + r + "&E" + r, result: bff } as ExcelJS.CellFormulaValue;
    row.getCell(8).value = Math.round(u.pricePerSqFt * u.sqft * 0.95);
    row.getCell(8).numFmt = "#,##0";
    row.getCell(9).value = "Terminated - re-offered to market";
    row.commit();
  }

  // ---------------- Build Summary from the data we just created ----------------
  const parsed = await readWorkbookData(workbook);
  const lookedUp = runLookup(parsed.plRows, parsed.saRows, parsed.rotRows);
  applyAutoFillAndErrors(workbook, lookedUp.plRows, lookedUp.saRows, lookedUp.rotRows, lookedUp.issues);
  buildSummarySheet(workbook, lookedUp.plRows, lookedUp.saRows, lookedUp.rotRows);

  return workbook;
}

export async function downloadSampleTemplate(): Promise<void> {
  const workbook = await buildSampleWorkbook();
  const buffer = await workbook.xlsx.writeBuffer();
  downloadWorkbookBuffer(buffer, "Record_of_Transaction_Sample_Template.xlsx");
}
