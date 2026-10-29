// Cross-sheet BFF lookup, auto-fill and validation logic. Pure, unit-testable.
import { normalizeBFFKey } from "./bff";
import type { ParsedIssue, PLRow, ROTRow, SARow } from "./types";

export interface LookupResult {
  plRows: PLRow[];
  saRows: SARow[];
  rotRows: ROTRow[];
  issues: ParsedIssue[];
}

/** Builds a BFF -> PLRow index (first occurrence wins). */
export function buildPLIndex(plRows: PLRow[]): Map<string, PLRow> {
  const map = new Map<string, PLRow>();
  for (const row of plRows) {
    const key = normalizeBFFKey(row.bff);
    if (!key) continue;
    if (!map.has(key)) map.set(key, row);
  }
  return map;
}

/** Builds a BFF -> SARow index (first occurrence wins). */
export function buildSAIndex(saRows: SARow[]): Map<string, SARow> {
  const map = new Map<string, SARow>();
  for (const row of saRows) {
    const key = normalizeBFFKey(row.bff);
    if (!key) continue;
    if (!map.has(key)) map.set(key, row);
  }
  return map;
}

/**
 * Runs the full cross-sheet lookup pipeline:
 *  - Fills non-tender SA Saleable Area / No. of Rooms from PL.
 *  - Fills ROT No. of Rms from SA, and computes ROT 呎價.
 *  - Collects validation issues (missing lookups, invalid data).
 * Does not mutate the input arrays; returns new row objects.
 */
export function runLookup(plRowsIn: PLRow[], saRowsIn: SARow[], rotRowsIn: ROTRow[]): LookupResult {
  const issues: ParsedIssue[] = [];
  const plRows = plRowsIn.map((r) => ({ ...r }));
  const saRows = saRowsIn.map((r) => ({ ...r }));
  const rotRows = rotRowsIn.map((r) => ({ ...r }));

  const plIndex = buildPLIndex(plRows);

  for (const sa of saRows) {
    if (sa.isTender === true) {
      // Tender units: user-provided area/rooms, never overwritten.
      if (sa.saleableAreaSqFt === null || sa.noOfRooms === null) {
        issues.push({
          sheet: "SA",
          row: sa.rowNumber,
          field: "Saleable Area / No. of Rooms (Column C/D)",
          bff: sa.bff,
          message: `Tender BFF "${sa.bff}" is missing Saleable Area or No. of Rooms (must be entered manually).`,
          severity: "warning",
        });
      }
      continue;
    }
    if (sa.isTender === false) {
      const key = normalizeBFFKey(sa.bff);
      const pl = key ? plIndex.get(key) : undefined;
      if (!pl) {
        sa.plMatchFound = false;
        issues.push({
          sheet: "SA",
          row: sa.rowNumber,
          field: "BFF (Column A)",
          bff: sa.bff,
          message: `Non-tender BFF "${sa.bff}" not found in PL.`,
          severity: "error",
        });
        continue;
      }
      sa.plMatchFound = true;
      sa.saleableAreaSqFt = pl.saleableAreaSqFt;
      sa.noOfRooms = pl.rms;
      sa.filledArea = true;
      sa.filledRooms = true;
    }
  }

  const saIndex = buildSAIndex(saRows);

  for (const rot of rotRows) {
    const key = normalizeBFFKey(rot.bff);
    const sa = key ? saIndex.get(key) : undefined;

    if (!sa) {
      rot.saMatchFound = false;
      issues.push({
        sheet: "ROT",
        row: rot.rowNumber,
        field: "BFF (Column G)",
        bff: rot.bff,
        message: `BFF "${rot.bff}" not found in SA.`,
        severity: "error",
      });
      rot.noOfRms = null;
      rot.rotPricePerSqFt = null;
      continue;
    }

    rot.saMatchFound = true;
    rot.isTender = sa.isTender;
    rot.noOfRms = sa.noOfRooms;

    if (rot.transactionPrice === null) {
      rot.rotPricePerSqFt = null;
      continue; // issue already raised during parsing
    }

    if (sa.saleableAreaSqFt === null || sa.saleableAreaSqFt <= 0) {
      rot.rotPricePerSqFt = null;
      issues.push({
        sheet: "ROT",
        row: rot.rowNumber,
        field: "ROT 呎價 (Column L)",
        bff: rot.bff,
        message: `Cannot compute ROT 呎價: SA BFF "${rot.bff}" has no valid saleable area.`,
        severity: "error",
      });
      continue;
    }

    rot.rotPricePerSqFt = rot.transactionPrice / sa.saleableAreaSqFt;
  }

  return { plRows, saRows, rotRows, issues };
}
