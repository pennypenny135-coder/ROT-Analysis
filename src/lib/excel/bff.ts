// BFF (Block & Flat & Floor) helpers.
import { toTrimmedString } from "./normalize";

/** Builds the BFF string from Block, Flat, Floor in that concatenation order. */
export function buildBFF(block: unknown, flat: unknown, floor: unknown): string {
  return `${toTrimmedString(block)}${toTrimmedString(flat)}${toTrimmedString(floor)}`;
}

/** Normalizes a BFF value into a case/space-insensitive join key. */
export function normalizeBFFKey(bff: unknown): string {
  if (bff === null || bff === undefined) return "";
  return String(bff).trim().toUpperCase().replace(/\s+/g, "");
}
