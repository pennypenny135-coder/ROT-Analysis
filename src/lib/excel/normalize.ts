// Pure normalization helpers used across the parsing / lookup pipeline.

/**
 * Extracts the first "(...)" group's numeric content from a saleable area
 * text block, e.g. "45.664 ( 492 )\n露台 Balcony: ..." -> 492
 * Tolerates spaces inside the parentheses. Returns null if it cannot be
 * extracted or is not numeric.
 */
export function extractSqFtFromAreaText(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;
  const text = String(raw);
  const match = text.match(/\(([^)]*)\)/);
  if (!match) return null;
  const inner = match[1].replace(/,/g, "").trim();
  if (!/^-?\d+(\.\d+)?$/.test(inner)) return null;
  const num = parseFloat(inner);
  return Number.isFinite(num) ? num : null;
}

/**
 * Normalizes a Max Discount value which can be a decimal fraction (0.05)
 * or a percentage string ("5%"), or a raw percentage number (5 meaning 5%).
 * Always returns a decimal fraction, e.g. 0.05.
 */
export function normalizeMaxDiscount(raw: unknown): number | null {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") {
    if (!Number.isFinite(raw)) return null;
    return Math.abs(raw) > 1 ? raw / 100 : raw;
  }
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (trimmed === "") return null;
    const isPct = trimmed.endsWith("%");
    const numPart = isPct ? trimmed.slice(0, -1) : trimmed;
    const num = parseFloat(numPart.replace(/,/g, ""));
    if (!Number.isFinite(num)) return null;
    if (isPct) return num / 100;
    return Math.abs(num) > 1 ? num / 100 : num;
  }
  return null;
}

export interface NormalizedSANo {
  normalized: string | null;
  isTender: boolean | null;
}

/**
 * Normalizes an SA "SA No." value: trims whitespace, removes internal
 * spaces, upper-cases the trailing "t" (Tender marker), e.g.
 * " 1t " / "1 T" / "1T" -> "1T". Plain numbers like "1", "2", "3" are PL.
 */
export function normalizeSANo(raw: unknown): NormalizedSANo {
  if (raw === null || raw === undefined) return { normalized: null, isTender: null };
  let s = String(raw).trim();
  if (s === "") return { normalized: null, isTender: null };
  s = s.replace(/\s+/g, "").toUpperCase();
  const match = s.match(/^(\d+)(T?)$/);
  if (!match) {
    // Cannot confidently normalize - still report best guess for tender detection.
    return { normalized: s, isTender: s.includes("T") };
  }
  return { normalized: s, isTender: match[2] === "T" };
}

/**
 * Parses a flexible date value which may already be a JS Date (from Excel
 * date-formatted cells) or a "DD-MM-YYYY" / "DD/MM/YYYY" text string.
 */
export function parseFlexibleDate(raw: unknown): Date | null {
  if (raw === null || raw === undefined || raw === "") return null;
  if (raw instanceof Date) {
    return Number.isNaN(raw.getTime()) ? null : raw;
  }
  if (typeof raw === "number") {
    // Excel serial date fallback (days since 1899-12-30)
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const ms = raw * 24 * 60 * 60 * 1000;
    const d = new Date(epoch.getTime() + ms);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (trimmed === "") return null;
    const m = trimmed.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10);
      const year = parseInt(m[3], 10);
      const d = new Date(year, month - 1, day);
      if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) {
        return d;
      }
      return null;
    }
    const iso = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (iso) {
      const year = parseInt(iso[1], 10);
      const month = parseInt(iso[2], 10);
      const day = parseInt(iso[3], 10);
      const d = new Date(year, month - 1, day);
      return Number.isNaN(d.getTime()) ? null : d;
    }
    const parsed = Date.parse(trimmed);
    return Number.isNaN(parsed) ? null : new Date(parsed);
  }
  return null;
}

/** Converts arbitrary cell content (number, numeric string, formula result) to a number. */
export function toNumber(raw: unknown): number | null {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  if (typeof raw === "string") {
    const cleaned = raw.replace(/[,$\s]/g, "");
    if (cleaned === "") return null;
    const num = parseFloat(cleaned);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

export function toTrimmedString(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  return String(raw).trim();
}

export function formatDateDDMMYYYY(date: Date | null): string {
  if (!date) return "";
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = date.getFullYear();
  return `${d}-${m}-${y}`;
}
