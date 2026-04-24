// Parses a P2P Excel file (FutureLog "Supplier Items List Report" format).
//
// File layout (from the reference Zam Zam P2P sample):
//   Row 1:  "Supplier Items List Report" (title)
//   Row 2:  "Division: {num} - {name}"
//   Row 3:  "Supplier: {num} - {name}"     ← we parse this for the supplier banner
//   Row 4:  blank
//   Row 5:  column headers (EN or VN)
//   Row 6+: data rows, terminated by a blank Article No.
//
// Header matching is tolerant: case-insensitive, whitespace-collapsed.
// Hotels occasionally have trailing spaces in headers (e.g. "Order Unit ").

import { P2P_HEADERS, P2P_FIELD_KEYS } from "./p2pHeaders.js";

const HEADER_ROW_INDEX = 4;  // zero-based → Excel row 5
const DATA_ROW_OFFSET  = 6;  // zero-based data start → Excel row 6

function normalizeHeader(h) {
  if (h === null || h === undefined) return "";
  return String(h).replace(/\s+/g, " ").trim().toLowerCase();
}

function cleanCell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return Number.isNaN(v) ? "" : v;
  if (v instanceof Date) {
    const d = String(v.getDate()).padStart(2, "0");
    const m = String(v.getMonth() + 1).padStart(2, "0");
    return `${d}.${m}.${v.getFullYear()}`;
  }
  return String(v);
}

function toNumericOrEmpty(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "") return "";
  const n = Number(s);
  return Number.isFinite(n) ? n : "";
}

/**
 * Extract "Supplier: 997286 - Zam Zam Trading Sdn Bhd" → {num: "997286", name: "Zam Zam Trading Sdn Bhd"}.
 * Returns null when the row doesn't match the expected shape — banner will be hidden.
 */
function parseLabelRow(cellValue, prefix) {
  if (!cellValue) return null;
  const s = String(cellValue).trim();
  const re = new RegExp(`^\\s*${prefix}\\s*:\\s*([^\\s-][^-]*?)\\s*-\\s*(.+?)\\s*$`, "i");
  const m = s.match(re);
  if (!m) return null;
  return { num: m[1].trim(), name: m[2].trim() };
}

/**
 * Given row 5 (the header row) and the chosen language, return an index map
 *   { fieldKey: colIndex | -1 }
 * plus a list of missing required fields for user-facing errors.
 */
function resolveHeaders(headerRow, lang, useNewPriceCol) {
  const canonical = P2P_HEADERS[lang] || P2P_HEADERS.EN;
  // Build normalized-header → colIndex lookup
  const headerToIdx = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    const norm = normalizeHeader(headerRow[i]);
    if (norm && !headerToIdx.has(norm)) headerToIdx.set(norm, i);
  }

  const resolved = {};
  const missing  = [];

  for (const key of P2P_FIELD_KEYS) {
    const label = canonical[key];
    let idx = headerToIdx.get(normalizeHeader(label)) ?? -1;
    // For newPrice specifically, fall back to the EN label if VN didn't find it.
    if (idx === -1 && key === "newPrice" && lang !== "EN") {
      idx = headerToIdx.get(normalizeHeader(P2P_HEADERS.EN.newPrice)) ?? -1;
    }
    resolved[key] = idx;

    const isRequired =
      key === "articleNo" ||
      (useNewPriceCol && key === "newPrice") ||
      (!useNewPriceCol && key === "priceOrderUnit");
    if (isRequired && idx === -1) missing.push(label);
  }

  return { resolved, missing };
}

function getCell(src, resolved, key) {
  const idx = resolved[key];
  if (idx === -1 || idx === undefined) return "";
  return cleanCell(src[idx]);
}

/**
 * @param {Array<Array<any>>} aoa  sheet parsed as array-of-arrays (SheetJS header:1)
 * @param {{lang: "EN"|"VN", useNewPriceCol: boolean}} opts
 * @returns {{
 *    supplier: {num,name}|null,
 *    division: {num,name}|null,
 *    rows: Array<{articleNo:string, itemName:string, priceOU:number|"", newPrice:number|"", ...}>
 * }}
 */
export function parseP2PFile(aoa, opts) {
  const lang = opts?.lang === "VN" ? "VN" : "EN";
  const useNewPriceCol = !!opts?.useNewPriceCol;

  if (!Array.isArray(aoa) || aoa.length <= HEADER_ROW_INDEX) {
    throw new Error("P2P file looks empty or malformed — expected headers on row 5.");
  }

  // Row 3 (index 2) carries the supplier label; row 2 (index 1) carries the division.
  // These are optional — silent fallback when missing.
  const division  = parseLabelRow(aoa[1]?.[0], "Division");
  const supplier  = parseLabelRow(aoa[2]?.[0], "Supplier");

  const headerRow = aoa[HEADER_ROW_INDEX] || [];
  const { resolved, missing } = resolveHeaders(headerRow, lang, useNewPriceCol);

  if (missing.length) {
    throw new Error(
      `This doesn't look like a ${lang} P2P file. Missing required header${
        missing.length === 1 ? "" : "s"
      } on row 5: ${missing.map((m) => `"${m}"`).join(", ")}. ` +
      `Try switching the language picker, or click "Show expected headers" to compare.`
    );
  }

  const rows = [];
  for (let r = DATA_ROW_OFFSET; r < aoa.length; r++) {
    const src = aoa[r] || [];
    // End-of-data signal: blank Article No.
    const articleRaw = getCell(src, resolved, "articleNo");
    const articleNo  = String(articleRaw).trim();
    if (articleNo === "") break;

    rows.push({
      articleNo,
      wsNo:          getCell(src, resolved, "wsNo"),
      itemName:      getCell(src, resolved, "itemName"),
      gtin:          getCell(src, resolved, "gtin"),
      orderUnit:     getCell(src, resolved, "orderUnit"),
      contentUnits:  getCell(src, resolved, "contentUnits"),
      packagingUnit: getCell(src, resolved, "packagingUnit"),
      priceOrderUnit: toNumericOrEmpty(src[resolved.priceOrderUnit]),
      newPrice:       toNumericOrEmpty(src[resolved.newPrice]),
      minOrderQty:   getCell(src, resolved, "minOrderQty"),
      originCountry: getCell(src, resolved, "originCountry"),
      // Excel row for error messages / debugging
      _excelRow: r + 1,
    });
  }

  return { supplier, division, rows };
}
