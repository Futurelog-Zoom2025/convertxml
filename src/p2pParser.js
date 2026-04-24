// Parses a P2P Excel file (FutureLog "Supplier Items List Report" format).
//
// File layout (from real P2P samples):
//   Row 1:  "Supplier Items List Report" (title)
//   Row 2:  "Division: {num} - {name}"       (EN)
//           "Khách sạn: {num} - {name}"      (VN)  — literally "Hotel"
//   Row 3:  "Supplier: {num} - {name}"       (EN)
//           "Nhà cung cấp: {num} - {name}"   (VN)
//   Row 4:  blank
//   Row 5:  column headers (EN or VN)
//   Row 6+: data rows, terminated by a blank Article No.
//
// Header matching is language-agnostic — we try all known aliases per field
// and take the first hit. NFC normalization handles VN files where diacritics
// are stored decomposed (e.g. "ã" as "a + combining tilde").

import {
  P2P_ALIASES, P2P_FIELD_KEYS,
  SUPPLIER_LABEL_PREFIXES, DIVISION_LABEL_PREFIXES,
} from "./p2pHeaders.js";

const HEADER_ROW_INDEX = 4;  // zero-based → Excel row 5
const DATA_ROW_OFFSET  = 6;  // zero-based data start → Excel row 6

function normalizeHeader(h) {
  if (h === null || h === undefined) return "";
  // NFC normalization collapses decomposed diacritics (e.g. Vietnamese files
  // often store "ã" as "a + combining tilde") into their composed form, so
  // file-row bytes can match our aliases typed in composed form.
  return String(h).normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
}

// Escape regex metachars so label prefixes can include "(" etc. in the future.
function escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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
 * Try each prefix against the given cell. For "Supplier: 997286 - Zam Zam..."
 * or "Nhà cung cấp: 143277 - Song Hanh..." returns { num, name } on match,
 * or null if no prefix matched.
 */
function parseLabelRow(cellValue, prefixes) {
  if (!cellValue) return null;
  const s = String(cellValue).normalize("NFC").trim();
  for (const prefix of prefixes) {
    const re = new RegExp(
      `^\\s*${escapeRegex(prefix)}\\s*:\\s*([^\\s-][^-]*?)\\s*-\\s*(.+?)\\s*$`,
      "i"
    );
    const m = s.match(re);
    if (m) return { num: m[1].trim(), name: m[2].trim() };
  }
  return null;
}

/**
 * Given row 5 (the header row), return an index map { fieldKey: colIndex | -1 }
 * plus a list of missing required fields for user-facing errors.
 */
function resolveHeaders(headerRow, useNewPriceCol) {
  const headerToIdx = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    const norm = normalizeHeader(headerRow[i]);
    if (norm && !headerToIdx.has(norm)) headerToIdx.set(norm, i);
  }

  const resolved = {};
  const missing  = [];

  for (const key of P2P_FIELD_KEYS) {
    const aliases = P2P_ALIASES[key] || [];
    let idx = -1;
    for (const alias of aliases) {
      const n = normalizeHeader(alias);
      if (headerToIdx.has(n)) {
        idx = headerToIdx.get(n);
        break;
      }
    }
    resolved[key] = idx;

    const isRequired =
      key === "articleNo" ||
      (useNewPriceCol && key === "newPrice") ||
      (!useNewPriceCol && key === "priceOrderUnit");
    if (isRequired && idx === -1) {
      // Report the first-listed alias (the "canonical" EN label)
      missing.push(aliases[0] || key);
    }
  }

  return { resolved, missing };
}

function getCell(src, resolved, key) {
  const idx = resolved[key];
  if (idx === -1 || idx === undefined) return "";
  return cleanCell(src[idx]);
}

/**
 * Parse an array-of-arrays representing a single P2P sheet.
 *
 * @param {Array<Array<any>>} aoa
 * @param {{useNewPriceCol: boolean}} opts  `lang` is no longer required for
 *                                          parsing — the matcher auto-detects
 *                                          language. Caller may still pass it
 *                                          but it's ignored here.
 * @returns {{supplier, division, rows}}
 */
export function parseP2PFile(aoa, opts) {
  const useNewPriceCol = !!opts?.useNewPriceCol;

  if (!Array.isArray(aoa) || aoa.length <= HEADER_ROW_INDEX) {
    throw new Error("Sheet looks empty or has fewer than 5 rows (expected headers on row 5).");
  }

  const division = parseLabelRow(aoa[1]?.[0], DIVISION_LABEL_PREFIXES);
  const supplier = parseLabelRow(aoa[2]?.[0], SUPPLIER_LABEL_PREFIXES);

  const headerRow = aoa[HEADER_ROW_INDEX] || [];
  const { resolved, missing } = resolveHeaders(headerRow, useNewPriceCol);

  if (missing.length) {
    throw new Error(
      `This doesn't look like a P2P file. Missing required header${
        missing.length === 1 ? "" : "s"
      } on row 5: ${missing.map((m) => `"${m}"`).join(", ")}. ` +
      `Click "Show expected headers" to compare.`
    );
  }

  const rows = [];
  for (let r = DATA_ROW_OFFSET; r < aoa.length; r++) {
    const src = aoa[r] || [];
    const articleRaw = getCell(src, resolved, "articleNo");
    const articleNo  = String(articleRaw).trim();
    if (articleNo === "") break;   // blank Article No. = end of data

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
      _excelRow: r + 1,
    });
  }

  return { supplier, division, rows };
}

/**
 * Higher-level entry point: given an array of sheets (each as an aoa), try
 * each in order and return the first one that parses successfully. Useful
 * when a workbook has leftover / empty / unrelated sheets that shouldn't
 * trip up the user.
 *
 * @param {Array<{name:string, aoa:Array}>} sheets
 * @param {object} opts  forwarded to parseP2PFile
 * @returns {{supplier, division, rows, sheetName, totalSheets}}
 */
export function parseFirstParseableSheet(sheets, opts) {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new Error("The workbook has no sheets.");
  }
  const errors = [];
  for (const { name, aoa } of sheets) {
    try {
      const result = parseP2PFile(aoa, opts);
      if (result.rows.length === 0) {
        errors.push(`sheet "${name}": parsed OK but contained no data rows`);
        continue;
      }
      return { ...result, sheetName: name, totalSheets: sheets.length };
    } catch (e) {
      errors.push(`sheet "${name}": ${e.message}`);
    }
  }
  throw new Error(
    `Could not find a usable P2P sheet in this workbook.\n\nTried ${sheets.length} sheet${sheets.length === 1 ? "" : "s"}:\n  - ${errors.join("\n  - ")}`
  );
}
