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

const HEADER_ROW_INDEX = 4;  // zero-based → Excel row 5 (column headers)
const DATA_ROW_OFFSET  = 5;  // zero-based → Excel row 6 (first data row)

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
 * Try each prefix against the given cell. Handles messy hotel data — the
 * regex is deliberately loose because suppliers ship all sorts of variants:
 *
 *   "Supplier: 997286 - Zam Zam Trading"            (strict, standard)
 *   "Nhà cung cấp: 143277 - Cong Ty Song Hanh"      (VN, standard)
 *   "Supplier 143277 - Song Hanh"                   (no colon)
 *   "Supplier: 143277 Song Hanh"                    (no dash)
 *   "Supplier: Song Hanh"                           (no number)
 *   "Supplier 997286"                               (no name)
 *   "Supplier:"                                     (prefix only, empty)
 *
 * Returns { num, name } when any of the known prefixes matches, or null when
 * the cell doesn't start with a known prefix. Never throws — a missing or
 * malformed label row just means the banner won't display, and parsing continues.
 */
function parseLabelRow(cellValue, prefixes) {
  if (!cellValue) return null;
  const raw = String(cellValue).normalize("NFC").trim();
  if (!raw) return null;

  for (const prefix of prefixes) {
    // Match the prefix followed by either ":" or a digit. This accepts the
    // standard "Supplier: 123 - Name" AND messier "Supplier 143277 - Name"
    // (missing colon), while rejecting plain titles like "Supplier Items List
    // Report" that just happen to start with the same word.
    const prefixRe = new RegExp(`^\\s*${escapeRegex(prefix)}\\s*(?::\\s*|(?=\\d))(.*)$`, "i");
    const m = raw.match(prefixRe);
    if (!m) continue;

    const rest = m[1].trim();
    if (!rest) return { num: "", name: "" };  // prefix + colon alone, nothing after

    // Prefer {num}{sep}{name} where sep is dash/colon/comma.
    let split = rest.match(/^(\S+?)\s*[-:,]\s*(.+)$/);
    if (split) return { num: split[1].trim(), name: split[2].trim() };

    // Fallback: "{digits} {text}" with only whitespace as separator.
    split = rest.match(/^(\d+)\s+(.+)$/);
    if (split) return { num: split[1].trim(), name: split[2].trim() };

    // No split available: all digits → num-only; otherwise name-only.
    if (/^\d+$/.test(rest)) return { num: rest, name: "" };
    return { num: "", name: rest };
  }
  return null;
}

/**
 * Scan the first few rows (typically rows 1-4) for a cell matching any of the
 * given label prefixes. Hotels occasionally shift rows around or add blank
 * lines, so we don't pin to a fixed row index.
 *
 * Only column A is scanned by default (that's where suppliers always put these
 * labels in practice). Falls back silently to null when nothing matches — the
 * caller just doesn't get a banner.
 */
function findLabelInRows(aoa, prefixes, maxRowsToScan = 4) {
  const limit = Math.min(maxRowsToScan, aoa.length);
  for (let i = 0; i < limit; i++) {
    const row = aoa[i];
    if (!row) continue;
    const hit = parseLabelRow(row[0], prefixes);
    if (hit) return hit;
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

  // Supplier and division lines are OPTIONAL. Scan the first 4 rows (not
  // fixed positions) to handle hotel files that shift things around or have
  // blank leading rows. If nothing matches, both stay null and the banner
  // just won't show — parsing continues normally.
  const division = findLabelInRows(aoa, DIVISION_LABEL_PREFIXES, 4);
  const supplier = findLabelInRows(aoa, SUPPLIER_LABEL_PREFIXES, 4);

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
