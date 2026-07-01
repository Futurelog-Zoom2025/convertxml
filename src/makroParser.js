// Parses a Makro (CPAxtra) price-list Excel file and runs the VBA VAT
// calculation on each row.
//
// File layout (new agreed format):
//   Row 1:  column headers (Thai) — see makroHeaders.js
//   Row 2+: data rows, terminated by a blank product code (รหัสสินค้า)
//
// Columns are located by HEADER NAME (row 1), not by position, so extra columns
// and column reordering are handled safely.
//
// ─────────────────────────────────────────────────────────────────────────────
// VAT CALCULATION — ported verbatim from the VBA (Module1
// Get_Data_From_FileCPAxtra…). DO NOT change this arithmetic; the business
// relies on it matching the workbook exactly. Only columns F (VAT amount) and
// G (price incl. VAT) feed the calc; the Ex-VAT column is display-only.
//
//   VAT%  (H) = IIf(VAT amount = 0, 0%, 7%)
//   I         = Round( InVAT / (1 + VAT%), 2 )        ' Price Exclude VAT
//   J         = Round( VAT amount + I, 2 )            ' Price Include VAT
//   K         = J - InVAT                             ' Diff (Decimal)   [not rounded]
//   L         = I - K                                 ' Price Exclude VAT(Adj)  ← new price
//   M         = L + VAT amount                        ' Price Include VAT(Adj)  [not rounded]
//   N         = M - InVAT                             ' Check Diff              [not rounded]
//
// The value carried forward as the row's "new price per order unit" is
// Round(L, 2) — matching the workbook's =ROUND(XLOOKUP(...L...),2) step.
// ─────────────────────────────────────────────────────────────────────────────

import { MAKRO_ALIASES, MAKRO_FIELD_KEYS, MAKRO_REQUIRED } from "./makroHeaders.js";

const HEADER_ROW_INDEX = 0;  // zero-based → Excel row 1 (column headers)
const DATA_ROW_OFFSET  = 1;  // zero-based → Excel row 2 (first data row)

const VAT_RATE = 0.07;       // 7% — the only non-zero VAT rate the VBA uses

// Normalize a header: NFC (collapse decomposed Thai/Latin diacritics), strip
// ALL whitespace, lowercase. Matching is by exact equality after normalization.
function normalizeHeader(h) {
  if (h === null || h === undefined) return "";
  return String(h).normalize("NFC").replace(/\s+/g, "").toLowerCase();
}

function cleanText(v) {
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

// Parse a numeric-ish cell (the Makro file stores prices as text like "81.12").
// Empty / non-numeric → 0, matching VBA's numeric coercion of blank cells.
function toNumber(v) {
  if (v === "" || v === null || v === undefined) return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

// VBA Round() — banker's rounding (round half to even). WorksheetFunction.Round
// differs (half away from zero), but the VBA here uses the built-in Round(),
// so we replicate half-to-even to stay byte-for-byte faithful.
export function vbaRound(value, digits = 0) {
  if (typeof value !== "number" || !Number.isFinite(value)) return value;
  const m = Math.pow(10, digits);
  const n = value * m;
  const sign = n < 0 ? -1 : 1;
  const abs = Math.abs(n);
  const floor = Math.floor(abs);
  const frac = abs - floor;
  let r;
  if (Math.abs(frac - 0.5) < 1e-9) {
    r = floor % 2 === 0 ? floor : floor + 1;   // half → nearest even
  } else {
    r = Math.round(abs);
  }
  return (sign * r) / m;
}

// Round to 2 decimals and normalize -0 → +0. Used to tidy the tiny
// floating-point residues in the diagnostic diff columns (Diff (Decimal),
// Check Diff) so the UI/export show a consistent "0" instead of a mix of
// 0 / 0.00 / -0.00. These columns are mathematically ~0 anyway, and the
// exported price is unaffected (it comes from the raw values).
export function clean2(v) {
  if (typeof v !== "number" || !Number.isFinite(v)) return v;
  const r = Math.round(v * 100) / 100;
  return r === 0 ? 0 : r;   // r === 0 also catches -0, so we return a clean +0
}

/**
 * Run the VBA VAT calculation for a single row.
 * @param {number} vatAmt   VAT amount per unit (column F)
 * @param {number} inVat    price including VAT (column G)
 * @returns computed fields incl. `priceExVatAdj` (L) and `newPrice` (Round(L,2)).
 */
export function computeVat(vatAmt, inVat) {
  const vatPct = vatAmt === 0 ? 0 : VAT_RATE;   // H
  const priceExVat = vbaRound(inVat / (1 + vatPct), 2);  // I
  const priceInVat = vbaRound(vatAmt + priceExVat, 2);   // J
  const diffDecimal = priceInVat - inVat;                // K (raw — feeds L below)
  const priceExVatAdj = priceExVat - diffDecimal;        // L (raw) ← new price basis
  const priceInVatAdj = priceExVatAdj + vatAmt;          // M (raw)
  const checkDiff = priceInVatAdj - inVat;               // N (raw)
  return {
    vatPct,
    priceExVat,
    priceInVat,
    // L and M keep their raw values (L feeds the exported price). K and N are
    // diagnostic-only and mathematically ~0, so we tidy their floating-point
    // dust to a clean 0 for display/export.
    diffDecimal: clean2(diffDecimal),
    priceExVatAdj,
    priceInVatAdj,
    checkDiff: clean2(checkDiff),
    // The value looked up into Report 1145 = ROUND(L, 2).
    newPrice: vbaRound(priceExVatAdj, 2),
  };
}

// Build a { fieldKey: colIndex | -1 } map from the header row, plus the list of
// missing required fields for user-facing errors.
function resolveHeaders(headerRow) {
  const headerToIdx = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    const norm = normalizeHeader(headerRow[i]);
    if (norm && !headerToIdx.has(norm)) headerToIdx.set(norm, i);
  }

  const resolved = {};
  const missing = [];
  for (const key of MAKRO_FIELD_KEYS) {
    const aliases = MAKRO_ALIASES[key] || [];
    let idx = -1;
    for (const alias of aliases) {
      const n = normalizeHeader(alias);
      if (headerToIdx.has(n)) { idx = headerToIdx.get(n); break; }
    }
    resolved[key] = idx;
    if (MAKRO_REQUIRED.has(key) && idx === -1) missing.push(aliases[0] || key);
  }
  return { resolved, missing };
}

function getText(src, resolved, key) {
  const idx = resolved[key];
  if (idx === -1 || idx === undefined) return "";
  return cleanText(src[idx]);
}
function getNum(src, resolved, key) {
  const idx = resolved[key];
  if (idx === -1 || idx === undefined) return 0;
  return toNumber(src[idx]);
}

/**
 * Parse an array-of-arrays representing a single Makro sheet.
 * @param {Array<Array<any>>} aoa
 * @returns {{rows: Array}}  rows keyed by artCode with VAT calc applied.
 */
export function parseMakroFile(aoa) {
  if (!Array.isArray(aoa) || aoa.length <= HEADER_ROW_INDEX) {
    throw new Error("The Makro sheet looks empty (expected headers on row 1).");
  }

  const headerRow = aoa[HEADER_ROW_INDEX] || [];
  const { resolved, missing } = resolveHeaders(headerRow);

  if (missing.length) {
    throw new Error(
      `This doesn't look like a Makro file. Missing required header${
        missing.length === 1 ? "" : "s"
      } on row 1: ${missing.map((m) => `"${m}"`).join(", ")}. ` +
      `Click "Show expected headers" to compare.`
    );
  }

  const rows = [];
  for (let r = DATA_ROW_OFFSET; r < aoa.length; r++) {
    const src = aoa[r] || [];
    const artCode = String(getText(src, resolved, "artCode")).trim();
    if (artCode === "") break;   // blank product code = end of data

    const vatAmt = getNum(src, resolved, "vat");
    const inVat = getNum(src, resolved, "priceInVat");
    const calc = computeVat(vatAmt, inVat);

    rows.push({
      artCode,
      itemName:   getText(src, resolved, "itemName"),
      artGroup:   getText(src, resolved, "artGroup"),
      status:     getText(src, resolved, "status"),
      // Raw "ราคาขาย (Ex. VAT)" from the file (display only). Named `srcExVat`
      // so the `...calc` spread's computed `priceExVat` (I) doesn't clobber it —
      // the two are usually equal but semantically distinct columns.
      srcExVat:   getNum(src, resolved, "priceExVat"),
      vatAmt,
      inVat,
      ...calc,
      _excelRow: r + 1,
    });
  }

  return { rows };
}

/**
 * Try each sheet in order and return the first that parses to at least one row.
 * Mirrors parseFirstParseableSheet in p2pParser.js so leftover/empty sheets in
 * the workbook don't trip up the user.
 */
export function parseFirstMakroSheet(sheets) {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new Error("The workbook has no sheets.");
  }
  const errors = [];
  for (const { name, aoa } of sheets) {
    try {
      const result = parseMakroFile(aoa);
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
    `Could not find a usable Makro sheet in this workbook.\n\nTried ${sheets.length} sheet${
      sheets.length === 1 ? "" : "s"
    }:\n  - ${errors.join("\n  - ")}`
  );
}
