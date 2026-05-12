// Excel export for the Full Data modal.
//
// Used by both tabs (R1145 and P2P). Reads the same `rows` + `columns` +
// `invalidCells` structures that the modal already maintains, then writes an
// .xlsx file with:
//
//   - Bold gray header row, frozen so it stays visible while scrolling
//   - Red fill on any cell flagged as a blocking error (from invalidCells)
//   - Yellow fill on cells related to warning-status pills (parser
//     modifications like Scaled price = 0, No price in scaled, etc.) — these
//     are not validation errors but rows the user should review
//   - Status column at the right edge with multiple pills joined by ", "
//   - Numeric values stored as numbers (so users can sort/filter naturally
//     in Excel) — prices and lead times specifically
//   - Default column widths sized for typical Report 1145 content
//
// The export uses xlsx-js-style (a fork of SheetJS that preserves cell
// styling). Plain SheetJS only writes structure, not styles, so red/yellow
// fills would not appear if we used it directly.
//
// Triggered by the "Export to Excel" button in the modal header. The button
// receives the CURRENT filtered view (errors-only + status + search) so the
// exported file matches what the user sees on screen.

import XLSX from "xlsx-js-style";
import { NA_MARKER } from "./reportParser.js";

// Yellow-fill candidate keys: cells related to parser modifications. We
// highlight these only when the row has a relevant warning-status pill.
const PRICE_LEAD_KEYS = new Set([
  "priceOU", "availability", "leadTimeRaw",
]);

// Status pill labels that indicate "this row had a parser modification".
// When present in a row's `statuses` array (or single `status` string), the
// price + lead-time cells get yellow fill in the export.
const WARNING_STATUSES = new Set([
  "Scaled price = 0",
  "No price in scaled price",
  "Both prices blank",
  "Lead time update",
  // P2P-tab statuses that imply a non-trivial parser action
  "Open lead time",
  "Price from report 1145",
  "No price update",
  "P2P-only item",
]);

// Stringify a row's status field(s) for the export. Both shapes supported:
//   - r.statuses: string[]    (R1145 tab, multi-pill)
//   - r.status:   string      (P2P tab, single)
function statusString(r) {
  if (Array.isArray(r.statuses)) return r.statuses.filter(Boolean).join(", ");
  if (typeof r.status === "string") return r.status;
  return "";
}

// Returns true if this row has any warning-class status pill.
function rowHasWarningStatus(r) {
  if (Array.isArray(r.statuses)) return r.statuses.some((s) => WARNING_STATUSES.has(s));
  if (typeof r.status === "string") return WARNING_STATUSES.has(r.status);
  return false;
}

// Convert a cell value to the most appropriate Excel type. Numbers stay as
// numbers (so sorting works); everything else is a string. NA marker is
// rendered as "#N/A" for clarity.
function cellValue(v) {
  if (v === NA_MARKER) return "#N/A";
  if (v === null || v === undefined) return "";
  if (typeof v === "number" && Number.isFinite(v)) return v;
  return String(v);
}

// Style objects reused by every cell that takes a given style.
const STYLE_HEADER = {
  font: { bold: true, color: { rgb: "FF000000" } },
  fill: { patternType: "solid", fgColor: { rgb: "FFD9D9D9" } },
  alignment: { horizontal: "left", vertical: "center", wrapText: true },
  border: {
    top:    { style: "thin", color: { rgb: "FFB0B0B0" } },
    bottom: { style: "thin", color: { rgb: "FFB0B0B0" } },
    left:   { style: "thin", color: { rgb: "FFB0B0B0" } },
    right:  { style: "thin", color: { rgb: "FFB0B0B0" } },
  },
};
const FILL_ERROR_RED    = { patternType: "solid", fgColor: { rgb: "FFF8D7DA" } };
const FILL_WARNING_AMB  = { patternType: "solid", fgColor: { rgb: "FFFFF3CD" } };
const FONT_ERROR        = { color: { rgb: "FF96231C" } };
const FONT_WARNING      = { color: { rgb: "FF8A4F16" } };

/**
 * Build & download an .xlsx file containing the given rows.
 *
 * @param {object} args
 * @param {Array<object>} args.rows           Row objects to export (already filtered)
 * @param {Array<{key:string,label:string}>} args.columns  Column definitions (same shape as modal's `columns`)
 * @param {Map<number, Set<string>>} args.invalidCells  Cells flagged as blocking errors (keys are row indexes IN THE ORIGINAL DATASET, not the filtered subset — see indexBy)
 * @param {Array<object>} args.allRows        The full unfiltered dataset, used to resolve invalidCells row indexes back to the row objects
 * @param {string} args.filename              Suggested filename without extension
 * @param {string} args.sheetName             Name of the worksheet inside the file (max 31 chars per Excel)
 */
export function exportFullDataToExcel({ rows, columns, invalidCells, allRows, filename, sheetName }) {
  // Build a lookup so we can find each visible row's original-dataset index.
  // invalidCells keys are indexes into `allRows`, not the filtered `rows`.
  const allRowsToIdx = new Map();
  (allRows || rows).forEach((r, idx) => allRowsToIdx.set(r, idx));

  // Columns to include in the export. Strip any non-displayable column
  // definitions (like the inline preview's column-class extras). The Status
  // column is appended at the end — never duplicated even if `columns`
  // already contains it.
  const baseCols = columns
    .filter((c) => c.key !== "statuses" && c.key !== "status" && c.key !== "__status")
    .map((c) => ({ key: c.key, label: c.label }));
  const exportCols = [...baseCols, { key: "__status", label: "Status" }];

  // Build the AOA (array-of-arrays) with cell objects for styling.
  //
  // SheetJS interprets:
  //   { v: value, t: "n"|"s", s: styleObject }
  // The "t" field is the type indicator: "n" = number, "s" = string. We let
  // SheetJS infer most types and only set "s" explicitly when applying styles.
  const aoa = [];

  // Header row
  aoa.push(exportCols.map((c) => ({ v: c.label, t: "s", s: STYLE_HEADER })));

  // Data rows
  for (const r of rows) {
    const originalIdx = allRowsToIdx.get(r);
    const invalidSet = originalIdx !== undefined ? invalidCells.get(originalIdx) : null;
    const hasWarning = rowHasWarningStatus(r);

    const cells = exportCols.map((c) => {
      const value = c.key === "__status" ? statusString(r) : cellValue(r[c.key]);
      const isInvalid = invalidSet && invalidSet.has(c.key);
      const isWarning = !isInvalid && hasWarning && PRICE_LEAD_KEYS.has(c.key);

      const cell = { v: value };
      // Set explicit type for numbers so Excel sorts/filters correctly.
      if (typeof value === "number") cell.t = "n";
      else cell.t = "s";

      if (isInvalid) {
        cell.s = { fill: FILL_ERROR_RED, font: FONT_ERROR };
      } else if (isWarning) {
        cell.s = { fill: FILL_WARNING_AMB, font: FONT_WARNING };
      }
      return cell;
    });
    aoa.push(cells);
  }

  // Build worksheet
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Column widths tuned for typical Report 1145 column lengths. Numeric
  // columns are narrower; description columns wider. Width "wch" is in
  // approximate characters.
  ws["!cols"] = exportCols.map((c) => {
    const k = c.key;
    if (k === "pos")               return { wch: 5 };
    if (k === "itemNo")            return { wch: 14 };
    if (k === "ean")               return { wch: 16 };
    if (k === "manArtId")          return { wch: 18 };
    if (k.startsWith("desc"))      return { wch: 32 };
    if (k === "itemName")          return { wch: 32 };
    if (k === "ou" || k === "cu")  return { wch: 6 };
    if (k === "cuou")              return { wch: 8 };
    if (k === "priceOU" || k === "__newPrice" || k === "__oldPrice")
                                   return { wch: 12 };
    if (k === "__diff")            return { wch: 10 };
    if (k === "origin")            return { wch: 8 };
    if (k === "customsNo")         return { wch: 14 };
    if (k === "leadTimeRaw")       return { wch: 12 };
    if (k === "availability")      return { wch: 14 };
    if (k === "specUrl")           return { wch: 24 };
    if (k === "offerStart" || k === "offerEnd")  return { wch: 12 };
    if (k === "customerId")        return { wch: 10 };
    if (k === "__status")          return { wch: 32 };
    return { wch: 14 };
  });

  // Freeze the header row so it stays visible while scrolling.
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  ws["!views"] = [{ state: "frozen", ySplit: 1 }];

  // Build & write workbook
  const wb = XLSX.utils.book_new();
  const sheetSafe = String(sheetName || "Full Data").slice(0, 31);
  XLSX.utils.book_append_sheet(wb, ws, sheetSafe);

  const today = new Date().toISOString().slice(0, 10);
  const finalName = `${filename || "Export"}_${today}.xlsx`;

  // SheetJS write → trigger browser download
  XLSX.writeFile(wb, finalName, { bookType: "xlsx", cellStyles: true });
}
