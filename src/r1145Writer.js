// Build a Report 1145 .xlsx file from rows produced by xmlParser.js.
//
// Output layout (compact form, simpler than the original Report 1145 template):
//   Row 1: column headers
//   Row 2+: data rows
//
// Per the business spec for this conversion direction:
//   - ID column: blank (user requested)
//   - Type column: blank (user requested)
//   - Price goes into column O ("Price per order unit") only
//   - Column P ("Scaled price") stays blank
//   - Customer ID appended as the final column (column W / 23)
//
// Note: this layout (no "Create import list (PRI)" title row) is intentional
// and matches the user's request. If users want to re-upload the file through
// the "Convert XML from 1145 report" tab, they'll need to either add a title
// row themselves or adjust the parser. The output is for review/editing, not
// necessarily round-trip.

import XLSX from "xlsx-js-style";

// FutureLog brand orange — warmer/amber shade picked from the actual logo
// (screenshot reference). Stays in the "FF" alpha-prefixed hex format that
// Excel expects.
const FUTURELOG_ORANGE = "FFF0A02E";   // amber-orange, matches FutureLog brand
const HEADER_GRAY      = "FFD9D9D9";   // neutral gray for empty-data columns
const HEADER_TEXT_DARK = "FF1A1A1A";   // near-black dark text for high contrast

// Shared font/border/alignment used by both the orange ("has data") and gray
// ("no data") header variants. Only the fill differs between them.
const HEADER_FONT = { bold: true, color: { rgb: HEADER_TEXT_DARK }, sz: 11 };
const HEADER_BORDER = {
  top:    { style: "thin", color: { rgb: "FFB0B0B0" } },
  bottom: { style: "thin", color: { rgb: "FFB0B0B0" } },
  left:   { style: "thin", color: { rgb: "FFB0B0B0" } },
  right:  { style: "thin", color: { rgb: "FFB0B0B0" } },
};
const HEADER_ALIGN = { horizontal: "left", vertical: "center", wrapText: true };

const HEADER_STYLE_ORANGE = {
  font: HEADER_FONT,
  fill: { patternType: "solid", fgColor: { rgb: FUTURELOG_ORANGE } },
  alignment: HEADER_ALIGN,
  border: HEADER_BORDER,
};
const HEADER_STYLE_GRAY = {
  font: HEADER_FONT,
  fill: { patternType: "solid", fgColor: { rgb: HEADER_GRAY } },
  alignment: HEADER_ALIGN,
  border: HEADER_BORDER,
};

// Column order matching the standard Report 1145 template, plus Customer ID
// at the end (column W / 23). Each entry maps a row-object key to the header
// label that appears in row 1.
const COLUMNS = [
  { key: null,        label: "ID" },              // intentionally blank per spec
  { key: null,        label: "Type" },            // intentionally blank per spec
  { key: "descDE",    label: "Item name (German)" },
  { key: "descFR",    label: "Item name (French)" },
  { key: "descIT",    label: "Item name (Italian)" },
  { key: "descGB",    label: "Item name (English)" },
  { key: "descExtra", label: "Item name" },
  { key: "itemNo",    label: "Article no." },
  { key: "ean",       label: "GTIN" },
  { key: "manArtId",  label: "Manufacturer's item number" },
  { key: null,        label: "Producer" },         // not in XML; stays blank
  { key: "ou",        label: "Order unit (OU)" },
  { key: "cu",        label: "Content unit (CU)" },
  { key: "cuou",      label: "Packaging unit" },
  { key: "priceOU",   label: "Price per order unit" },  // column O — gets the value
  { key: null,        label: "Scaled price" },     // column P — intentionally blank
  { key: "origin",    label: "Country of origin" },
  { key: "customsNo", label: "Customs tariff number" },
  { key: "availability", label: "Article<BR>lead time" },
  { key: "specUrl",   label: "Item link supplier<br>" },
  { key: "offerStart",label: "Start of special offer" },
  { key: "offerEnd",  label: "End of special offer" },
  { key: "customerId",label: "Customer ID" },      // appended per spec
];

// Approximate column widths (in characters) for the .xlsx. Tuned to the
// content type in each column so the file looks reasonable when opened.
const COL_WIDTHS = [
  10,  // ID
  6,   // Type
  22, 22, 22, 32, 32,  // 5 description columns
  14,  // Article no.
  16,  // GTIN
  22,  // Mfg item no
  14,  // Producer
  14, 14, 14,  // OU, CU, Packaging unit
  16,  // Price per order unit
  14,  // Scaled price
  18, 18,  // Origin, Customs no
  14,  // Lead time
  22,  // Spec URL
  16, 16,  // Offer start/end
  14,  // Customer ID
];

/**
 * Build the .xlsx workbook in memory and return it as a downloadable Blob.
 *
 * @param {object[]} rows  Output of xmlParser.parseFuturelogXml().rows
 * @returns {{ blob: Blob, suggestedFilename: string }}
 */
export function buildReport1145Xlsx(rows, opts = {}) {
  // Build the array-of-arrays sheet content.
  // Headers go in row 1, data starts at row 2.
  //
  // Header styling rule (per-column, per-file):
  //   - "has data" columns get orange fill (FutureLog brand color)
  //   - "all empty" columns get gray fill
  //   - Columns intentionally blank (key === null) are always gray
  //
  // The detection runs once over the dataset before building the header so
  // we can pick the right style for each cell. If a column had no data in
  // this file but gets filled in a future file, it'll automatically switch
  // to orange — no code change needed.
  //
  // xlsx-js-style accepts cell objects of the form { v, t, s } where:
  //   v = the value to display
  //   t = type indicator: "s" = string, "n" = number
  //   s = style object (fonts, fills, borders, etc.)

  // Step 1: detect which columns have at least one non-empty value
  const columnHasData = COLUMNS.map((c) => {
    if (c.key === null) return false;        // intentionally blank columns
    for (const r of rows) {
      const v = r[c.key];
      if (v === null || v === undefined) continue;
      // Treat empty strings as no data; everything else (incl. 0, 0.0) counts.
      if (typeof v === "string" && v.trim() === "") continue;
      return true;
    }
    return false;
  });

  // Step 2: build the styled header row using the detection result
  const headerRow = COLUMNS.map((c, i) => ({
    v: c.label,
    t: "s",
    s: columnHasData[i] ? HEADER_STYLE_ORANGE : HEADER_STYLE_GRAY,
  }));

  const aoa = [headerRow];

  for (const r of rows) {
    aoa.push(COLUMNS.map((c) => {
      if (c.key === null) return "";          // intentionally blank columns
      const v = r[c.key];
      if (v === null || v === undefined) return "";
      return v;
    }));
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = COL_WIDTHS.map((w) => ({ wch: w }));

  // Freeze the header row so it stays visible while scrolling — useful when
  // the file has hundreds of articles.
  ws["!views"] = [{ state: "frozen", ySplit: 1 }];

  // Enable autofilter on the header row so users can filter/sort columns
  // directly in Excel. Range covers all header columns and all data rows.
  const lastColLetter = XLSX.utils.encode_col(COLUMNS.length - 1);
  const lastRow = aoa.length; // 1-indexed; header is row 1, data fills the rest
  ws["!autofilter"] = { ref: `A1:${lastColLetter}${lastRow}` };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Tabelle1");

  // Generate as a Uint8Array, wrap in a Blob
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const today = new Date().toISOString().slice(0, 10);
  const suggestedFilename =
    opts.filename ||
    `Report_1145_from_XML_${today}.xlsx`;

  return { blob, suggestedFilename };
}
