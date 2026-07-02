// Merges parsed Makro price data into parsed Report 1145 rows, producing the
// Tabelle1-shaped rows the XML generator consumes.
//
// Ported from the VBA Get_Data_From_FileReport1145 (Module2). Key differences
// from the P2P merger:
//   - The DRIVING list is Report 1145 (every 1145 row appears in the output),
//     NOT the Makro file. Makro codes with no matching 1145 article are ignored
//     (counted in the summary only), exactly like the workbook's XLOOKUP does.
//   - The join key is Report 1145 "Article no." (itemNo) === Makro "รหัสสินค้า".
//   - Lead time is set UNCONDITIONALLY (no toggle): matched → "1", unmatched → "0".
//
// Per-row rule (VBA):
//   AA "New Price per order unit"      = XLOOKUP(Article no., รหัสสินค้า, L, "No Information")
//   AB "New Price per order unit(Adj)" = ROUND( XLOOKUP(..., L, <old 1145 price>), 2)
//   AC "Diff Price Between Old and New"= <old 1145 price> - AB
//   AF "Article lead time(Adj)"        = IIf(AA = "No Information", "0", "1")
//
// So:
//   matched   → price = Round(Makro L, 2)   (Makro.newPrice); lead time = "1"
//   unmatched → price = old 1145 price;      lead time = "0"; flagged "No Information"

import { vbaRound, clean2 } from "./makroParser.js";

// Status labels carried on each merged row (single-status shape, like P2P).
export const MAKRO_STATUS = {
  MATCHED:      "Price updated from Makro",
  NO_INFO:      "No Information",
  DISCONTINUED: "Discontinued",
  NO_PRICE:     "No Makro price",
};

function r1145PriceOf(r) {
  const v = r?.priceOU;
  return typeof v === "number" && Number.isFinite(v) ? v : 0;
}

// A Makro row is "discontinued" when its สถานะ contains "discontinu"
// (case-insensitive) — covers "Discontinue", "Discontinued", "DISCONTINUED",
// etc. Anything else (active, blank, missing) is treated as normal/active.
function isDiscontinued(makroRow) {
  return String(makroRow?.status || "").trim().toLowerCase().includes("discontinu");
}

// True when the Makro row has a usable (> 0) computed new price. A blank/zero
// In-VAT price yields newPrice ≤ 0 (or negative when a VAT amount is present);
// such rows must NOT overwrite the Report 1145 price.
function hasUsableMakroPrice(makroRow) {
  const p = makroRow?.newPrice;
  return typeof p === "number" && Number.isFinite(p) && p > 0;
}

// The Makro-derived display columns (raw file values + VAT calc), shared by the
// matched and discontinued row builders.
function makroDisplayFields(makroRow) {
  return {
    // ----- Raw Makro columns (red in the full-data view) -----
    __makroArtGroup:  makroRow.artGroup,
    __makroName:      makroRow.itemName,
    __makroExVat:     makroRow.srcExVat,         // ราคาขาย (Ex. VAT) — raw from file
    __makroVat:       makroRow.vatAmt,           // VAT amount
    __makroInVat:     makroRow.inVat,            // ราคาขาย (In. VAT)
    __makroStatus:    makroRow.status,           // สถานะ
    // ----- VAT calculation columns (yellow in the full-data view) -----
    __makroVatPct:        makroRow.vatPct,        // H  VAT%
    __makroPriceExVat:    makroRow.priceExVat,    // I  Price Exclude VAT
    __makroPriceInVat:    makroRow.priceInVat,    // J  Price Include VAT
    __makroDiffDecimal:   makroRow.diffDecimal,   // K  Diff (Decimal)
    __makroExVatAdj:      makroRow.priceExVatAdj, // L  Price Exclude VAT(Adj)
    __makroPriceInVatAdj: makroRow.priceInVatAdj, // M  Price Include VAT(Adj)
    __makroCheckDiff:     makroRow.checkDiff,     // N  Check Diff
  };
}

// Matched row: 1145 base with the Makro new price + lead time "1" overlaid.
function buildMatchedRow({ r1145Row, makroRow, pos }) {
  const oldPrice = r1145PriceOf(r1145Row);
  const newPrice = makroRow.newPrice;   // = ROUND(Makro L, 2)
  return {
    ...r1145Row,
    pos,
    priceOU:      newPrice,
    availability: "1",       // AF — matched ⇒ open lead time
    leadTimeRaw:  "1",       // keep in sync so validation reflects the XML output
    // The Makro price replaces the 1145 price entirely, so the 1145 parser's
    // scaled-price flags no longer describe this row — clear them to avoid
    // misleading "Scaled price…" warnings from the validator.
    __scaledPriceWasZero: false,
    __scaledPriceWasBlank: false,
    __priceBothBlank: false,
    // Drop the 1145 multi-pill status array so the single-status model (below)
    // is the one the modal filter / Excel export read.
    statuses: undefined,
    // Preview / full-data extras
    __oldPrice:   oldPrice,
    __newPrice:   newPrice,
    __diff:       clean2(oldPrice - newPrice),   // AC = old - new (dust snapped to 0)
    ...makroDisplayFields(makroRow),
    status:   MAKRO_STATUS.MATCHED,
    __source: "matched",
  };
}

// Matched code, but the Makro price is NOT applied — keep the Report 1145 price
// and close the lead time to "0". Used for two cases:
//   • Discontinue    — สถานะ says the item is discontinued
//   • No Makro price — the Makro row has no usable price (≤ 0 / blank)
// Makro/VAT columns are still populated (and highlighted) for reference, but
// __newPrice is blank since nothing was applied. Stays __source "matched" so the
// red/yellow highlighting fires; `extra` carries the flag driving the counts.
function buildKeepR1145Row({ r1145Row, makroRow, pos, status, extra }) {
  const oldPrice = r1145PriceOf(r1145Row);
  return {
    ...r1145Row,
    pos,
    priceOU:      vbaRound(oldPrice, 2),   // keep the Report 1145 price
    availability: "0",                     // close the lead time
    leadTimeRaw:  "0",
    __scaledPriceWasZero: false,
    __scaledPriceWasBlank: false,
    __priceBothBlank: false,
    statuses: undefined,
    __oldPrice:   oldPrice,
    __newPrice:   "",                      // Makro price NOT applied
    __diff:       "",
    ...makroDisplayFields(makroRow),
    status,
    __source: "matched",
    ...extra,
  };
}

// Unmatched 1145 row: keep its existing price, lead time "0", flag "No Information".
function buildNoInfoRow({ r1145Row, pos }) {
  const oldPrice = r1145PriceOf(r1145Row);
  return {
    ...r1145Row,
    pos,
    priceOU:      vbaRound(oldPrice, 2),   // VBA: AB = ROUND(Z, 2) when not found
    availability: "0",        // AF — not found ⇒ lead time 0
    leadTimeRaw:  "0",
    // Single-status model — drop the 1145 multi-pill array (see buildMatchedRow).
    statuses: undefined,
    __oldPrice:   oldPrice,
    __newPrice:   "",
    __diff:       "",
    // Raw Makro + VAT-calc columns are empty for rows with no Makro match.
    __makroArtGroup:  "",
    __makroName:      "",
    __makroExVat:     "",
    __makroVat:       "",
    __makroInVat:     "",
    __makroStatus:    "",
    __makroVatPct:        "",
    __makroPriceExVat:    "",
    __makroPriceInVat:    "",
    __makroDiffDecimal:   "",
    __makroExVatAdj:      "",
    __makroPriceInVatAdj: "",
    __makroCheckDiff:     "",
    status:   MAKRO_STATUS.NO_INFO,
    __source: "no-info",
  };
}

/**
 * @param {Array} r1145Rows  rows from parseReport1145(...)
 * @param {Array} makroRows  rows from parseMakroFile(...).rows
 * @returns {{rows: Array, summary: {total, matched, noInfo, makroOnly}}}
 */
export function mergeMakroAndReport1145(r1145Rows, makroRows) {
  // Makro dict keyed by product code (first occurrence wins).
  const makroByCode = new Map();
  for (const m of makroRows) {
    const key = String(m.artCode || "").trim();
    if (key && !makroByCode.has(key)) makroByCode.set(key, m);
  }

  const merged = [];
  const usedCodes = new Set();
  let pos = 1;

  // Drive by the Report 1145 list — every 1145 row appears in the output.
  for (const r of r1145Rows) {
    const key = String(r.itemNo || "").trim();
    const m = key ? makroByCode.get(key) : undefined;
    if (m) {
      if (isDiscontinued(m)) {
        // Discontinued in Makro → keep the 1145 price, lead time 0.
        merged.push(buildKeepR1145Row({
          r1145Row: r, makroRow: m, pos: pos++,
          status: MAKRO_STATUS.DISCONTINUED, extra: { __discontinued: true },
        }));
      } else if (!hasUsableMakroPrice(m)) {
        // Makro price ≤ 0 / blank → don't overwrite; keep the 1145 price, lead time 0.
        merged.push(buildKeepR1145Row({
          r1145Row: r, makroRow: m, pos: pos++,
          status: MAKRO_STATUS.NO_PRICE, extra: { __noMakroPrice: true },
        }));
      } else {
        merged.push(buildMatchedRow({ r1145Row: r, makroRow: m, pos: pos++ }));
      }
      usedCodes.add(key);
    } else {
      merged.push(buildNoInfoRow({ r1145Row: r, pos: pos++ }));
    }
  }

  // Makro codes never matched to any 1145 article. These don't appear in the
  // XML output (the VBA's XLOOKUP is 1145-driven), but we return them so the UI
  // can list them in a separate viewer. Deduped by product code, matching the
  // lookup semantics; a sequential `pos` is added for display.
  const makroOnlyRows = [];
  let posU = 1;
  for (const [key, m] of makroByCode) {
    if (!usedCodes.has(key)) makroOnlyRows.push({ ...m, pos: posU++ });
  }
  const makroOnly = makroOnlyRows.length;

  const summary = {
    total:        merged.length,
    // "matched" = price actually updated (found, active, and had a usable price).
    // Discontinued and no-price rows matched too but kept the 1145 price, so
    // they're counted separately.
    matched:      merged.filter((m) => m.__source === "matched" && !m.__discontinued && !m.__noMakroPrice).length,
    discontinued: merged.filter((m) => m.__discontinued).length,
    noPrice:      merged.filter((m) => m.__noMakroPrice).length,
    noInfo:       merged.filter((m) => m.__source === "no-info").length,
    makroOnly,
  };

  return { rows: merged, summary, makroOnlyRows };
}
