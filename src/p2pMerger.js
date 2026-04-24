// Merges parsed P2P data with parsed Report 1145 data into the Tabelle1-shaped
// row format used by the XML generator.
//
// Matches the VBA logic from Module8 (Get_Data_From_Report1145_revised) of
// Importlist_TH_XML (For MDM)__new.xlsm:
//
//  - Iterate each P2P row (keyed by Article No.)
//  - Look up the matching Report 1145 row by supplier article no. (R1145 `itemNo`)
//  - If matched: merge R1145 descriptions/units/origin with the P2P price
//  - Apply price / lead-time rules (see computePriceAndStatus)
//  - Append any R1145 items NOT found in P2P with status "No price update"
//
// Output: rows in the same shape as parseReport1145, plus a `status` field
// carrying one of:
//    ""                         — ordinary, price came from P2P as expected
//    "Open lead time"           — lead-time flipped to 1 because new price > 1
//    "Price from report 1145"   — P2P price missing/zero, fell back to R1145
//    "No price update"          — item only in R1145, no matching P2P row
//    "P2P-only item"            — item only in P2P, no matching R1145 row

/**
 * @typedef {Object} MergeOptions
 * @property {boolean} useNewPriceCol  use the NEW PRICE column instead of Price / Order unit
 * @property {boolean} openLeadTime    flip availability to "1" when a new price > 1 exists
 */

// Extract the effective P2P price for a given row, honoring the toggle.
function p2pPriceFor(p2pRow, useNewPriceCol) {
  const v = useNewPriceCol ? p2pRow.newPrice : p2pRow.priceOrderUnit;
  return typeof v === "number" && Number.isFinite(v) ? v : 0;
}

// The R1145 "effective price" is whatever parseReport1145 already computed into `priceOU`.
function r1145PriceOf(r) {
  const v = r?.priceOU;
  return typeof v === "number" && Number.isFinite(v) ? v : 0;
}

// Decide the final price + availability + status for a merged row.
// Mirrors Module8's price-fallback + Open Lead Time block.
function computePriceAndStatus({ p2pRow, r1145Row, opts }) {
  const useNewPriceCol = !!opts.useNewPriceCol;
  const openLeadTime   = !!opts.openLeadTime;

  const p2pPrice = p2pRow ? p2pPriceFor(p2pRow, useNewPriceCol) : 0;
  const r1145Price = r1145Row ? r1145PriceOf(r1145Row) : 0;
  const r1145Lt    = r1145Row ? r1145Row.leadTimeRaw : "";

  // Fallback rule (VBA: "If Cells(b, 23).Value <= 1 Then ... Price from report 1145"):
  // when P2P price is missing or <= 0, use R1145 price and flag the row.
  let finalPrice, status, availability;

  if (p2pPrice > 0) {
    finalPrice = p2pPrice;
    availability = r1145Lt !== "" && r1145Lt !== null && r1145Lt !== undefined
      ? r1145Lt
      : "0";

    if (openLeadTime && p2pPrice > 1) {
      // Module8: "Open lead time if there is price in new price column"
      // sets lead time to 1 and stamps "Open lead time".
      availability = "1";
      status = "Open lead time";
    } else {
      status = "";
    }
  } else {
    // No usable P2P price → fall back to Report 1145
    finalPrice = r1145Price || 0;
    availability = r1145Lt !== "" && r1145Lt !== null && r1145Lt !== undefined
      ? r1145Lt
      : "0";
    status = "Price from report 1145";
  }

  return { finalPrice, status, availability };
}

// Produce a Tabelle1 row starting from an R1145 base, overlaying P2P where relevant.
// This is for matched rows (both P2P and R1145 present).
function buildMatchedRow({ p2pRow, r1145Row, opts, pos }) {
  const { finalPrice, status, availability } = computePriceAndStatus({ p2pRow, r1145Row, opts });

  return {
    ...r1145Row,
    pos,
    priceOU:      finalPrice,
    availability,
    // Preserve an old-price reference for the preview diff column
    __oldPrice:   r1145PriceOf(r1145Row),
    __newPrice:   p2pPriceFor(p2pRow, opts.useNewPriceCol),
    status,
    // Where the row came from — useful for "show only P2P items" filters later
    __source:     "matched",
  };
}

// Row for a P2P item that had no matching R1145 entry. Descriptions, units,
// origin are unknown — validator will flag missing mandatory fields.
function buildP2POnlyRow({ p2pRow, opts, pos }) {
  const price = p2pPriceFor(p2pRow, opts.useNewPriceCol);
  return {
    pos,
    descDE: "",
    descFR: "",
    descIT: "",
    descGB: p2pRow.itemName || "",
    descExtra: "",
    itemNo: p2pRow.articleNo,
    ean:    p2pRow.gtin || "",
    manArtId: "",
    manLiefID: "",
    // P2P unit labels ("Package", "Can"…) are not in the MDM Unit_List — the
    // validator will catch this and prompt the user to edit.
    ou: p2pRow.orderUnit || "",
    cu: p2pRow.contentUnits || "",
    cuou: p2pRow.packagingUnit || "",
    priceOU: price,
    priceLevel: "",
    origin: p2pRow.originCountry || "",
    customsNo: "",
    availability: opts.openLeadTime && price > 1 ? "1" : "0",
    leadTimeRaw:  opts.openLeadTime && price > 1 ? "1" : "0",
    specUrl: "", offerStart: "", offerEnd: "",
    customerId: "0000",
    __oldPrice: "",
    __newPrice: price,
    status: "P2P-only item",
    __source: "p2p-only",
  };
}

// Row for an R1145 item that had no matching P2P row — appended at the end.
function buildR1145OnlyRow({ r1145Row, pos }) {
  return {
    ...r1145Row,
    pos,
    __oldPrice: r1145PriceOf(r1145Row),
    __newPrice: "",
    status: "No price update",
    __source: "r1145-only",
  };
}

/**
 * @param {Array} r1145Rows   rows from parseReport1145(...)
 * @param {Array} p2pRows     rows from parseP2PFile(...).rows
 * @param {MergeOptions} opts
 * @returns {{rows: Array, summary: {matched:number, p2pOnly:number, r1145Only:number, total:number}}}
 */
export function mergeP2PAndReport1145(r1145Rows, p2pRows, opts) {
  const options = {
    useNewPriceCol: !!opts?.useNewPriceCol,
    openLeadTime:   !!opts?.openLeadTime,
  };

  // Build R1145 dict by itemNo (= supplier article number), using the FIRST occurrence.
  // Duplicates are flagged as errors by the normal validator, so we don't fuss over them here.
  const r1145ByArticle = new Map();
  for (const r of r1145Rows) {
    const key = String(r.itemNo || "").trim();
    if (key && !r1145ByArticle.has(key)) r1145ByArticle.set(key, r);
  }

  const merged = [];
  const matchedKeys = new Set();
  let pos = 1;

  // 1) Walk P2P as the driving list
  for (const p of p2pRows) {
    const key = String(p.articleNo || "").trim();
    if (!key) continue;

    const r = r1145ByArticle.get(key);
    if (r) {
      merged.push(buildMatchedRow({ p2pRow: p, r1145Row: r, opts: options, pos: pos++ }));
      matchedKeys.add(key);
    } else {
      merged.push(buildP2POnlyRow({ p2pRow: p, opts: options, pos: pos++ }));
    }
  }

  // 2) Append R1145 items that were not in P2P, with "No price update"
  for (const [key, r] of r1145ByArticle) {
    if (!matchedKeys.has(key)) {
      merged.push(buildR1145OnlyRow({ r1145Row: r, pos: pos++ }));
    }
  }

  const summary = {
    total:      merged.length,
    matched:    merged.filter((m) => m.__source === "matched").length,
    p2pOnly:    merged.filter((m) => m.__source === "p2p-only").length,
    r1145Only:  merged.filter((m) => m.__source === "r1145-only").length,
  };

  return { rows: merged, summary };
}
