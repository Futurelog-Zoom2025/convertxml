// Parses a Report 1145 Excel file and maps it to the 23-column Tabelle1 format
// used by the FUTURELOG XML generator.
//
// Columns are located by their HEADER NAME on row 4 (not by fixed position),
// so extra columns in the uploaded file are safely ignored and column order
// doesn't matter. Header matching is tolerant: case-insensitive, whitespace-
// collapsed, and <br>/<BR> tags stripped.
//
// VBA reference (Get_Data_From_File1145):
//   Data starts at row 5, stops at the first blank "Type" row.
//   Price logic:
//     priceOU = scaledPrice (rounded to 2dp) if scaledPrice is a positive number, else unitPrice.
//     availability = leadTime when scaled branch was used, else "0".

const HEADER_ROW_INDEX = 3; // Excel row 4

// Sentinel used internally so the validator can distinguish "cell was #N/A"
// from "cell was blank". The XML generator treats this as empty.
export const NA_MARKER = "__NA__";

// Canonical Report 1145 header names -> internal field keys used by the rest
// of the pipeline. Each header on the LEFT is how it appears (roughly) in the
// real file; values on the RIGHT are the Tabelle1 field keys.
//
// - `display`: the canonical human-readable header name (shown in error messages)
// - `aliases`: all accepted spellings (matched case-insensitively, whitespace-collapsed,
//    with <br> tags stripped). First match wins.
// - `mandatory`: if true, the file is rejected when no matching header is found.
const HEADER_MAP = {
  id:               { display: "ID",                          aliases: ["id"],                                         mandatory: false },
  recordType:       { display: "Type",                        aliases: ["type"],                                       mandatory: true  }, // structurally required: parser uses empty Type as end-of-data signal (matches VBA `Do Until Cells(Z, 2) = ""`)
  descDE:           { display: "Item name (German)",          aliases: ["item name (german)", "item name german"],     mandatory: false },
  descFR:           { display: "Item name (French)",          aliases: ["item name (french)", "item name french"],     mandatory: false },
  descIT:           { display: "Item name (Italian)",         aliases: ["item name (italian)", "item name italian"],   mandatory: false },
  descGB:           { display: "Item name (English)",         aliases: ["item name (english)", "item name english"],   mandatory: false },
  descExtra:        { display: "Item name",                   aliases: ["item name"],                                  mandatory: false },
  itemNo:           { display: "Article no.",                 aliases: ["article no.", "article no", "article number", "artikel nr.", "artikel nr"], mandatory: true  },
  ean:              { display: "GTIN",                        aliases: ["gtin", "ean"],                                mandatory: false },
  manArtId:         { display: "Manufacturer's item number",  aliases: ["manufacturer's item number", "manufacturers item number", "manufacturer item number"], mandatory: false },
  producer:         { display: "Producer",                    aliases: ["producer", "manufacturer"],                   mandatory: false },
  ou:               { display: "Order unit (OU)",             aliases: ["order unit (ou)", "order unit"],              mandatory: true  },
  cu:               { display: "Content unit (CU)",           aliases: ["content unit (cu)", "content unit"],          mandatory: true  },
  cuou:             { display: "Packaging unit",              aliases: ["packaging unit"],                             mandatory: false },
  priceUnit:        { display: "Price per order unit",        aliases: ["price per order unit", "price"],              mandatory: false },
  scaledPrice:      { display: "Scaled price",                aliases: ["scaled price"],                               mandatory: false },
  origin:           { display: "Country of origin",           aliases: ["country of origin", "origin"],                mandatory: false },
  customsNo:        { display: "Customs tariff number",       aliases: ["customs tariff number", "customs no.", "customs no", "customs number"], mandatory: false },
  leadTime:         { display: "Article lead time",           aliases: ["article lead time", "lead time", "articlelead time", "article<br>lead time"], mandatory: true  },
  specUrl:          { display: "Item link supplier",          aliases: ["item link supplier", "spec url", "supplier link"], mandatory: false },
  offerStart:       { display: "Start of special offer",      aliases: ["start of special offer", "offer start"],      mandatory: false },
  offerEnd:         { display: "End of special offer",        aliases: ["end of special offer", "offer end"],          mandatory: false },
};

// Normalize a header string so minor differences don't break matching.
//  - strip <br>, <BR>, <br/>, <br /> tags
//  - lowercase
//  - collapse all whitespace into a single space
//  - trim
function normalizeHeader(h) {
  if (h === null || h === undefined) return "";
  return String(h)
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function cleanCell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") {
    const t = v.trim();
    if (t === "#N/A" || t === "N/A" || t === "#N/A N/A") return NA_MARKER;
    return t;
  }
  if (typeof v === "number") {
    if (Number.isNaN(v)) return "";
    return v;
  }
  if (v instanceof Date) {
    const d = String(v.getDate()).padStart(2, "0");
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const y = v.getFullYear();
    return `${d}.${m}.${y}`;
  }
  return String(v);
}

function toNumericOrEmpty(v) {
  if (v === "" || v === null || v === undefined || v === NA_MARKER) return "";
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "") return "";
  const n = Number(s);
  return Number.isFinite(n) ? n : "";
}

/**
 * Scan the header row and return an index map { fieldKey: columnIndex | -1 }.
 * Extra / unknown columns are ignored silently.
 */
function resolveHeaders(headerRow) {
  // Build normalized -> column-index lookup
  const headerToIdx = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    const norm = normalizeHeader(headerRow[i]);
    if (norm !== "" && !headerToIdx.has(norm)) {
      headerToIdx.set(norm, i);
    }
  }

  const resolved = {};
  const missingMandatory = [];

  for (const [fieldKey, spec] of Object.entries(HEADER_MAP)) {
    let idx = -1;
    for (const alias of spec.aliases) {
      const normAlias = normalizeHeader(alias);
      if (headerToIdx.has(normAlias)) {
        idx = headerToIdx.get(normAlias);
        break;
      }
    }
    resolved[fieldKey] = idx;
    if (idx === -1 && spec.mandatory) {
      missingMandatory.push(spec.display || spec.aliases[0]);
    }
  }

  return { resolved, missingMandatory };
}

// Safely read a cell from a source row using the resolved column index.
// Returns "" if the column wasn't found in the header row.
function getCell(src, resolved, key) {
  const idx = resolved[key];
  if (idx === -1 || idx === undefined) return "";
  return cleanCell(src[idx]);
}

/**
 * @param {Array<Array<any>>} aoa  sheet parsed as array-of-arrays (header:1 from SheetJS)
 * @returns {Array<Object>}        rows in Tabelle1 format
 */
export function parseReport1145(aoa) {
  if (!Array.isArray(aoa) || aoa.length <= HEADER_ROW_INDEX) {
    throw new Error("Report 1145 appears to be empty or malformed (no data below header row 4).");
  }

  const headerRow = aoa[HEADER_ROW_INDEX] || [];
  const { resolved, missingMandatory } = resolveHeaders(headerRow);

  if (missingMandatory.length) {
    throw new Error(
      "This doesn't look like a Report 1145 file. Missing required header" +
        (missingMandatory.length === 1 ? "" : "s") +
        " on row 4: " +
        missingMandatory.map((h) => `"${h}"`).join(", ") +
        "."
    );
  }

  const rows = [];

  for (let r = HEADER_ROW_INDEX + 1; r < aoa.length; r++) {
    const src = aoa[r] || [];

    // Stop at the first row with empty "Type" — matches VBA `Do Until Cells(Z, 2) = ""`.
    const typeVal = getCell(src, resolved, "recordType");
    if (typeVal === "" || typeVal === 0 || typeVal === NA_MARKER) break;

    // Price logic (ported from VBA)
    const priceOU = toNumericOrEmpty(src[resolved.priceUnit]);
    const scaled = toNumericOrEmpty(src[resolved.scaledPrice]);
    const leadTime = getCell(src, resolved, "leadTime");

    const scaledRounded =
      typeof scaled === "number" ? Math.round(scaled * 100) / 100 : "";

    let finalPrice;
    let usedScaled = false;
    if (scaledRounded === "" || scaledRounded === 0) {
      finalPrice = priceOU === "" ? 0 : priceOU;
    } else {
      finalPrice = scaledRounded;
      usedScaled = true;
    }

    const availability = usedScaled ? leadTime : "0";

    const row = {
      pos: rows.length + 1,
      recordType: getCell(src, resolved, "recordType"),
      descDE: getCell(src, resolved, "descDE"),
      descFR: getCell(src, resolved, "descFR"),
      descIT: getCell(src, resolved, "descIT"),
      descGB: getCell(src, resolved, "descGB"),
      descExtra: getCell(src, resolved, "descExtra"),
      itemNo: String(getCell(src, resolved, "itemNo")).trim(),
      ean: String(getCell(src, resolved, "ean")).trim(),
      manArtId: getCell(src, resolved, "manArtId"),
      manLiefID: "", // Not present in Report 1145; Tabelle1 slot retained
      ou: getCell(src, resolved, "ou"),
      cu: getCell(src, resolved, "cu"),
      cuou: getCell(src, resolved, "cuou"),
      priceOU: finalPrice,
      priceLevel: "",
      origin: getCell(src, resolved, "origin"),
      customsNo: getCell(src, resolved, "customsNo"),
      availability: availability,      // computed value written to XML <VLZ>
      leadTimeRaw: leadTime,           // raw value from source file, used only for validation
      specUrl: getCell(src, resolved, "specUrl"),
      offerStart: getCell(src, resolved, "offerStart"),
      offerEnd: getCell(src, resolved, "offerEnd"),
      customerId: "0000",
    };
    rows.push(row);
  }

  return rows;
}
