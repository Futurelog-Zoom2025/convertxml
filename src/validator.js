// Port of the VBA validation logic in generate_xml / Convert_to_UNI
// Returns { errors: [...], warnings: [...], invalidCells: Map<rowIdx,Set<col>> }

import { VALID_UNITS, VALID_COUNTRIES, VALID_LANGUAGES } from "./referenceData.js";
import { NA_MARKER } from "./reportParser.js";

// Tabelle1 columns that are mandatory.
// VBA:  For col = 1 To lastCol: If Trim(Cells(2, col).Value) <> "" And col <> 20 Then add to mandatory
// Column 20 (specUrl / T) is excluded. In practice Report 1145 never populates columns 4, 5, 16, 21, 22, 23
// on row 2, so those would not be treated as mandatory by the VBA either.
// To be faithful, we treat as mandatory ONLY the columns that have a value in row 2 AND are not col 20.
// We evaluate that dynamically from the first row of rows[] (matching the VBA reference to row 2).
//
// This mirrors "whatever row-2 has is mandatory" behavior exactly.
const EXCLUDED_MANDATORY_KEYS = new Set(["specUrl"]); // col 20 in VBA

const ROW_KEYS = [
  "pos","recordType","descDE","descFR","descIT","descGB","descExtra",
  "itemNo","ean","manArtId","manLiefID","ou","cu","cuou","priceOU",
  "priceLevel","origin","customsNo","availability","leadTimeRaw","specUrl","offerStart",
  "offerEnd","customerId",
];

// Human-readable labels for error messages
const KEY_LABELS = {
  pos: "Pos",
  recordType: "Record Type",
  descDE: "German description",
  descFR: "French description",
  descIT: "Italian description",
  descGB: "English description",
  descExtra: "Local description",
  itemNo: "Article no.",
  ean: "EAN",
  manArtId: "Manufacturer Article ID",
  manLiefID: "Manufacturer/Supplier ID",
  ou: "Order unit (OU)",
  cu: "Content unit (CU)",
  cuou: "CU per OU",
  priceOU: "Price / OU",
  priceLevel: "Price Level",
  origin: "Country of origin",
  customsNo: "Customs No.",
  availability: "Availability",
  leadTimeRaw: "Article lead time",
  specUrl: "Spec URL",
  offerStart: "Offer Start",
  offerEnd: "Offer End",
  customerId: "Customer ID",
};

function isBlank(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === "string") return v.trim() === "";
  return false;
}

function isNA(v) {
  return v === NA_MARKER;
}

export function validate(rows, params) {
  const errors = [];
  const warnings = [];
  // rowIdx is 0-based in `rows`; the user-facing row number in the VBA is rowIdx + 2
  const invalidCells = new Map();

  const markCell = (rowIdx, colKey) => {
    if (!invalidCells.has(rowIdx)) invalidCells.set(rowIdx, new Set());
    invalidCells.get(rowIdx).add(colKey);
  };

  if (rows.length === 0) {
    errors.push("No data rows found in the uploaded file.");
    return { errors, warnings, invalidCells };
  }

  // Step 0: parameter validation (mirrors Step 4 in VBA; we run it upfront for better UX)
  if (!/^\d{3}$/.test(params.companyId || "")) {
    errors.push("WebShop Company ID must be exactly 3 numeric digits (e.g. 169).");
  }
  if (!/^\d{6}$/.test(params.supplierNo || "")) {
    errors.push("Supplier Number must be exactly 6 numeric digits (e.g. 123456).");
  }
  if (!VALID_LANGUAGES.includes(String(params.language || "").toUpperCase())) {
    errors.push(`Language must be one of: ${VALID_LANGUAGES.join(", ")}.`);
  }
  if (!/^\d{8}$/.test(params.validityDate || "")) {
    errors.push("Validity Date must be exactly 8 numeric digits in DDMMYYYY format (e.g. 31122025).");
  }

  // Step 1: Customer ID "0000" core-entry rule
  // If ANY row has customerId = "0000", then EVERY itemNo must have at least one "0000" row.
  const itemHas0000 = new Set();
  const allItemNos = new Set();
  let anyHas0000 = false;
  rows.forEach((r) => {
    const itemNo = String(r.itemNo || "").trim();
    const cust = String(r.customerId || "").trim();
    if (itemNo !== "") {
      allItemNos.add(itemNo);
      if (cust === "0000") {
        anyHas0000 = true;
        itemHas0000.add(itemNo);
      }
    }
  });
  if (anyHas0000) {
    const missing = [];
    for (const it of allItemNos) {
      if (!itemHas0000.has(it)) missing.push(it);
    }
    if (missing.length) {
      errors.push(
        `The following Item Nos do not have a Customer ID = "0000" core entry:\n  ${missing.join(", ")}`
      );
      rows.forEach((r, idx) => {
        if (missing.includes(String(r.itemNo || "").trim())) {
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
        }
      });
    }
  }

  // Step 2: Mandatory fields — an explicit list agreed with the business.
  // Every data row must have a non-blank, non-#N/A value in each of these columns.
  //
  // Source: Report 1145 mandatory columns per FutureLog spec:
  //   - Article no.          (itemNo)
  //   - Order unit (OU)      (ou)
  //   - Content unit (CU)    (cu)
  //   - Article lead time    (leadTimeRaw — the raw source-file value, not the post-processed availability)
  const mandatoryKeys = ["itemNo", "ou", "cu", "leadTimeRaw"];

  // Track blanks and #N/A separately so the user can see exactly which cells to fix.
  const blankDetails = []; // [{row, col}]
  const naDetails = [];    // [{row, col}]

  rows.forEach((r, idx) => {
    for (const k of mandatoryKeys) {
      if (isNA(r[k])) {
        markCell(idx, k);
        // Preview shows the computed `availability` column, not `leadTimeRaw`.
        // Mark both so the row visibly highlights in the preview.
        if (k === "leadTimeRaw") markCell(idx, "availability");
        naDetails.push({ row: idx + 2, col: KEY_LABELS[k] || k });
      } else if (isBlank(r[k])) {
        markCell(idx, k);
        if (k === "leadTimeRaw") markCell(idx, "availability");
        blankDetails.push({ row: idx + 2, col: KEY_LABELS[k] || k });
      }
    }
  });

  if (naDetails.length) {
    const sample = naDetails
      .slice(0, 15)
      .map((d) => `  Row ${d.row}, column "${d.col}"`)
      .join("\n");
    const more = naDetails.length > 15 ? `\n  …and ${naDetails.length - 15} more` : "";
    errors.push(
      `Found #N/A in ${naDetails.length} mandatory cell${naDetails.length === 1 ? "" : "s"}. ` +
        `Please replace every #N/A with the correct value before generating:\n${sample}${more}`
    );
  }

  if (blankDetails.length) {
    const sample = blankDetails
      .slice(0, 15)
      .map((d) => `  Row ${d.row}, column "${d.col}"`)
      .join("\n");
    const more = blankDetails.length > 15 ? `\n  …and ${blankDetails.length - 15} more` : "";
    errors.push(
      `Found ${blankDetails.length} blank mandatory cell${blankDetails.length === 1 ? "" : "s"}. ` +
        `All mandatory columns must be filled:\n${sample}${more}`
    );
  }

  // Step 3: Triplet duplicate / EAN conflict / ItemNo-EAN mismatch
  const tripletSeen = new Map(); // key -> firstRowIdx
  const eanToItem = new Map(); // ean -> itemNo
  const itemToEan = new Map(); // itemNo -> ean
  const itemNoFirstSeen = new Map(); // itemNo -> rowIdx (for "duplicate without customerID")
  const duplicateTriplets = [];
  const eanConflicts = [];
  const itemEanMismatches = [];

  rows.forEach((r, idx) => {
    const itemNo = String(r.itemNo || "").trim();
    const cust = String(r.customerId || "").trim();
    const ean = String(r.ean || "").trim();

    if (itemNo !== "") {
      if (itemNoFirstSeen.has(itemNo)) {
        if (cust === "") {
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
          errors.push(`Row ${idx + 2}: Item No "${itemNo}" repeats without a Customer ID.`);
        }
      } else {
        itemNoFirstSeen.set(itemNo, idx);
      }
    }

    if (itemNo !== "" && cust !== "" && ean !== "" && ean !== "0000000000000") {
      const key = `${itemNo}|${cust}|${ean}`;
      if (tripletSeen.has(key)) {
        const firstIdx = tripletSeen.get(key);
        markCell(firstIdx, "itemNo"); markCell(firstIdx, "ean"); markCell(firstIdx, "customerId");
        markCell(idx, "itemNo"); markCell(idx, "ean"); markCell(idx, "customerId");
        duplicateTriplets.push(`Row ${idx + 2}: ItemNo ${itemNo}, CustID ${cust}, EAN ${ean}`);
      } else {
        tripletSeen.set(key, idx);
      }

      if (eanToItem.has(ean)) {
        if (eanToItem.get(ean) !== itemNo) {
          markCell(idx, "ean");
          eanConflicts.push(
            `Row ${idx + 2}: EAN ${ean} used with different Item No (${itemNo} vs ${eanToItem.get(ean)})`
          );
        }
      } else {
        eanToItem.set(ean, itemNo);
      }

      if (itemToEan.has(itemNo)) {
        if (itemToEan.get(itemNo) !== ean) {
          markCell(idx, "itemNo");
          itemEanMismatches.push(
            `Row ${idx + 2}: Item No ${itemNo} used with different EAN (${ean} vs ${itemToEan.get(itemNo)})`
          );
        }
      } else {
        itemToEan.set(itemNo, ean);
      }
    }
  });

  if (duplicateTriplets.length) {
    errors.push("Duplicate ItemNo + CustomerID + EAN combinations:\n  " + duplicateTriplets.join("\n  "));
  }
  if (eanConflicts.length) {
    errors.push("Same EAN used with different Item Nos:\n  " + eanConflicts.join("\n  "));
  }
  if (itemEanMismatches.length) {
    errors.push("Same Item No used with different EANs:\n  " + itemEanMismatches.join("\n  "));
  }

  // Step 4: Unit List / Country List
  const badUnits = [];
  rows.forEach((r, idx) => {
    const ou = String(r.ou || "").trim();
    const cu = String(r.cu || "").trim();
    if (ou !== "" && !VALID_UNITS.has(ou)) {
      markCell(idx, "ou");
      badUnits.push(`Row ${idx + 2}: OU "${ou}"`);
    }
    if (cu !== "" && !VALID_UNITS.has(cu)) {
      markCell(idx, "cu");
      badUnits.push(`Row ${idx + 2}: CU "${cu}"`);
    }
  });
  if (badUnits.length) {
    errors.push("Invalid unit codes (not in Unit_List):\n  " + badUnits.slice(0, 50).join("\n  "));
  }

  const badCountries = [];
  rows.forEach((r, idx) => {
    const c = String(r.origin || "").trim();
    if (c !== "" && !VALID_COUNTRIES.has(c)) {
      markCell(idx, "origin");
      badCountries.push(`Row ${idx + 2}: "${c}"`);
    }
  });
  if (badCountries.length) {
    errors.push("Invalid country codes (not in Country_List):\n  " + badCountries.slice(0, 50).join("\n  "));
  }

  // Warning: all availability == 0
  const allAvailZero = rows.every((r) => {
    const v = r.availability;
    return v === 0 || v === "0" || v === "";
  });
  if (allAvailZero) {
    warnings.push('All "availability" (column S) values are 0 — please double-check the prices in Report 1145.');
  }

  return { errors, warnings, invalidCells };
}
