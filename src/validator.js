// Port of the VBA validation logic in generate_xml / Convert_to_UNI
// Returns { errors: [...], warnings: [...], invalidCells: Map<rowIdx,Set<col>> }

import { VALID_UNITS, VALID_COUNTRIES, VALID_LANGUAGES } from "./referenceData.js";
import { NA_MARKER } from "./reportParser.js";

// Convert 0-based row index to the Excel row number the user sees.
// Report 1145 layout: rows 1-3 metadata, row 4 headers, row 5+ data.
const DATA_ROW_OFFSET = 5;
function excelRow(idx) {
  return idx + DATA_ROW_OFFSET;
}

// Human-readable labels used in error messages
const KEY_LABELS = {
  pos: "Pos",
  descDE: "German description",
  descFR: "French description",
  descIT: "Italian description",
  descGB: "English description",
  descExtra: "Local description",
  itemNo: "Article no.",
  ean: "EAN",
  manArtId: "Mfg Item No",
  ou: "Order unit (OU)",
  cu: "Content unit (CU)",
  cuou: "CU per OU",
  priceOU: "Price",
  origin: "Origin",
  customsNo: "Customs No.",
  availability: "Availability",
  leadTimeRaw: "Lead time",
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

// Compact list rendering for error bodies. Keeps messages short:
// shows up to `limit` items, then "and N more".
function compact(items, limit = 10) {
  if (items.length <= limit) return items.join("\n");
  return items.slice(0, limit).join("\n") + `\n…and ${items.length - limit} more`;
}

export function validate(rows, params) {
  const errors = [];
  const warnings = [];
  const invalidCells = new Map();

  const markCell = (rowIdx, colKey) => {
    if (!invalidCells.has(rowIdx)) invalidCells.set(rowIdx, new Set());
    invalidCells.get(rowIdx).add(colKey);
  };

  if (rows.length === 0) {
    errors.push("No data rows found in the uploaded file.");
    return { errors, warnings, invalidCells };
  }

  // ---------- Parameter checks ----------
  if (!/^\d{3}$/.test(params.companyId || "")) {
    errors.push("Company ID must be 3 digits.");
  }
  if (!/^\d{6}$/.test(params.supplierNo || "")) {
    errors.push("Supplier Number must be 6 digits.");
  }
  if (!VALID_LANGUAGES.includes(String(params.language || "").toUpperCase())) {
    errors.push(`Language must be one of: ${VALID_LANGUAGES.join(", ")}.`);
  }
  if (!/^\d{8}$/.test(params.validityDate || "")) {
    errors.push("Validity Date must be 8 digits (DDMMYYYY).");
  }

  // ---------- Customer ID "0000" core-entry rule ----------
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
      const list = compact(missing, 10);
      errors.push(`Missing Customer ID "0000" for ${missing.length} item(s):\n${list}`);
      rows.forEach((r, idx) => {
        if (missing.includes(String(r.itemNo || "").trim())) {
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
        }
      });
    }
  }

  // ---------- Mandatory fields ----------
  const mandatoryKeys = ["itemNo", "ou", "cu", "leadTimeRaw"];

  const blankDetails = [];
  const naDetails = [];

  rows.forEach((r, idx) => {
    for (const k of mandatoryKeys) {
      if (isNA(r[k])) {
        markCell(idx, k);
        if (k === "leadTimeRaw") markCell(idx, "availability");
        naDetails.push(`  Row ${excelRow(idx)}: ${KEY_LABELS[k] || k}`);
      } else if (isBlank(r[k])) {
        markCell(idx, k);
        if (k === "leadTimeRaw") markCell(idx, "availability");
        blankDetails.push(`  Row ${excelRow(idx)}: ${KEY_LABELS[k] || k}`);
      }
    }
  });

  if (naDetails.length) {
    errors.push(
      `${naDetails.length} cell(s) contain #N/A — please replace with a real value:\n${compact(naDetails, 10)}`
    );
  }
  if (blankDetails.length) {
    errors.push(
      `${blankDetails.length} required cell(s) are blank:\n${compact(blankDetails, 10)}`
    );
  }

  // ---------- Duplicate / conflict checks ----------
  const tripletSeen = new Map();
  const eanToItem = new Map();
  const itemToEan = new Map();
  const itemNoFirstSeen = new Map();
  const itemCustSeen = new Map();
  const duplicateTriplets = [];
  const eanConflicts = [];
  const itemEanMismatches = [];
  const itemCustDuplicates = [];
  const itemNoNoCust = [];

  rows.forEach((r, idx) => {
    const itemNo = String(r.itemNo || "").trim();
    const cust = String(r.customerId || "").trim();
    const ean = String(r.ean || "").trim();

    if (itemNo !== "") {
      if (itemNoFirstSeen.has(itemNo)) {
        if (cust === "") {
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
          itemNoNoCust.push(`  Row ${excelRow(idx)}: Item "${itemNo}" repeats without Customer ID`);
        }
      } else {
        itemNoFirstSeen.set(itemNo, idx);
      }

      // Article no. must be unique per Customer ID
      if (cust !== "") {
        const itemCustKey = `${itemNo}|${cust}`;
        if (itemCustSeen.has(itemCustKey)) {
          const firstIdx = itemCustSeen.get(itemCustKey);
          markCell(firstIdx, "itemNo");
          markCell(firstIdx, "customerId");
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
          itemCustDuplicates.push(
            `  Row ${excelRow(idx)}: Item "${itemNo}" + Cust "${cust}" (also row ${excelRow(firstIdx)})`
          );
        } else {
          itemCustSeen.set(itemCustKey, idx);
        }
      }
    }

    if (itemNo !== "" && cust !== "" && ean !== "" && ean !== "0000000000000") {
      const key = `${itemNo}|${cust}|${ean}`;
      if (tripletSeen.has(key)) {
        const firstIdx = tripletSeen.get(key);
        markCell(firstIdx, "itemNo"); markCell(firstIdx, "ean"); markCell(firstIdx, "customerId");
        markCell(idx, "itemNo"); markCell(idx, "ean"); markCell(idx, "customerId");
        duplicateTriplets.push(
          `  Row ${excelRow(idx)}: Item ${itemNo} / Cust ${cust} / EAN ${ean}`
        );
      } else {
        tripletSeen.set(key, idx);
      }

      if (eanToItem.has(ean)) {
        if (eanToItem.get(ean) !== itemNo) {
          markCell(idx, "ean");
          eanConflicts.push(
            `  Row ${excelRow(idx)}: EAN ${ean} → items ${itemNo} vs ${eanToItem.get(ean)}`
          );
        }
      } else {
        eanToItem.set(ean, itemNo);
      }

      if (itemToEan.has(itemNo)) {
        if (itemToEan.get(itemNo) !== ean) {
          markCell(idx, "itemNo");
          itemEanMismatches.push(
            `  Row ${excelRow(idx)}: Item ${itemNo} → EANs ${ean} vs ${itemToEan.get(itemNo)}`
          );
        }
      } else {
        itemToEan.set(itemNo, ean);
      }
    }
  });

  if (itemNoNoCust.length) {
    errors.push(
      `${itemNoNoCust.length} item(s) repeat without a Customer ID:\n${compact(itemNoNoCust, 10)}`
    );
  }
  if (itemCustDuplicates.length) {
    errors.push(
      `${itemCustDuplicates.length} duplicate Article no. + Customer ID:\n${compact(itemCustDuplicates, 10)}`
    );
  }
  if (duplicateTriplets.length) {
    errors.push(
      `${duplicateTriplets.length} duplicate Item + Cust + EAN row(s):\n${compact(duplicateTriplets, 10)}`
    );
  }
  if (eanConflicts.length) {
    errors.push(
      `Same EAN used with different items:\n${compact(eanConflicts, 10)}`
    );
  }
  if (itemEanMismatches.length) {
    errors.push(
      `Same Item No used with different EANs:\n${compact(itemEanMismatches, 10)}`
    );
  }

  // ---------- Unit / Country lists ----------
  const badUnits = [];
  rows.forEach((r, idx) => {
    const ou = String(r.ou || "").trim();
    const cu = String(r.cu || "").trim();
    if (ou !== "" && !VALID_UNITS.has(ou)) {
      markCell(idx, "ou");
      badUnits.push(`  Row ${excelRow(idx)}: OU "${ou}"`);
    }
    if (cu !== "" && !VALID_UNITS.has(cu)) {
      markCell(idx, "cu");
      badUnits.push(`  Row ${excelRow(idx)}: CU "${cu}"`);
    }
  });
  if (badUnits.length) {
    errors.push(`Invalid unit code(s):\n${compact(badUnits, 10)}`);
  }

  const badCountries = [];
  rows.forEach((r, idx) => {
    const c = String(r.origin || "").trim();
    if (c !== "" && !VALID_COUNTRIES.has(c)) {
      markCell(idx, "origin");
      badCountries.push(`  Row ${excelRow(idx)}: "${c}"`);
    }
  });
  if (badCountries.length) {
    errors.push(`Invalid country code(s):\n${compact(badCountries, 10)}`);
  }

  // ---------- Price checks ----------
  const negativePrices = [];
  const zeroPriceRows = [];
  rows.forEach((r, idx) => {
    const p = r.priceOU;
    if (typeof p === "number") {
      if (p < 0) {
        markCell(idx, "priceOU");
        negativePrices.push(`  Row ${excelRow(idx)}: Item "${r.itemNo}" = ${p}`);
      } else if (p === 0) {
        zeroPriceRows.push(excelRow(idx));
      }
    }
  });

  if (negativePrices.length) {
    errors.push(
      `${negativePrices.length} negative price(s) — not allowed:\n${compact(negativePrices, 10)}`
    );
  }

  if (zeroPriceRows.length) {
    const sample = zeroPriceRows.slice(0, 10).join(", ");
    const more = zeroPriceRows.length > 10 ? `, +${zeroPriceRows.length - 10} more` : "";
    warnings.push(
      `${zeroPriceRows.length} row(s) have price = 0 (row${zeroPriceRows.length === 1 ? "" : "s"}: ${sample}${more}). Check if a price column is missing.`
    );
  }

  // Warning: all availability == 0
  const allAvailZero = rows.every((r) => {
    const v = r.availability;
    return v === 0 || v === "0" || v === "";
  });
  if (allAvailZero) {
    warnings.push("All availability values are 0 — please double-check prices.");
  }

  return { errors, warnings, invalidCells };
}
