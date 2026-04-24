// Port of the VBA validation logic in generate_xml / Convert_to_UNI.
// Errors sharing a common cause are collapsed into a single line listing
// every affected row — minimal identifier, all rows shown.

import { VALID_UNITS, VALID_COUNTRIES, VALID_LANGUAGES } from "./referenceData.js";
import { NA_MARKER } from "./reportParser.js";

// Convert 0-based row index to the Excel row number the user sees.
// Report 1145 layout: rows 1-3 metadata, row 4 headers, row 5+ data.
const DATA_ROW_OFFSET = 5;
function excelRow(idx) {
  return idx + DATA_ROW_OFFSET;
}

const KEY_LABELS = {
  itemNo: "Article no.",
  ean: "EAN",
  ou: "Order unit (OU)",
  cu: "Content unit (CU)",
  priceOU: "Price",
  origin: "Origin",
  availability: "Availability",
  leadTimeRaw: "Lead time",
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

// List every unique row number, sorted ascending — no truncation.
function fmtRows(rowNums) {
  return [...new Set(rowNums)].sort((a, b) => a - b).join(", ");
}

// Build a grouped-error message.
// groups: Array<{ key: string, rows: number[] }>
// One group → inline on one line. Multiple groups → one line per group.
function formatGrouped(count, title, groups) {
  if (groups.length === 1) {
    const g = groups[0];
    return `${count} ${title} : ${g.key} — rows ${fmtRows(g.rows)}`;
  }
  const body = groups.map((g) => `  ${g.key} — rows ${fmtRows(g.rows)}`).join("\n");
  return `${count} ${title}:\n${body}`;
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

  // ========== Parameter checks ==========
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

  // ========== Customer ID "0000" core-entry rule ==========
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
      errors.push(`${missing.length} item(s) missing Customer ID "0000" : ${missing.join(", ")}`);
      rows.forEach((r, idx) => {
        if (missing.includes(String(r.itemNo || "").trim())) {
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
        }
      });
    }
  }

  // ========== Mandatory fields (grouped by column) ==========
  const mandatoryKeys = ["itemNo", "ou", "cu", "leadTimeRaw"];
  const naByColumn = new Map();
  const blankByColumn = new Map();

  rows.forEach((r, idx) => {
    for (const k of mandatoryKeys) {
      if (isNA(r[k])) {
        markCell(idx, k);
        if (k === "leadTimeRaw") markCell(idx, "availability");
        if (!naByColumn.has(k)) naByColumn.set(k, []);
        naByColumn.get(k).push(excelRow(idx));
      } else if (isBlank(r[k])) {
        markCell(idx, k);
        if (k === "leadTimeRaw") markCell(idx, "availability");
        if (!blankByColumn.has(k)) blankByColumn.set(k, []);
        blankByColumn.get(k).push(excelRow(idx));
      }
    }
  });

  if (naByColumn.size > 0) {
    const total = Array.from(naByColumn.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(naByColumn.entries()).map(([k, rs]) => ({
      key: KEY_LABELS[k] || k,
      rows: rs,
    }));
    errors.push(formatGrouped(total, "#N/A cell(s)", groups));
  }

  if (blankByColumn.size > 0) {
    const total = Array.from(blankByColumn.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(blankByColumn.entries()).map(([k, rs]) => ({
      key: KEY_LABELS[k] || k,
      rows: rs,
    }));
    errors.push(formatGrouped(total, "blank required cell(s)", groups));
  }

  // ========== Duplicate / conflict checks ==========
  // Detection is still per (Item, Cust) pair — duplicates are allowed across
  // different Customer IDs — but the display is regrouped by Item only so the
  // error bubble shows one line per Article no. with every affected row.
  const tripletSeen = new Map();
  const eanFirstSeen = new Map();
  const itemEanFirstSeen = new Map();
  const itemNoFirstSeen = new Map();
  const itemCustSeen = new Map();

  const itemCustDupRows = new Map();      // itemNo -> [rowNums] (duplicates only)
  const tripletDupByItem = new Map();     // itemNo -> [rowNums] (duplicates only)
  const eanConflictGroups = new Map();    // ean -> { firstRow, conflictRows[] }
  const itemEanConflictGroups = new Map();// itemNo -> { firstRow, conflictRows[] }
  const itemNoRepeatNoCust = new Map();   // itemNo -> [rowNums]

  rows.forEach((r, idx) => {
    const itemNo = String(r.itemNo || "").trim();
    const cust = String(r.customerId || "").trim();
    const ean = String(r.ean || "").trim();

    if (itemNo !== "") {
      if (itemNoFirstSeen.has(itemNo)) {
        if (cust === "") {
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
          if (!itemNoRepeatNoCust.has(itemNo)) itemNoRepeatNoCust.set(itemNo, []);
          itemNoRepeatNoCust.get(itemNo).push(excelRow(idx));
        }
      } else {
        itemNoFirstSeen.set(itemNo, idx);
      }

      if (cust !== "") {
        const itemCustKey = `${itemNo}|${cust}`;
        if (itemCustSeen.has(itemCustKey)) {
          const firstIdx = itemCustSeen.get(itemCustKey);
          markCell(firstIdx, "itemNo");
          markCell(firstIdx, "customerId");
          markCell(idx, "itemNo");
          markCell(idx, "customerId");
          if (!itemCustDupRows.has(itemNo)) itemCustDupRows.set(itemNo, []);
          itemCustDupRows.get(itemNo).push(excelRow(idx));
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
        if (!tripletDupByItem.has(itemNo)) tripletDupByItem.set(itemNo, []);
        tripletDupByItem.get(itemNo).push(excelRow(idx));
      } else {
        tripletSeen.set(key, idx);
      }

      // EAN conflict — same EAN used with a different item
      if (eanFirstSeen.has(ean)) {
        const first = eanFirstSeen.get(ean);
        if (first.itemNo !== itemNo) {
          markCell(idx, "ean");
          if (!eanConflictGroups.has(ean)) {
            eanConflictGroups.set(ean, {
              firstRow: excelRow(first.firstIdx),
              conflictRows: [],
            });
          }
          eanConflictGroups.get(ean).conflictRows.push(excelRow(idx));
        }
      } else {
        eanFirstSeen.set(ean, { itemNo, firstIdx: idx });
      }

      // Item/EAN mismatch — same item used with a different EAN
      if (itemEanFirstSeen.has(itemNo)) {
        const first = itemEanFirstSeen.get(itemNo);
        if (first.ean !== ean) {
          markCell(idx, "itemNo");
          if (!itemEanConflictGroups.has(itemNo)) {
            itemEanConflictGroups.set(itemNo, {
              firstRow: excelRow(first.firstIdx),
              conflictRows: [],
            });
          }
          itemEanConflictGroups.get(itemNo).conflictRows.push(excelRow(idx));
        }
      } else {
        itemEanFirstSeen.set(itemNo, { ean, firstIdx: idx });
      }
    }
  });

  // --- Emit grouped duplicate errors ---

  if (itemNoRepeatNoCust.size > 0) {
    const total = Array.from(itemNoRepeatNoCust.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(itemNoRepeatNoCust.entries()).map(([itemNo, rs]) => ({
      key: `Item "${itemNo}"`,
      rows: rs,
    }));
    errors.push(formatGrouped(total, "Article no. repeats without Customer ID", groups));
  }

  if (itemCustDupRows.size > 0) {
    const total = Array.from(itemCustDupRows.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(itemCustDupRows.entries()).map(([itemNo, rs]) => ({
      key: `Item "${itemNo}"`,
      rows: rs,
    }));
    errors.push(formatGrouped(total, "duplicate Article no.", groups));
  }

  if (tripletDupByItem.size > 0) {
    const total = Array.from(tripletDupByItem.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(tripletDupByItem.entries()).map(([itemNo, rs]) => ({
      key: `Item "${itemNo}"`,
      rows: rs,
    }));
    errors.push(formatGrouped(total, "duplicate Article no. + EAN", groups));
  }

  if (eanConflictGroups.size > 0) {
    const total = Array.from(eanConflictGroups.values()).reduce((s, g) => s + g.conflictRows.length, 0);
    const groups = Array.from(eanConflictGroups.entries()).map(([ean, g]) => ({
      key: `EAN ${ean}`,
      rows: g.conflictRows,
    }));
    errors.push(formatGrouped(total, "EAN used with different Article no.", groups));
  }

  if (itemEanConflictGroups.size > 0) {
    const total = Array.from(itemEanConflictGroups.values()).reduce((s, g) => s + g.conflictRows.length, 0);
    const groups = Array.from(itemEanConflictGroups.entries()).map(([itemNo, g]) => ({
      key: `Item "${itemNo}"`,
      rows: g.conflictRows,
    }));
    errors.push(formatGrouped(total, "Article no. used with different EANs", groups));
  }

  // ========== Unit / Country lists ==========
  const badUnitsByCode = new Map();
  rows.forEach((r, idx) => {
    const ou = String(r.ou || "").trim();
    const cu = String(r.cu || "").trim();
    if (ou !== "" && !VALID_UNITS.has(ou)) {
      markCell(idx, "ou");
      const key = `OU "${ou}"`;
      if (!badUnitsByCode.has(key)) badUnitsByCode.set(key, []);
      badUnitsByCode.get(key).push(excelRow(idx));
    }
    if (cu !== "" && !VALID_UNITS.has(cu)) {
      markCell(idx, "cu");
      const key = `CU "${cu}"`;
      if (!badUnitsByCode.has(key)) badUnitsByCode.set(key, []);
      badUnitsByCode.get(key).push(excelRow(idx));
    }
  });
  if (badUnitsByCode.size > 0) {
    const total = Array.from(badUnitsByCode.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(badUnitsByCode.entries()).map(([key, rs]) => ({ key, rows: rs }));
    errors.push(formatGrouped(total, "invalid unit code(s)", groups));
  }

  const badCountriesByCode = new Map();
  rows.forEach((r, idx) => {
    const c = String(r.origin || "").trim();
    if (c !== "" && !VALID_COUNTRIES.has(c)) {
      markCell(idx, "origin");
      const key = `"${c}"`;
      if (!badCountriesByCode.has(key)) badCountriesByCode.set(key, []);
      badCountriesByCode.get(key).push(excelRow(idx));
    }
  });
  if (badCountriesByCode.size > 0) {
    const total = Array.from(badCountriesByCode.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(badCountriesByCode.entries()).map(([key, rs]) => ({ key, rows: rs }));
    errors.push(formatGrouped(total, "invalid country code(s)", groups));
  }

  // ========== Price checks ==========
  const negativePricesByItem = new Map();
  const zeroPriceRows = [];
  rows.forEach((r, idx) => {
    const p = r.priceOU;
    if (typeof p === "number") {
      if (p < 0) {
        markCell(idx, "priceOU");
        const key = String(r.itemNo || "").trim();
        if (!negativePricesByItem.has(key)) {
          negativePricesByItem.set(key, { price: p, rows: [] });
        }
        negativePricesByItem.get(key).rows.push(excelRow(idx));
      } else if (p === 0) {
        zeroPriceRows.push(excelRow(idx));
      }
    }
  });

  if (negativePricesByItem.size > 0) {
    const total = Array.from(negativePricesByItem.values()).reduce((s, g) => s + g.rows.length, 0);
    const groups = Array.from(negativePricesByItem.entries()).map(([itemNo, g]) => ({
      key: `Item "${itemNo}" = ${g.price}`,
      rows: g.rows,
    }));
    errors.push(formatGrouped(total, "negative price(s)", groups));
  }

  if (zeroPriceRows.length) {
    warnings.push(
      `${zeroPriceRows.length} row(s) have price = 0 (rows: ${zeroPriceRows.join(", ")}). Check if a price column is missing.`
    );
  }

  const allAvailZero = rows.every((r) => {
    const v = r.availability;
    return v === 0 || v === "0" || v === "";
  });
  if (allAvailZero) {
    warnings.push("All availability values are 0 — please double-check prices.");
  }

  return { errors, warnings, invalidCells };
}
