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

// List every unique row number, sorted ascending, with consecutive runs
// collapsed into ranges. Examples:
//   [5, 6, 7, 8, 9, 12, 14, 15, 16]  →  "5-9, 12, 14-16"
//   [8, 9]                            →  "8, 9"  (run of 2 stays as-is for clarity)
//   [5, 6, 7]                         →  "5-7"
// Rationale: real-world bulk warnings (e.g. "1908 rows have Scaled price
// blank") used to print every single row number, making the message unusable.
// Ranges preserve all the information without the wall of text.
function fmtRows(rowNums) {
  const sorted = [...new Set(rowNums)].sort((a, b) => a - b);
  if (sorted.length === 0) return "";

  const parts = [];
  let runStart = sorted[0];
  let runEnd = sorted[0];

  const flushRun = () => {
    if (runStart === runEnd) {
      parts.push(String(runStart));
    } else if (runEnd === runStart + 1) {
      // Run of exactly 2 → list both rather than "5-6" (slightly clearer in
      // small messages and only one extra char).
      parts.push(`${runStart}, ${runEnd}`);
    } else {
      parts.push(`${runStart}-${runEnd}`);
    }
  };

  for (let i = 1; i < sorted.length; i++) {
    const n = sorted[i];
    if (n === runEnd + 1) {
      runEnd = n;
    } else {
      flushRun();
      runStart = n;
      runEnd = n;
    }
  }
  flushRun();

  return parts.join(", ");
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
  const itemEanFirstSeen = new Map();
  const itemNoFirstSeen = new Map();
  const itemCustSeen = new Map();

  const itemCustDupRows = new Map();      // itemNo -> [rowNums] (duplicates only)
  const tripletDupByItem = new Map();     // itemNo -> [rowNums] (duplicates only)
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

      // NOTE: Previously this block flagged "same GTIN used with a different
      // Article no." as a blocking error. The business confirmed that GTINs
      // can legitimately be shared across different Article nos. (e.g. the
      // same product sold under multiple article codes), so that rule was
      // removed. The mirror check below — "same Article no. used with a
      // different GTIN" — is preserved because Article nos. shouldn't
      // duplicate in the first place per the existing duplicate-Article
      // rule, but kept here as a defensive belt-and-braces guard.

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

  if (itemEanConflictGroups.size > 0) {
    const total = Array.from(itemEanConflictGroups.values()).reduce((s, g) => s + g.conflictRows.length, 0);
    const groups = Array.from(itemEanConflictGroups.entries()).map(([itemNo, g]) => ({
      key: `Item "${itemNo}"`,
      rows: g.conflictRows,
    }));
    errors.push(formatGrouped(total, "Article no. used with different EANs", groups));
  }

  // ========== GTIN / EAN format validation ==========
  //
  // Once a GTIN is present (i.e. not the placeholder "0000000000000" and not
  // blank), it must be a valid EAN-13:
  //   1. Exactly 13 characters
  //   2. All digits, no letters or whitespace
  //   3. The 13th digit must equal the check digit computed from digits 1-12
  //
  // We bucket the failures by error type so the user sees a single grouped
  // message per failure mode (length, non-digit, check-digit) instead of one
  // line per row. Each group lists which Excel rows are affected.

  const gtinByLengthIssue = new Map();   // gtin string -> [rowNums]
  const gtinByDigitsIssue = new Map();   // gtin string -> [rowNums]
  const gtinByCheckIssue  = new Map();   // gtin string -> [rowNums]

  // Compute the EAN-13 check digit from a 12-digit prefix.
  // Standard formula: weight digits at odd positions (1,3,5,...) by 1, digits
  // at even positions (2,4,6,...) by 3. Check digit = (10 - sum mod 10) mod 10.
  function computeEan13CheckDigit(twelveDigits) {
    let sum = 0;
    for (let i = 0; i < 12; i++) {
      const d = twelveDigits.charCodeAt(i) - 48;
      sum += i % 2 === 0 ? d : d * 3;
    }
    return (10 - (sum % 10)) % 10;
  }

  rows.forEach((r, idx) => {
    const ean = String(r.ean || "").trim();
    // Skip placeholder GTINs — these are "no GTIN" by convention. Cases:
    //   - blank / empty cell
    //   - all-zeros of any length ("0000000000000")
    //   - 13 chars where the leading 12 are zeros ("0000000000001", "0000000000007")
    //   - 13 identical digits ("1111111111111", "9999999999999")
    //
    // Real GTINs always start with a non-zero GS1 issuer prefix (country code)
    // and never use the same digit thirteen times, so any of these patterns is
    // unambiguously a hand-typed sentinel meaning "we don't have a real GTIN."
    // Flagging them as "check digit invalid" is mathematically correct but
    // misleading to the user, so we suppress those errors entirely here.
    if (ean === "") return;
    if (/^0+$/.test(ean)) return;
    if (ean.length === 13 && /^0{12}/.test(ean)) return;
    if (ean.length === 13 && /^(\d)\1{12}$/.test(ean)) return;

    if (ean.length !== 13) {
      markCell(idx, "ean");
      const key = `"${ean}" (${ean.length} chars)`;
      if (!gtinByLengthIssue.has(key)) gtinByLengthIssue.set(key, []);
      gtinByLengthIssue.get(key).push(excelRow(idx));
      return;  // Stop further checks on this GTIN — wrong length is the root cause
    }

    if (!/^\d{13}$/.test(ean)) {
      markCell(idx, "ean");
      const key = `"${ean}"`;
      if (!gtinByDigitsIssue.has(key)) gtinByDigitsIssue.set(key, []);
      gtinByDigitsIssue.get(key).push(excelRow(idx));
      return;
    }

    const expected = computeEan13CheckDigit(ean.slice(0, 12));
    const actual = ean.charCodeAt(12) - 48;
    if (expected !== actual) {
      markCell(idx, "ean");
      const key = `"${ean}" (expected check digit ${expected}, got ${actual})`;
      if (!gtinByCheckIssue.has(key)) gtinByCheckIssue.set(key, []);
      gtinByCheckIssue.get(key).push(excelRow(idx));
    }
  });

  if (gtinByLengthIssue.size > 0) {
    const total = Array.from(gtinByLengthIssue.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(gtinByLengthIssue.entries()).map(([key, rs]) => ({ key, rows: rs }));
    errors.push(formatGrouped(total, "GTIN must be exactly 13 digits", groups));
  }
  if (gtinByDigitsIssue.size > 0) {
    const total = Array.from(gtinByDigitsIssue.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(gtinByDigitsIssue.entries()).map(([key, rs]) => ({ key, rows: rs }));
    errors.push(formatGrouped(total, "GTIN contains non-digit characters", groups));
  }
  if (gtinByCheckIssue.size > 0) {
    const total = Array.from(gtinByCheckIssue.values()).reduce((s, rs) => s + rs.length, 0);
    const groups = Array.from(gtinByCheckIssue.entries()).map(([key, rs]) => ({ key, rows: rs }));
    errors.push(formatGrouped(total, "GTIN check digit is invalid (likely a typo)", groups));
  }

  // ========== Field-length limits ==========
  //
  // Hard caps imposed by FutureLog / MDM:
  //   - Article no. → max 20 characters
  //   - Each description field (DE/FR/IT/GB/local) → max 100 characters
  //
  // The description limit applies INDEPENDENTLY to each language column. A
  // row could pass on descGB but fail on descExtra. We track each language
  // separately so the user knows which column has the long text — but at the
  // user's request, the offending Article no. itself is no longer printed
  // (just the row numbers).
  const ARTICLE_NO_MAX = 20;
  const DESCRIPTION_MAX = 100;
  const DESC_FIELDS = [
    { key: "descDE",    label: "German" },
    { key: "descFR",    label: "French" },
    { key: "descIT",    label: "Italian" },
    { key: "descGB",    label: "English" },
    { key: "descExtra", label: "Local" },
  ];

  // Flat row lists — no per-Article grouping, since users prefer not to see
  // the long item identifier echoed in the error message.
  const articleNoTooLongRows = [];
  // language label → flat row list (description errors still split by column)
  const descTooLongByLang = new Map();

  rows.forEach((r, idx) => {
    const itemNo = String(r.itemNo || "").trim();

    if (itemNo.length > ARTICLE_NO_MAX) {
      markCell(idx, "itemNo");
      articleNoTooLongRows.push(excelRow(idx));
    }

    for (const f of DESC_FIELDS) {
      const v = r[f.key];
      if (typeof v !== "string" || v === NA_MARKER) continue;
      if (v.length > DESCRIPTION_MAX) {
        markCell(idx, f.key);
        if (!descTooLongByLang.has(f.label)) descTooLongByLang.set(f.label, []);
        descTooLongByLang.get(f.label).push(excelRow(idx));
      }
    }
  });

  if (articleNoTooLongRows.length > 0) {
    errors.push(
      `${articleNoTooLongRows.length} Article no. exceeds ${ARTICLE_NO_MAX}-character limit — rows ${fmtRows(articleNoTooLongRows)}`
    );
  }
  if (descTooLongByLang.size > 0) {
    const total = Array.from(descTooLongByLang.values()).reduce((s, rs) => s + rs.length, 0);
    if (descTooLongByLang.size === 1) {
      // Single language → inline
      const [lang, rs] = [...descTooLongByLang.entries()][0];
      errors.push(
        `${total} Item name (${lang}) exceeds ${DESCRIPTION_MAX}-character limit — rows ${fmtRows(rs)}`
      );
    } else {
      // Multiple languages → one line per language
      const body = Array.from(descTooLongByLang.entries())
        .map(([lang, rs]) => `  ${lang} — rows ${fmtRows(rs)}`)
        .join("\n");
      errors.push(`${total} Item name exceeds ${DESCRIPTION_MAX}-character limit:\n${body}`);
    }
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

  // ========== Scaled-price fallback warnings ==========
  //
  // The parser handles two scenarios silently:
  //   - Scaled price = 0      → keep the 0 as price, lead time = 0
  //   - Scaled price = blank  → fall back to Price per order unit, lead time = 0
  //
  // Both behaviors are correct, but invisible to the user. Surface them as
  // warnings so the user can confirm they're intentional. Rows mentioned here
  // are EXCLUDED from the generic "price = 0" warning below — that warning is
  // only useful when nothing else explains the zero price.
  const scaledZeroRows = [];
  const scaledBlankRows = [];
  const scaledExplainedSet = new Set();   // row numbers covered by a scaled warning
  rows.forEach((r, idx) => {
    if (r.__scaledPriceWasZero) {
      scaledZeroRows.push(excelRow(idx));
      scaledExplainedSet.add(excelRow(idx));
    }
    if (r.__scaledPriceWasBlank) {
      scaledBlankRows.push(excelRow(idx));
      scaledExplainedSet.add(excelRow(idx));
    }
  });

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
        // Skip rows where a scaled-price warning will already explain the 0.
        if (!scaledExplainedSet.has(excelRow(idx))) {
          zeroPriceRows.push(excelRow(idx));
        }
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
      `${zeroPriceRows.length} row(s) have price = 0 (rows: ${fmtRows(zeroPriceRows)}). Check if a price column is missing.`
    );
  }

  if (scaledZeroRows.length) {
    warnings.push(
      `${scaledZeroRows.length} row(s) have Scaled price = 0. ` +
      `Kept the price as 0 and lead time = 0 (rows: ${fmtRows(scaledZeroRows)}).`
    );
  }
  if (scaledBlankRows.length) {
    warnings.push(
      `${scaledBlankRows.length} row(s) have Scaled price blank. ` +
      `Used Price per order unit instead and lead time = 0 (rows: ${fmtRows(scaledBlankRows)}).`
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
