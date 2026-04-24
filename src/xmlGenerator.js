// Direct port of the VBA generate_xml writer.
// Matches the tag structure and field order of the original output.

import { NA_MARKER } from "./reportParser.js";

// Replaces doubled spaces with a single space, and "&" with "+".
// Matches VBA ReplaceAll: collapse " " doubles, replace Chr(38) "&" with "+", trim.
// Also strips our internal NA_MARKER sentinel so it never reaches the XML.
function replaceAll(v) {
  if (v === null || v === undefined) return "";
  if (v === NA_MARKER) return "";
  let s = typeof v === "number" ? String(v) : String(v);
  // Collapse repeated spaces
  let prev;
  do {
    prev = s;
    s = s.replace(/  /g, " ");
  } while (s !== prev);
  s = s.replace(/&/g, "+");
  return s.trim();
}

// XML-escape AFTER the VBA transforms (for characters VBA would have written raw).
// The original VBA writes raw text into the XML stream (with "&" replaced by "+"),
// so the only characters left that need escaping are `<` and `>` (and stray `"` in attributes,
// but we never write those into attributes). We escape them defensively here so malformed
// product names can't break the XML.
function xmlEscape(v) {
  return replaceAll(v)
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Strip dots for OFFSTART / OFFEND — VBA: Replace(val, ".", "", 1, -1, 1)
function stripDots(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\./g, "");
}

/**
 * @param {Array<Object>} rows  Tabelle1-shaped rows
 * @param {{companyId:string, supplierNo:string, language:string, validityDate:string}} params
 * @returns {{xml:string, filename:string}}
 */
export function generateXml(rows, params) {
  const varLangName = String(params.language || "").toUpperCase();
  const isPrimaryLang = /^(DE|GB|FR|IT)$/.test(varLangName);

  const parts = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" ?>');
  parts.push("<FUTURELOG>");
  parts.push("<HEAD>");
  parts.push(`<VALIDITY>${params.validityDate}</VALIDITY>`);
  parts.push("</HEAD>");
  parts.push("<ARTICLES>");

  for (const r of rows) {
    parts.push("<ARTICLE>");

    // The language-name block. Original VBA uses <NAME>...</NAME> — an earlier
    // version of this file had a typo (<n>...</n>) which was visible in the
    // downloaded XML compared to VBA output. Keep this exactly as NAME.
    parts.push("<NAME>");
    parts.push(`<DE>${xmlEscape(r.descDE)}</DE>`);
    parts.push(`<FR>${xmlEscape(r.descFR)}</FR>`);
    parts.push(`<IT>${xmlEscape(r.descIT)}</IT>`);
    parts.push(`<GB>${xmlEscape(r.descGB)}</GB>`);
    if (!isPrimaryLang) {
      parts.push(`<XX>${xmlEscape(r.descExtra)}</XX>`);
    }
    parts.push("</NAME>");

    parts.push("<ARTICLEDATA>");
    parts.push(`<ARTNO>${xmlEscape(r.itemNo)}</ARTNO>`);
    parts.push(`<EAN>${xmlEscape(r.ean)}</EAN>`);
    parts.push(`<CUSTNO>${xmlEscape(r.customsNo)}</CUSTNO>`); // VBA: Cells(Z, 18) = customs no.
    parts.push(`<MANARTNO>${xmlEscape(r.manArtId)}</MANARTNO>`);
    parts.push(`<ORG>${xmlEscape(r.origin)}</ORG>`);
    parts.push(`<PICURL>${xmlEscape(r.specUrl)}</PICURL>`);
    parts.push("</ARTICLEDATA>");

    parts.push("<PRICES>");
    parts.push("<PRICE>");
    const cust = String(r.customerId || "").trim();
    parts.push(`<CUSTID>${cust === "" ? "0000" : xmlEscape(cust)}</CUSTID>`);
    parts.push(`<PRCOU>${xmlEscape(r.priceOU)}</PRCOU>`);
    parts.push(`<OU>${xmlEscape(r.ou)}</OU>`);
    parts.push(`<CU>${xmlEscape(r.cu)}</CU>`);
    parts.push(`<NUCUOU>${xmlEscape(r.cuou)}</NUCUOU>`);
    parts.push(`<VLZ>${xmlEscape(r.availability)}</VLZ>`);
    parts.push(`<OFFSTART>${stripDots(r.offerStart)}</OFFSTART>`);
    parts.push(`<OFFEND>${stripDots(r.offerEnd)}</OFFEND>`);
    parts.push("</PRICE>");
    parts.push("</PRICES>");

    parts.push("</ARTICLE>");
  }

  parts.push("</ARTICLES>");
  parts.push("</FUTURELOG>");

  const xml = parts.join("");

  // Filename: {Company}{Supplier}{YYYY}{MM}{DD}.cat.xml
  // VBA: varDateNrYear = Right(varDateNr,4); varDateNrDay = Left(varDateNr,2); varDateNrMonth = Mid(varDateNr,3,2)
  // So the YYYYMMDD portion in the filename is YYYY + MM + DD
  const dd = params.validityDate.slice(0, 2);
  const mm = params.validityDate.slice(2, 4);
  const yyyy = params.validityDate.slice(4, 8);
  const filename = `${params.companyId}${params.supplierNo}${yyyy}${mm}${dd}.cat.xml`;

  return { xml, filename };
}
