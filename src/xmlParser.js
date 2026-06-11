// FUTURELOG XML → row objects for the "Convert XML to Report 1145" tab.
//
// Input: a .cat.xml file with the schema produced by xmlGenerator.js:
//
//   <FUTURELOG>
//     <HEAD><VALIDITY>DDMMYYYY</VALIDITY></HEAD>
//     <ARTICLES>
//       <ARTICLE>
//         <NAME><DE/><FR/><IT/><GB/><XX/></NAME>
//         <ARTICLEDATA><ARTNO/><EAN/><CUSTNO/><MANARTNO/><ORG/><PICURL/></ARTICLEDATA>
//         <PRICES><PRICE>
//           <CUSTID/><PRCOU/><OU/><CU/><NUCUOU/><VLZ/>
//           <OFFSTART/><OFFEND/>
//         </PRICE></PRICES>
//       </ARTICLE>
//       …
//     </ARTICLES>
//   </FUTURELOG>
//
// Note: older files exported by this app may use <n> instead of <NAME> due to
// an earlier code bug. We accept both for backwards compatibility.
//
// Output: { validity, rows[] }
//   - validity: { dd, mm, yyyy, raw } from the HEAD/VALIDITY tag (for filename)
//   - rows: one entry per <ARTICLE>, with the same field keys the writer
//     expects (descDE, descFR, descIT, descGB, descExtra, itemNo, ean, etc.)
//
// The writer (r1145Writer.js) then maps these fields into the 22-column
// Tabelle1 layout for the .xlsx output.
//
// Implementation uses getElementsByTagName rather than querySelector — it's
// universally supported by XML parsers (including older browser DOMParser
// flavors and xmldom polyfills) and doesn't trip on XML namespaces.

const STRIP_DOTS_DATE_RE = /^(\d{2})(\d{2})(\d{4})$/;

/**
 * Parse FUTURELOG XML text into structured rows.
 *
 * @param {string} xmlText  Raw .xml file contents
 * @returns {{ validity: object, rows: object[] }}
 * @throws Error if the XML can't be parsed or has no <ARTICLE> blocks
 */
export function parseFuturelogXml(xmlText) {
  if (typeof xmlText !== "string" || xmlText.trim() === "") {
    throw new Error("XML file is empty.");
  }

  // Strip BOM if present — some generators (including ours) prepend U+FEFF so
  // tools detect UTF-8 encoding correctly. DOMParser tolerates it but cleaner
  // input keeps later regex/text checks predictable.
  let text = xmlText;
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  // Use the browser's DOMParser — no dependency. Returns a Document object;
  // any malformed XML produces a <parsererror> element we need to check for.
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, "application/xml");

  // Modern browsers emit <parsererror> for malformed XML. Use the generic
  // getElementsByTagName here too since some parser implementations don't
  // support querySelector.
  const perrList = doc.getElementsByTagName("parsererror");
  if (perrList && perrList.length > 0) {
    throw new Error(
      "This doesn't look like a valid XML file. " +
      "Make sure the file isn't truncated and starts with <?xml version=\"1.0\"…?>."
    );
  }

  const roots = doc.getElementsByTagName("FUTURELOG");
  if (roots.length === 0) {
    throw new Error(
      "Expected a <FUTURELOG> root element. This may not be a FUTURELOG export."
    );
  }
  const root = roots[0];

  // Extract validity date from HEAD/VALIDITY (used for the output filename)
  const validityText = textByPath(root, ["HEAD", "VALIDITY"]);
  const validity = parseValidity(validityText);

  // Find all article elements — direct ARTICLES > ARTICLE children
  const articlesContainer = firstByTag(root, "ARTICLES");
  if (!articlesContainer) {
    throw new Error("Expected an <ARTICLES> block inside <FUTURELOG>.");
  }
  const articleEls = articlesContainer.getElementsByTagName("ARTICLE");
  if (articleEls.length === 0) {
    throw new Error(
      "No <ARTICLE> blocks found inside <ARTICLES>. The XML appears to be empty."
    );
  }

  const rows = [];
  let pos = 1;
  for (let i = 0; i < articleEls.length; i++) {
    rows.push(extractRow(articleEls[i], pos++));
  }

  return { validity, rows };
}

// ─── helpers ────────────────────────────────────────────────────────────────

// Return the first child element with the given tag, or null.
function firstByTag(parent, tag) {
  const els = parent.getElementsByTagName(tag);
  return els.length > 0 ? els[0] : null;
}

// Walk a path of tags (e.g. ["HEAD", "VALIDITY"]) and return the trimmed text
// content of the deepest element, or "" if any step is missing.
function textByPath(parent, tagPath) {
  let cur = parent;
  for (const tag of tagPath) {
    cur = firstByTag(cur, tag);
    if (!cur) return "";
  }
  const t = cur.textContent;
  return t === null || t === undefined ? "" : t.trim();
}

// Inside an <ARTICLE>, look for <NAME> first (current schema). Fall back to
// <n> if not present — older files exported by this app used <n> due to a
// historical bug, and we don't want those to fail to import.
function nameElement(article) {
  return firstByTag(article, "NAME") || firstByTag(article, "n");
}

// Pull a child of <NAME> by tag, e.g. DE, FR, IT, GB, XX.
function nameField(article, tag) {
  const name = nameElement(article);
  if (!name) return "";
  const el = firstByTag(name, tag);
  return el ? (el.textContent || "").trim() : "";
}

function parseValidity(raw) {
  // The forward generator writes DDMMYYYY (e.g. "02122025"). We also accept
  // YYYYMMDD just in case some other tool produces the file. If the string
  // doesn't match either, we return raw so the user still sees something.
  if (!raw) return { dd: "", mm: "", yyyy: "", raw: "" };
  const m = String(raw).match(STRIP_DOTS_DATE_RE);
  if (m) {
    const [, dd, mm, yyyy] = m;
    // 8-digit could be DDMMYYYY or YYYYMMDD. Distinguish by the first 4 chars:
    // if they look like a year (>= 2000 and <= 2200), assume YYYYMMDD.
    const head = Number(raw.slice(0, 4));
    if (head >= 2000 && head <= 2200) {
      return { dd: raw.slice(6, 8), mm: raw.slice(4, 6), yyyy: raw.slice(0, 4), raw };
    }
    return { dd, mm, yyyy, raw };
  }
  return { dd: "", mm: "", yyyy: "", raw: String(raw) };
}

// PRICES > PRICE > <tag>
function priceField(article, tag) {
  const prices = firstByTag(article, "PRICES");
  if (!prices) return "";
  const price = firstByTag(prices, "PRICE");
  if (!price) return "";
  const el = firstByTag(price, tag);
  return el ? (el.textContent || "").trim() : "";
}

// ARTICLEDATA > <tag>
function articleDataField(article, tag) {
  const ad = firstByTag(article, "ARTICLEDATA");
  if (!ad) return "";
  const el = firstByTag(ad, tag);
  return el ? (el.textContent || "").trim() : "";
}

// Build one row object from an <ARTICLE> element.
// Field keys mirror the parser-output shape used by the Show Full Data modal
// so the existing UI components can render them without changes.
function extractRow(article, pos) {
  // Numeric coercion for PRCOU and NUCUOU — Excel sees them as numbers when
  // they're numeric, otherwise as strings.
  const prcouRaw = priceField(article, "PRCOU");
  const nucouRaw = priceField(article, "NUCUOU");
  const vlzRaw   = priceField(article, "VLZ");

  return {
    pos,
    // Description fields — one per language
    descDE:    nameField(article, "DE"),
    descFR:    nameField(article, "FR"),
    descIT:    nameField(article, "IT"),
    descGB:    nameField(article, "GB"),
    descExtra: nameField(article, "XX"),
    // Article identifiers
    itemNo:    articleDataField(article, "ARTNO"),
    ean:       articleDataField(article, "EAN"),
    manArtId:  articleDataField(article, "MANARTNO"),
    origin:    articleDataField(article, "ORG"),
    customsNo: articleDataField(article, "CUSTNO"),
    specUrl:   articleDataField(article, "PICURL"),
    // Price block
    customerId:   priceField(article, "CUSTID"),
    priceOU:      toNumberOrString(prcouRaw),
    ou:           priceField(article, "OU"),
    cu:           priceField(article, "CU"),
    cuou:         toNumberOrString(nucouRaw),
    availability: toNumberOrString(vlzRaw),
    offerStart:   priceField(article, "OFFSTART"),
    offerEnd:     priceField(article, "OFFEND"),
    // leadTimeRaw mirrors availability — the XML only carries VLZ (the final
    // value), not the raw source-file lead time, so they're identical here.
    leadTimeRaw:  toNumberOrString(vlzRaw),
  };
}

// Coerce a string to a number when it looks like one; otherwise return the
// trimmed string as-is. Empty strings stay empty.
function toNumberOrString(v) {
  if (v === null || v === undefined) return "";
  const s = String(v).trim();
  if (s === "") return "";
  // Match the same liberal number parsing the forward parser uses: strip
  // thousands separators (rare in our XML output but possible) before parsing.
  const cleaned = s.replace(/,/g, "");
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return s;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : s;
}
