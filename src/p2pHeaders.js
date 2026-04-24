// P2P header definitions.
//
// Two data structures exported here:
//
//   P2P_HEADERS — per-language label map, used by the UI to render the
//                 "Show expected headers" popup. Hotels may prefer to fill in
//                 their P2P file using their local language, so we show them
//                 EN or VN strings depending on the picker.
//
//   P2P_ALIASES — field → [all known aliases across every language]. Used by
//                 the parser for header matching. This way the parser is
//                 language-agnostic: it accepts files with EN headers, VN
//                 headers, or a mix. First match wins.
//
// Vietnamese strings were verified against a real VN P2P file
// (P2p_Report_SONG_HANH_T3_2026.xlsx). Some suppliers use "Đơn vị Nội dung"
// (capital N) rather than "Đơn vị nội dung" — the header matcher lowercases
// everything so either works.

export const P2P_HEADERS = {
  EN: {
    wsNo:           "WS No.",
    itemName:       "Item Name",
    articleNo:      "Article No.",
    gtin:           "GTIN",
    orderUnit:      "Order Unit",
    contentUnits:   "Content Units",
    packagingUnit:  "Packaging unit",
    priceOrderUnit: "Price / Order unit",
    newPrice:       "NEW PRICE",
    minOrderQty:    "Minimum Order Quantity",
    originCountry:  "Country of origin",
  },
  VN: {
    wsNo:           "Mã WS.",
    itemName:       "Tên mặt hàng",
    articleNo:      "Mã sản phẩm",
    gtin:           "GTIN",
    orderUnit:      "Đơn vị đơn đặt hàng (OU)",
    contentUnits:   "Đơn vị Nội dung",
    packagingUnit:  "Đơn vị đóng gói",
    priceOrderUnit: "Đơn giá",
    // Hotels using VN usually don't ship a NEW PRICE column at all (they
    // overwrite the price/OU column instead). If they do, they use the EN
    // label — no Vietnamese translation exists in FutureLog's Language sheet.
    newPrice:       "NEW PRICE",
    minOrderQty:    "Số lượng đặt hàng tối thiểu",
    originCountry:  "Nguồn gốc xuất xứ",
  },
};

export const P2P_FIELD_KEYS = [
  "wsNo", "itemName", "articleNo", "gtin", "orderUnit", "contentUnits",
  "packagingUnit", "priceOrderUnit", "newPrice", "minOrderQty", "originCountry",
];

// Merged alias list — parser uses this to match headers regardless of language.
// Order matters for display in error messages but not for matching.
export const P2P_ALIASES = {
  wsNo:           ["WS No.", "Mã WS."],
  itemName:       ["Item Name", "Tên mặt hàng"],
  articleNo:      ["Article No.", "Mã sản phẩm"],
  gtin:           ["GTIN"],
  orderUnit:      ["Order Unit", "Đơn vị đơn đặt hàng (OU)"],
  contentUnits:   ["Content Units", "Đơn vị Nội dung"],
  packagingUnit:  ["Packaging unit", "Đơn vị đóng gói"],
  priceOrderUnit: ["Price / Order unit", "Đơn giá"],
  newPrice:       ["NEW PRICE"],
  minOrderQty:    ["Minimum Order Quantity", "Số lượng đặt hàng tối thiểu"],
  originCountry:  ["Country of origin", "Nguồn gốc xuất xứ"],
};

// Prefixes for the supplier / division label rows (rows 2 and 3 of the file).
// The hotel sends either English or localized versions of these labels.
// VN files use "Khách sạn" (literally "Hotel") where EN uses "Division".
export const SUPPLIER_LABEL_PREFIXES = ["Supplier", "Nhà cung cấp"];
export const DIVISION_LABEL_PREFIXES = ["Division", "Khách sạn"];

/**
 * Ordered list of header strings for display in the "expected headers" popup.
 * Required fields are flagged so the UI can highlight them. The `newPrice`
 * entry is omitted entirely when the toggle is OFF — in that case the parser
 * doesn't look for a NEW PRICE column at all, so showing it in the example
 * would just confuse users into thinking it's needed.
 */
export function headerDisplayList(lang, useNewPriceCol) {
  const hdrs = P2P_HEADERS[lang] || P2P_HEADERS.EN;
  return P2P_FIELD_KEYS
    .filter((key) => !(key === "newPrice" && !useNewPriceCol))
    .map((key) => ({
      key,
      label: hdrs[key],
      required:
        key === "articleNo" ||
        (useNewPriceCol && key === "newPrice") ||
        (!useNewPriceCol && key === "priceOrderUnit"),
    }));
}
