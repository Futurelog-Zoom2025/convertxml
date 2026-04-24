// P2P header definitions extracted from the "Language" sheet of
// Importlist_TH_XML (For MDM)__new.xlsm.
//
// The P2P file lays out headers on row 5 (1-based). Row 6 is the first data row.
// Headers may be in English or Vietnamese depending on the supplier's region.
// Column positions in the file are NOT fixed — we match by header name.

// Canonical header strings per language, in the "natural" column order the
// hotels use. EN matches the reference xlsm; VN matches the Language sheet.
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
    // The Language sheet has no VN translation for "NEW PRICE" — hotels
    // that use VN either don't have a NEW PRICE column, or label it in EN.
    // We accept the EN string "NEW PRICE" as a fallback here.
    newPrice:       "NEW PRICE",
    minOrderQty:    "Số lượng đặt hàng tối thiểu",
    originCountry:  "Nguồn gốc xuất xứ",
  },
};

// The field keys we actually need downstream — they map P2P headers to our
// internal row-object keys. `articleNo` and one of the two price fields are
// the only required ones; everything else is optional.
export const P2P_FIELD_KEYS = [
  "wsNo", "itemName", "articleNo", "gtin", "orderUnit", "contentUnits",
  "packagingUnit", "priceOrderUnit", "newPrice", "minOrderQty", "originCountry",
];

/**
 * Ordered list of header strings for display in the "expected headers" popup.
 * Required fields are flagged so the UI can highlight them.
 *
 * @param {"EN"|"VN"} lang
 * @param {boolean} useNewPriceCol  — true when user enabled the "NEW PRICE" toggle
 * @returns {Array<{key:string,label:string,required:boolean}>}
 */
export function headerDisplayList(lang, useNewPriceCol) {
  const hdrs = P2P_HEADERS[lang] || P2P_HEADERS.EN;
  return P2P_FIELD_KEYS.map((key) => ({
    key,
    label: hdrs[key],
    required:
      key === "articleNo" ||
      (useNewPriceCol && key === "newPrice") ||
      (!useNewPriceCol && key === "priceOrderUnit"),
  }));
}
