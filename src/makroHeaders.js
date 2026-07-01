// Makro (CPAxtra / Siam Makro) price-list header definitions.
//
// Unlike the P2P file, the Makro file has NO language variants — the supplier
// always ships Thai headers. So there is a single, Thai-only header set here
// (no EN/VN picker on the tab).
//
// New format agreed with the business (differs from the original VBA workbook,
// which expected the raw dump to start further down the sheet): the Makro file
// the user uploads now has its HEADER ON ROW 1 and DATA FROM ROW 2.
//
// Required columns (must be present):
//   รหัสสินค้า            product code — the join key against Report 1145 "Article no."
//   ชื่อสินค้า            product name (display only)
//   ราคาขาย (Ex. VAT)    selling price excluding VAT (display only — NOT used in the calc)
//   VAT                  the VAT amount per unit (drives the VAT% branch)
//   ราคาขาย (In. VAT)    selling price including VAT (drives the exclude-VAT calc)
//
// Optional columns (some raw dumps don't include them):
//   สถานะ                status (display only)
//   Art. Group           article group (display only)

// Internal field keys used throughout the Makro pipeline.
export const MAKRO_FIELD_KEYS = [
  "artGroup", "artCode", "itemName", "priceExVat", "vat", "priceInVat", "status",
];

// field → list of accepted header spellings. Header matching normalizes by
// stripping ALL whitespace + NFC + lowercase, so spacing differences (e.g.
// "ราคาขาย (Ex. VAT)" vs "ราคาขาย(Ex. VAT)") collapse to the same key and only
// genuinely different wordings need their own alias.
export const MAKRO_ALIASES = {
  artGroup:   ["Art. Group", "Art.Group", "กลุ่มสินค้า"],
  artCode:    ["รหัสสินค้า"],
  itemName:   ["ชื่อสินค้า"],
  priceExVat: ["ราคาขาย (Ex. VAT)", "ราคาขาย(Ex. VAT)", "ราคาขาย (Ex.VAT)"],
  vat:        ["VAT"],
  // The real workbook uses "รวม VAT"; the business also refers to it as
  // "In. VAT". Accept both, plus a couple of common spellings.
  priceInVat: [
    "ราคาขาย (รวม VAT)", "ราคาขาย(รวม VAT)",
    "ราคาขาย (In. VAT)", "ราคาขาย(In. VAT)",
    "ราคาขาย (Incl. VAT)",
  ],
  status:     ["สถานะ"],
};

// Which fields are mandatory. สถานะ and Art. Group are optional because some
// raw dumps omit them.
export const MAKRO_REQUIRED = new Set([
  "artCode", "itemName", "priceExVat", "vat", "priceInVat",
]);

// Human-readable labels shown in the "expected headers" popup, in display order.
// Thai-only — there is no language choice for this file.
const MAKRO_DISPLAY_LABELS = {
  artGroup:   "Art. Group",
  artCode:    "รหัสสินค้า",
  itemName:   "ชื่อสินค้า",
  priceExVat: "ราคาขาย (Ex. VAT)",
  vat:        "VAT",
  priceInVat: "ราคาขาย (In. VAT)",
  status:     "สถานะ",
};

/**
 * Ordered list of header strings for the "expected headers" popup. Required
 * fields are flagged so the UI can highlight them. Optional fields (สถานะ,
 * Art. Group) are shown but not highlighted.
 *
 * Display order matches the real Makro dump: Art. Group first, then the
 * product/price columns, then สถานะ last.
 */
export function makroHeaderDisplayList() {
  const order = ["artCode", "itemName", "priceExVat", "vat", "priceInVat", "status", "artGroup"];
  return order.map((key) => ({
    key,
    label: MAKRO_DISPLAY_LABELS[key],
    required: MAKRO_REQUIRED.has(key),
  }));
}
