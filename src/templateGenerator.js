// Generates a blank Report 1145 template (.xlsx) matching the exact layout of
// the sample file Report_1145_Sample.xls:
//   - Row 1: "Create import list (PRI)"
//   - Rows 2-3: blank
//   - Row 4: column headers
//   - Row 5+: data
//
// Users can download this, fill it in, and upload it back into the converter.

import * as XLSX from "xlsx";

const HEADERS = [
  "ID",
  "Type",
  "Item name (German)",
  "Item name (French)",
  "Item name (Italian)",
  "Item name (English)",
  "Item name",
  "Article no.",
  "GTIN",
  "Manufacturer's item number",
  "Producer",
  "Order unit (OU)",
  "Content unit (CU)",
  "Packaging unit",
  "Price per order unit",
  "Scaled price",
  "Country of origin",
  "Customs tariff number",
  "Article<BR>lead time",
  "Item link supplier<br>",
  "Start of special offer",
  "End of special offer",
];

// One sample data row users can use as reference (remove before actually uploading)
const SAMPLE_ROW = [
  1700144, 1, "Thai Item", "", "",
  "Alcohol, solid, Aro, bucket, 14kg",
  "แอลกอฮอล์, แบบแข็ง, เอโร่, ชนิดถัง, 14กก",
  155018, "0000000000000", "", "",
  "EI", "EI", 1, 1335.05, "",
  "XX", "", 10, "", "", "",
];

/**
 * Build a workbook matching the Report 1145 layout and trigger a browser download.
 */
export function downloadTemplate() {
  // Construct the array-of-arrays layout
  // Row 1: "Create import list (PRI)" in col A
  // Row 2-3: blank
  // Row 4: headers
  // Row 5: sample data row (so users see the expected format)
  const aoa = [
    ["Create import list (PRI)"],
    [],
    [],
    HEADERS,
    SAMPLE_ROW,
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Column widths to make the template easy to read
  ws["!cols"] = [
    { wch: 10 }, // ID
    { wch: 6 },  // Type
    { wch: 22 }, // German
    { wch: 22 }, // French
    { wch: 22 }, // Italian
    { wch: 32 }, // English
    { wch: 32 }, // Item name
    { wch: 12 }, // Article no.
    { wch: 16 }, // GTIN
    { wch: 22 }, // Mfg item no
    { wch: 14 }, // Producer
    { wch: 14 }, // OU
    { wch: 14 }, // CU
    { wch: 14 }, // Packaging
    { wch: 16 }, // Price
    { wch: 14 }, // Scaled
    { wch: 18 }, // Origin
    { wch: 18 }, // Customs
    { wch: 16 }, // Lead time
    { wch: 22 }, // Link
    { wch: 16 }, // Offer start
    { wch: 16 }, // Offer end
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Tabelle1");

  // Build the file as a Uint8Array and trigger download
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Report_1145_Template.xlsx";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}
