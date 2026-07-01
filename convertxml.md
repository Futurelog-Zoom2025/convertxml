# XML Converter

A static, browser-only web app (FutureLog **Master Data Tools**) for converting supplier price data to and from the **FUTURELOG** MDM XML format.

Everything runs client-side — **no backend, nothing is uploaded to any server**. Excel/XML files are parsed in the browser with [SheetJS](https://sheetjs.com/). Designed for free static hosting on **Cloudflare Pages**.

The app has **four tabs**, each a self-contained converter:

| Tab | Input | Output |
| --- | --- | --- |
| **Convert XML from 1145 report** | Report 1145 `.xls/.xlsx` | FUTURELOG `.cat.xml` |
| **Convert XML from P2P report** | P2P file + matching Report 1145 | FUTURELOG `.cat.xml` |
| **Convert XML from Makro file** | Makro price list + matching Report 1145 | FUTURELOG `.cat.xml` |
| **Convert XML to report 1145** | FUTURELOG `.xml` | Report 1145 `.xlsx` |

All logic (VAT calculation, price/lead-time rules, validation, the XML writer) is ported 1:1 from the original Excel/VBA macros so the output matches the workbooks exactly.

---

## Tab 1 — Convert XML from 1145 report

Upload a Report 1145 Excel file → validate → download the FUTURELOG XML.

- Columns are matched by **header name** on row 4 (EN / TH / VN supported), so column order and extra columns don't matter.
- Price rule (from `Get_Data_From_File1145`): `priceOU = scaledPrice` when scaled price is a positive number, otherwise the unit price; lead time is carried through only when the scaled price was used, otherwise closed to `0`.
- Select **one or more WebShop Company IDs** to generate a separate XML per company.
- Filename pattern: `{CompanyID}{SupplierNo}{YYYYMMDD}.cat.xml`.
- A blank Report 1145 template can be downloaded from the tab.

## Tab 2 — Convert XML from P2P report

Upload a **P2P file** (supplier list from the hotel) plus the **matching Report 1145**; the app merges them by **Article No.** and reuses the same validation + XML pipeline.

- Header language picker: **EN / VN** (auto-detected; "Show expected headers" popup available).
- **Options:**
  - *NEW PRICE column* — pull the new price from a dedicated `NEW PRICE` column instead of `Price / Order unit`.
  - *Open lead time when a new price exists* — rows with a new price get lead time `1`, rows without get `0`.
- Rows with no usable P2P price fall back to the Report 1145 price ("Price from Report 1145"); items only in P2P or only in 1145 are labelled accordingly.

## Tab 3 — Convert XML from Makro file

Upload a **Makro (CPAxtra) price list** plus the **matching Report 1145**. Ported from the *For Makro Project* workbook (Module1 VAT calc + Module2 lookup).

**Makro file format** (headers on **row 1**, data from row 2, Thai — no language picker):

| Header | Required | Notes |
| --- | --- | --- |
| `รหัสสินค้า` | ✔ | product code — the join key |
| `ชื่อสินค้า` | ✔ | product name |
| `ราคาขาย (Ex. VAT)` | ✔ | selling price excl. VAT (display only) |
| `VAT` | ✔ | VAT amount per unit (drives the VAT % branch) |
| `ราคาขาย (In. VAT)` | ✔ | selling price incl. VAT (also accepts `ราคาขาย (รวม VAT)`) |
| `สถานะ` | — | status (Active / Discontinue) |
| `Art. Group` | — | article group |

**VAT calculation** (ported verbatim — verified 0 mismatches against the workbook across all rows):

```
VAT%             = 0% if VAT amount = 0, else 7%
Price Exclude VAT (I) = ROUND( InVAT / (1 + VAT%), 2 )
Price Include VAT (J) = ROUND( VAT amount + I, 2 )
Diff (Decimal)   (K) = J − InVAT
Price Exclude VAT(Adj) (L) = I − K          ← the "new price" basis
Price Include VAT(Adj) (M) = L + VAT amount
Check Diff       (N) = M − InVAT
New price            = ROUND(L, 2)
```

**Merge** (driven by **all Report 1145 rows**; join `Article no.` ↔ Makro `รหัสสินค้า`):

- **Matched (active):** price = Makro `ROUND(L, 2)`, lead time = `1`.
- **Not found:** keep the Report 1145 price, lead time = `0` — flagged **"No Information"**.
- **Discontinued** (`สถานะ` contains *Discontinue*): keep the Report 1145 price and close lead time to `0` — the merge summary and a warning report **how many** were discontinued.

Descriptions, units, GTIN and origin come from the Report 1145 file; the exported price (`<PRCOU>`) is the Makro Ex-VAT(Adj) new price.

**Viewers:**
- **Show Full Data** — every merged row, with the raw Makro columns tinted **red** and the VAT-calc columns tinted **yellow**; search, sort, status filter, and Export to Excel.
- **Makro not in 1145** — Makro products whose code had no matching Report 1145 article (not part of the XML output); Makro-only columns, search, and Export to Excel.

## Tab 4 — Convert XML to report 1145

Upload a FUTURELOG `.xml` (the kind this app produces) and get a Report 1145 `.xlsx` back — the reverse direction. Division, supplier number and validity date are decoded from the filename; `PRCOU` fills *Price per order unit*, and a *Customer ID* column is appended so multi-tier pricing isn't lost.

---

## Shared behavior

- **Validation** (`validator.js`, ported from `generate_xml`): mandatory fields; duplicate `ItemNo + CustomerID + EAN`; Item No / EAN conflicts; GTIN 13-digit format; field-length limits; unit-code and country-code lists; Customer ID `"0000"` core-entry rule; price/lead-time warnings.
- **Full Data modal** (`fullDataModal.js`): sortable columns, search-any-column, "Errors only" toggle, status filter, and **Export to Excel** (`exportExcel.js`, preserves error/warning highlighting) — shared by all data tabs.
- **First-visit popup** announces the newest feature once per browser (dismissal stored in `localStorage`).
- **Privacy:** files never leave the browser — parsing, validation and generation are entirely client-side.

---

## Project structure

```
convertxml/
├── index.html               # UI shell (tabs, panels, modals)
├── style.css                # Styles + full-data highlight classes
├── package.json
├── vite.config.js
├── wrangler.jsonc
├── public/
│   └── _headers             # Cloudflare security headers
└── src/
    ├── main.js              # Tab switching, tab init, first-visit popup
    ├── shared.js            # $, company multi-select, loading overlay, download, helpers
    ├── reportParser.js      # Report 1145 → Tabelle1 mapping (+ price logic)
    ├── validator.js         # Port of VBA generate_xml validation
    ├── xmlGenerator.js      # Port of VBA XML writer
    ├── xmlParser.js         # FUTURELOG XML → rows (Tab 4)
    ├── referenceData.js     # Unit_List, Country_List, valid languages
    ├── templateGenerator.js # Blank Report 1145 template download
    ├── r1145Writer.js       # Rows → Report 1145 .xlsx (Tab 4)
    ├── exportExcel.js       # Full-data modal → styled .xlsx
    ├── fullDataModal.js     # Shared full-data viewer (search/sort/filter/export)
    ├── tab1145.js           # Tab 1 controller
    ├── tabP2P.js            # Tab 2 controller
    ├── p2pHeaders.js        # P2P header aliases (EN/VN) + expected-headers list
    ├── p2pParser.js         # P2P file parser
    ├── p2pMerger.js         # P2P ↔ Report 1145 merge
    ├── tabMakro.js          # Tab 3 controller
    ├── makroHeaders.js      # Makro header aliases (Thai) + expected-headers list
    ├── makroParser.js       # Makro parser + VAT calculation
    ├── makroMerger.js       # Makro ↔ Report 1145 merge (+ discontinue rule)
    └── tabXmlToR1145.js     # Tab 4 controller
```

---

## Local development

Requires [Node.js](https://nodejs.org/) 18+ and npm.

```bash
npm install
npm run dev      # dev server (usually http://localhost:5173)
npm run build    # production build → dist/
npm run preview  # preview the production build
```

---

## Deploy to Cloudflare Pages (via GitHub)

1. Push the repo to GitHub.
2. In [dash.cloudflare.com](https://dash.cloudflare.com) → **Workers & Pages** → **Create application** → **Pages** → **Connect to Git**, select the repo.
3. Build settings:

   | Setting | Value |
   | --- | --- |
   | Framework preset | `None` (or `Vite`) |
   | Build command | `npm run build` |
   | Build output directory | `dist` |
   | `NODE_VERSION` (env var) | `20` |

4. **Save and Deploy.** Every push to `main` auto-rebuilds; other branches get preview URLs.

---

## Output XML schema

```xml
<?xml version="1.0" encoding="UTF-8" ?>
<FUTURELOG>
  <HEAD><VALIDITY>DDMMYYYY</VALIDITY></HEAD>
  <ARTICLES>
    <ARTICLE>
      <NAME>
        <DE>…</DE><FR>…</FR><IT>…</IT><GB>…</GB>
        <XX>…</XX>             <!-- only when language is NOT DE/FR/IT/GB -->
      </NAME>
      <ARTICLEDATA>
        <ARTNO>…</ARTNO><EAN>…</EAN><CUSTNO>…</CUSTNO>
        <MANARTNO>…</MANARTNO><ORG>…</ORG><PICURL>…</PICURL>
      </ARTICLEDATA>
      <PRICES><PRICE>
        <CUSTID>…</CUSTID><PRCOU>…</PRCOU>
        <OU>…</OU><CU>…</CU><NUCUOU>…</NUCUOU>
        <VLZ>…</VLZ>
        <OFFSTART>…</OFFSTART><OFFEND>…</OFFEND>
      </PRICE></PRICES>
    </ARTICLE>
    <!-- … more articles … -->
  </ARTICLES>
</FUTURELOG>
```

---

## Notes

- SheetJS is pinned to `xlsx@0.18.5`. Its advisory-flagged issues require parsing untrusted input; since files are only ever the ones the user opens in their own browser, the practical risk is minimal. To use the latest, point the `xlsx` dependency at `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz` and reinstall.
- Pure static HTML + JS — no Workers, no paid features required.
