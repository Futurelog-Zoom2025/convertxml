# Report 1145 → XML Converter

A static web app that converts **Report 1145** Excel files into the FUTURELOG XML format used for MDM (Master Data Management) import.

Built to run entirely in the browser — no backend, no data upload to any server, no cost to run. Designed for deployment on **Cloudflare Pages** via GitHub.

---

## Features

- Upload `.xls` / `.xlsx` Report 1145 files (drag & drop or click)
- In-browser parsing with **SheetJS**
- Full validation (ported 1:1 from the original `generate_xml` VBA macro):
  - Mandatory fields
  - Duplicate `ItemNo + CustomerID + EAN` combinations
  - EAN conflicts (same EAN with different Item Nos)
  - Item No conflicts (same Item No with different EANs)
  - Unit code validation against the official `Unit_List`
  - Country code validation against the official `Country_List`
  - "All availability = 0" warning
  - Customer ID `"0000"` core-entry rule
- Generates a valid FUTURELOG XML file with the correct filename pattern:
  `{CompanyID}{SupplierNo}{YYYYMMDD}.cat.xml`
- Download the XML directly in the browser

---

## Local Development

### Prerequisites

- [Node.js](https://nodejs.org/) 18+ and npm

### Setup

```bash
npm install
npm run dev
```

Open the URL Vite prints (usually `http://localhost:5173`).

### Build for production

```bash
npm run build
```

The production build is emitted to `dist/`.

---

## Deploy to Cloudflare Pages (via GitHub)

This app is designed to be deployed as a **static site** on Cloudflare Pages. Follow these steps.

### Step 1 — Push the project to GitHub

If you haven't already, create a new repository on GitHub (e.g., `report-1145-xml-converter`), then:

```bash
cd report-1145-xml-converter
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/report-1145-xml-converter.git
git push -u origin main
```

### Step 2 — Connect Cloudflare Pages to GitHub

1. Log in to [dash.cloudflare.com](https://dash.cloudflare.com).
2. In the left sidebar, go to **Workers & Pages**.
3. Click **Create application** → **Pages** tab → **Connect to Git**.
4. Authorize Cloudflare to access your GitHub account (first time only).
5. Select your repository (`report-1145-xml-converter`) and click **Begin setup**.

### Step 3 — Configure the build

Fill in these exact settings:

| Setting                 | Value           |
| ----------------------- | --------------- |
| **Project name**        | `report-1145-xml-converter` (or any name you like) |
| **Production branch**   | `main`          |
| **Framework preset**    | `None` (or `Vite` if offered) |
| **Build command**       | `npm run build` |
| **Build output directory** | `dist`       |
| **Root directory**      | (leave blank)   |
| **Node version** (env var) | `18` or `20` |

To set the Node version, under **Environment variables** add:

- **Variable name:** `NODE_VERSION`
- **Value:** `20`

Click **Save and Deploy**. Cloudflare will build and deploy the site. You'll get a URL like `https://report-1145-xml-converter.pages.dev`.

### Step 4 — Continuous deployment

Every time you `git push` to `main`, Cloudflare will automatically rebuild and deploy. Pushes to other branches create preview deployments with their own URL.

### Step 5 (optional) — Custom domain

In the Pages project, go to **Custom domains** → **Set up a custom domain** and follow the instructions.

---

## Project structure

```
report-1145-xml-converter/
├── index.html              # UI shell
├── style.css               # Styles
├── package.json
├── vite.config.js
├── public/
│   └── _headers            # Cloudflare security headers
└── src/
    ├── main.js             # UI wiring, file handling, download
    ├── reportParser.js     # Report 1145 → Tabelle1 mapping
    ├── validator.js        # Port of VBA generate_xml validation
    ├── xmlGenerator.js     # Port of VBA XML writer
    └── referenceData.js    # Unit_List and Country_List constants
```

---

## How the conversion works

The original VBA workflow had two stages:

1. **Import** (`Get_Data_From_File1145`) — reads Report 1145, copies columns B–N, O, P, Q, R, S, T–V into a staging sheet called **Tabelle1**, then computes `priceOU = IF(scaledPrice empty or 0, unitPrice, scaledPrice)` and `availability = IF(priceOU == scaledPrice, leadTime, 0)`.
2. **Generate XML** (`generate_xml`) — validates the staged data, then streams it to a FUTURELOG XML file.

This web app performs both stages in one click:

- `reportParser.js` handles stage 1 (mapping Report 1145 columns to the 23-column Tabelle1 shape).
- `validator.js` performs all validations from the VBA.
- `xmlGenerator.js` writes out the XML, keeping the exact tag structure, order, and filename format.

---

## Output XML schema

```xml
<?xml version="1.0" encoding="UTF-8" ?>
<FUTURELOG>
  <HEAD><VALIDITY>DDMMYYYY</VALIDITY></HEAD>
  <ARTICLES>
    <ARTICLE>
      <n>
        <DE>…</DE><FR>…</FR><IT>…</IT><GB>…</GB>
        <XX>…</XX>             <!-- only if language is NOT DE/FR/IT/GB -->
      </n>
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
    <!-- … more articles -->
  </ARTICLES>
</FUTURELOG>
```

---

## Privacy

Your Excel file never leaves your browser. Parsing, validation, and XML generation all happen client-side.

---

## Notes

- The `xlsx` (SheetJS) package used is `0.18.5` from the npm registry. It has two advisory-flagged issues (Prototype Pollution and ReDoS). Both require untrusted input to be parsed — since this app only parses files the user themselves uploaded in their own browser, the practical risk is minimal. If your organization requires the latest patched version, change the dependency in `package.json` to `"xlsx": "https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz"` and run `npm install` again.
- The app is pure static HTML + JS. Cloudflare Pages hosts it as a CDN-served static site — no Cloudflare Workers, no paid features needed.
