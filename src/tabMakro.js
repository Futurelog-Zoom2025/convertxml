// Tab 4: Makro → XML converter.
// Uploads a Makro (CPAxtra) price list plus the matching Report 1145 file,
// runs the VBA VAT calculation on the Makro rows, then merges by product code
// (Makro รหัสสินค้า ↔ 1145 Article no.) and runs the SAME validation + XML
// pipeline as the other tabs.
//
// Mirrors tabP2P.js. Differences: no Options card (the Makro new price and
// lead-time rules are fixed by the VBA), and no language picker (the Makro file
// is always Thai).

import * as XLSX from "xlsx";
import { parseReport1145, NA_MARKER } from "./reportParser.js";
import { validate } from "./validator.js";
import { generateXml } from "./xmlGenerator.js";
import { parseFirstMakroSheet } from "./makroParser.js";
import { mergeMakroAndReport1145 } from "./makroMerger.js";
import { makroHeaderDisplayList } from "./makroHeaders.js";
import {
  $, escapeHtml, formatBytes, todayDDMMYYYY, delay,
  runWithLoading, downloadBlob, buildCompanyMultiselect, restrictToDigits,
  summarizeErrorForHint,
} from "./shared.js";
import { openFullDataModal } from "./fullDataModal.js";

export function initMakroTab() {
  const els = {
    // Upload Makro
    resetBtn:       $("#makroResetBtn"),
    dropZone:       $("#makroDropZone"),
    fileInput:      $("#makroFileInput"),
    fileInfo:       $("#makroFileInfo"),
    showHeadersBtn: $("#makroShowHeadersBtn"),
    hdrModal:       $("#makroHdrModal"),
    hdrBox:         $("#makroHdrBox"),
    hdrClose:       $("#makroHdrClose"),
    hdrCopyBtn:     $("#makroHdrCopyBtn"),
    // Upload R1145
    r1145Drop:      $("#makroR1145DropZone"),
    r1145FileInput: $("#makroR1145FileInput"),
    r1145FileInfo:  $("#makroR1145FileInfo"),
    // Status
    mergeStatus:    $("#makroMergeStatus"),
    // Preview
    previewCard:    $("#makroPreviewCard"),
    previewSummary: $("#makroPreviewSummary"),
    previewTable:   $("#makroPreviewTable"),
    showFullBtn:    $("#makroShowFullBtn"),
    unmatchedBtn:   $("#makroUnmatchedBtn"),
    // Params
    supplierNo:     $("#makroSupplierNo"),
    language:       $("#makroLanguage"),
    validityDate:   $("#makroValidityDate"),
    generateBtn:    $("#makroGenerateBtn"),
    genHint:        $("#makroGenHint"),
  };

  const companyPicker = buildCompanyMultiselect({
    rootId: "makroCompanyMultiselect",
    btnId: "makroCompanyBtn",
    labelId: "makroCompanyBtnLabel",
    menuId: "makroCompanyMenu",
    optionsId: "makroCompanyOptions",
    selectAllId: "makroSelectAllCompanies",
    clearAllId: "makroClearAllCompanies",
  });

  const state = {
    makroParsed: null,  // {rows, sheetName, totalSheets}
    r1145Rows: null,    // rows[]
    mergedRows: [],
    makroOnlyRows: [],  // Makro products with no matching 1145 article
    invalidCells: new Map(),
    summary: null,
  };

  // Enable/label the "Makro not in 1145" button from the current unmatched count.
  function updateUnmatchedBtn() {
    const n = state.makroOnlyRows.length;
    els.unmatchedBtn.disabled = n === 0;
    els.unmatchedBtn.innerHTML =
      `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;vertical-align:-2px;margin-right:4px"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>` +
      `Makro not in 1145${n ? ` (${n})` : ""}`;
  }

  // ---------- Status + generate gating ----------
  function setStatus(kind, html) {
    els.mergeStatus.className = `status ${kind}`;
    els.mergeStatus.innerHTML = html;
    els.mergeStatus.classList.remove("hidden");
  }
  function clearStatus() {
    els.mergeStatus.className = "status hidden";
    els.mergeStatus.innerHTML = "";
  }

  function setGenerateReady(kind, detail) {
    if (kind === "empty") {
      els.generateBtn.disabled = true;
      els.genHint.textContent = detail || "Upload both files first.";
      els.genHint.className = "gen-hint";
    } else if (kind === "error") {
      els.generateBtn.disabled = true;
      if (typeof detail === "string" && detail.trim() !== "") {
        els.genHint.textContent = `⚠ ${detail}`;
      } else {
        const n = typeof detail === "number" ? detail : 0;
        els.genHint.textContent = n > 0
          ? `⚠ Fix the ${n} validation issue${n === 1 ? "" : "s"} above before generating.`
          : "⚠ Fix the validation issues above before generating.";
      }
      els.genHint.className = "gen-hint warn";
    } else if (kind === "ready") {
      els.generateBtn.disabled = false;
      els.genHint.textContent = detail || "Ready — select one or more Company IDs and click Generate.";
      els.genHint.className = "gen-hint ready";
    }
  }

  function resetAll() {
    state.makroParsed = null;
    state.r1145Rows = null;
    state.mergedRows = [];
    state.makroOnlyRows = [];
    state.invalidCells = new Map();
    state.summary = null;
    updateUnmatchedBtn();
    els.fileInput.value = "";
    els.r1145FileInput.value = "";
    els.fileInfo.classList.add("hidden");
    els.fileInfo.innerHTML = "";
    els.r1145FileInfo.classList.add("hidden");
    els.r1145FileInfo.innerHTML = "";
    els.previewCard.classList.add("hidden");
    els.hdrModal.classList.add("hidden");
    companyPicker.reset();
    els.supplierNo.value = "";
    els.validityDate.value = "";
    els.language.value = "TH";
    clearStatus();
    setGenerateReady("empty");
  }

  // ---------- Expected-headers popup ----------
  function renderHeaderBox() {
    const list = makroHeaderDisplayList();
    els.hdrBox.innerHTML = list
      .map((h) => (h.required ? `<span class="req">${escapeHtml(h.label)}</span>` : escapeHtml(h.label)))
      .join("\t");
  }

  // ---------- File handling ----------
  async function handleMakroFile(file) {
    clearStatus();
    try {
      const data = await file.arrayBuffer();
      const parsed = await runWithLoading(
        "Parsing Makro file…",
        "Reading prices and running the VAT calculation.",
        () => {
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          const sheets = wb.SheetNames.map((name) => ({
            name,
            aoa: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: true }),
          }));
          return parseFirstMakroSheet(sheets);
        }
      );

      if (parsed.rows.length === 0) throw new Error("No data rows found in the Makro file.");
      state.makroParsed = parsed;

      const sheetNote = parsed.totalSheets > 1
        ? ` · sheet "${escapeHtml(parsed.sheetName)}" of ${parsed.totalSheets}`
        : "";

      els.fileInfo.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
        <span class="name">${escapeHtml(file.name)}</span>
        <span class="size">· ${formatBytes(file.size)} · ${parsed.rows.length} product${parsed.rows.length === 1 ? "" : "s"}${sheetNote}</span>
      `;
      els.fileInfo.classList.remove("hidden");

      if (!els.validityDate.value) els.validityDate.value = todayDDMMYYYY();
      await tryMerge();
    } catch (err) {
      console.error(err);
      state.makroParsed = null;
      els.fileInfo.classList.add("hidden");
      setStatus("error", `<h3>Could not read the Makro file</h3>${escapeHtml(err.message || String(err))}`);
      setGenerateReady("empty");
    }
  }

  async function handleR1145File(file) {
    clearStatus();
    try {
      const data = await file.arrayBuffer();
      const rows = await runWithLoading(
        "Parsing Report 1145 file…",
        "Indexing articles for lookup.",
        () => {
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
          return parseReport1145(aoa);
        }
      );
      if (rows.length === 0) throw new Error("No data rows found in the Report 1145 file.");
      state.r1145Rows = rows;

      els.r1145FileInfo.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
        <span class="name">${escapeHtml(file.name)}</span>
        <span class="size">· ${formatBytes(file.size)} · ${rows.length} article${rows.length === 1 ? "" : "s"}</span>
      `;
      els.r1145FileInfo.classList.remove("hidden");

      await tryMerge();
    } catch (err) {
      console.error(err);
      state.r1145Rows = null;
      els.r1145FileInfo.classList.add("hidden");
      setStatus("error", `<h3>Could not read the Report 1145 file</h3>${escapeHtml(err.message || String(err))}`);
      setGenerateReady("empty");
    }
  }

  // ---------- Merge + validate ----------
  async function tryMerge() {
    if (!state.makroParsed || !state.r1145Rows) {
      const missing = [];
      if (!state.makroParsed) missing.push("Makro file");
      if (!state.r1145Rows) missing.push("Report 1145 file");
      setGenerateReady("empty", `Upload the ${missing.join(" and ")} to continue.`);
      return;
    }

    const { rows, summary, makroOnlyRows } = await runWithLoading(
      "Merging Makro prices with Report 1145…",
      `Matching ${state.r1145Rows.length.toLocaleString()} Report 1145 articles against ${state.makroParsed.rows.length.toLocaleString()} Makro prices.`,
      () => mergeMakroAndReport1145(state.r1145Rows, state.makroParsed.rows),
    );

    state.mergedRows = rows;
    state.makroOnlyRows = makroOnlyRows;
    state.summary = summary;
    updateUnmatchedBtn();

    // Row-level validation with dummy params (same as P2P).
    const { invalidCells, errors, warnings } = await runWithLoading(
      "Validating merged data…",
      `Checking ${rows.length.toLocaleString()} row${rows.length === 1 ? "" : "s"} against all rules.`,
      () => validate(rows, { companyId: "000", supplierNo: "000000", language: "TH", validityDate: "01012026" })
    );
    state.invalidCells = invalidCells;

    // Per the business: warnings (e.g. "Scaled price blank → used the 1145
    // price") should report a TOTAL count only, not list every row number.
    // The count is already in the message ("N row(s)…"); strip the "(rows: …)".
    const cleanWarnings = warnings.map((w) => w.replace(/\s*\(rows:[^)]*\)/gi, ""));

    renderPreview(rows, invalidCells);

    // Discontinued rows report a total count both in the summary and as a
    // warning (per the business request to surface how many were discontinued).
    if (summary.discontinued > 0) {
      cleanWarnings.push(
        `${summary.discontinued} row(s) marked "Discontinue" in Makro (สถานะ) — ` +
        `kept the Report 1145 price and closed lead time (0).`
      );
    }
    if (summary.noPrice > 0) {
      cleanWarnings.push(
        `${summary.noPrice} row(s) had no usable Makro price (≤ 0 / blank) — ` +
        `kept the Report 1145 price and closed lead time (0).`
      );
    }

    const discontinuedNote = summary.discontinued > 0
      ? `  <li><strong>${summary.discontinued}</strong> article${summary.discontinued === 1 ? "" : "s"} marked <strong>Discontinue</strong> in Makro (kept the Report 1145 price, lead time 0)</li>`
      : "";
    const noPriceNote = summary.noPrice > 0
      ? `  <li><strong>${summary.noPrice}</strong> article${summary.noPrice === 1 ? "" : "s"} with <strong>no usable Makro price</strong> (≤ 0 / blank) (kept the Report 1145 price, lead time 0)</li>`
      : "";
    const makroOnlyNote = summary.makroOnly > 0
      ? `  <li><strong>${summary.makroOnly}</strong> Makro price${summary.makroOnly === 1 ? "" : "s"} had no matching Report 1145 article (ignored — not in output)</li>`
      : "";
    const summaryHtml =
      `<h3>Merge complete</h3>` +
      `<ul>` +
      `  <li><strong>${summary.matched}</strong> article${summary.matched === 1 ? "" : "s"} updated with the Makro price (lead time set to 1)</li>` +
      `  <li><strong>${summary.noInfo}</strong> article${summary.noInfo === 1 ? "" : "s"} with "No Information" (kept the Report 1145 price, lead time 0)</li>` +
      discontinuedNote +
      noPriceNote +
      makroOnlyNote +
      `</ul>`;

    if (errors.length) {
      const list = errors.map((e) => `<li>${escapeHtml(e).replace(/\n/g, "<br>")}</li>`).join("");
      const warnList = cleanWarnings.length
        ? `<p style="margin-top:10px"><strong>Warnings:</strong></p><ul>${cleanWarnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("")}</ul>`
        : "";
      setStatus("error",
        summaryHtml +
        `<p style="margin-top:10px"><strong>Validation failed — ${errors.length} issue${errors.length === 1 ? "" : "s"}:</strong></p><ul>${list}</ul>${warnList}`
      );
      setGenerateReady("error", errors.length);
      return;
    }

    if (cleanWarnings.length) {
      const list = cleanWarnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("");
      setStatus("warn", summaryHtml + `<p style="margin-top:10px"><strong>Warnings:</strong></p><ul>${list}</ul>`);
    } else {
      setStatus("success", summaryHtml + `<p style="margin-top:6px">All rows validated successfully.</p>`);
    }
    setGenerateReady("ready");
  }

  // ---------- Preview ----------
  const STATUS_PILL = {
    "":                          { label: "—",             cls: "pill-none" },
    "Price updated from Makro":  { label: "Price updated", cls: "pill-open" },
    "No Information":            { label: "No Information", cls: "pill-nopu" },
    "Discontinued":              { label: "Discontinued",  cls: "pill-p1145" },
    "No Makro price":            { label: "No Makro price", cls: "pill-p2ponly" },
  };

  const PREVIEW_COLS = [
    { key: "pos",          label: "#",           cls: "c-pos" },
    { key: "itemNo",       label: "Article no.", cls: "c-item" },
    { key: "descGB",       label: "Description", cls: "c-en" },
    { key: "ou",           label: "OU",          cls: "c-unit" },
    { key: "__oldPrice",   label: "Old",         cls: "c-price" },
    { key: "priceOU",      label: "New",         cls: "c-price" },
    { key: "__diff",       label: "Diff",        cls: "c-cuou" },
    { key: "availability", label: "LT",          cls: "c-avail" },
    { key: "status",       label: "Status",      cls: "c-cust" },
  ];

  // Full-data modal — all fields + the Makro-specific calc columns, mirroring
  // the VBA Calculate_Page layout so the user sees the same data as the workbook.
  //
  // Column-group highlighting:
  //   red    (makro-*) = raw data pulled straight from the Makro file
  //   yellow (vat-*)   = values produced by the VBA VAT calculation
  // Headers are always tinted; cells are tinted only on rows that actually
  // carry Makro data (matched rows) — "No Information" rows stay plain.
  const MAKRO_HDR = "makro-header";
  const VAT_HDR   = "vat-header";
  const makroCell = (r) => (r.__source === "matched" ? "makro-cell" : "");
  const vatCell   = (r) => (r.__source === "matched" ? "vat-cell" : "");
  const fmtPct = (r) =>
    typeof r.__makroVatPct === "number" ? `${(r.__makroVatPct * 100).toFixed(2)}%` : "";

  const FULL_COLS = [
    // ----- From Report 1145 -----
    { key: "pos",        label: "#" },
    { key: "itemNo",     label: "Article no." },
    { key: "ean",        label: "GTIN" },
    { key: "descGB",     label: "Item name (English)" },
    { key: "descExtra",  label: "Item name (Local)" },
    { key: "ou",         label: "OU" },
    { key: "cu",         label: "CU" },
    { key: "cuou",       label: "Packaging unit" },
    { key: "origin",     label: "Country of origin" },
    { key: "__oldPrice", label: "Old Price/OU (1145)" },
    // ----- Raw Makro data (red) -----
    { key: "__makroArtGroup", label: "Art. Group",       headerClass: MAKRO_HDR, cellClass: makroCell },
    { key: "__makroName",     label: "ชื่อสินค้า (Makro)",  headerClass: MAKRO_HDR, cellClass: makroCell },
    { key: "__makroExVat",    label: "ราคาขาย (Ex. VAT)",  headerClass: MAKRO_HDR, cellClass: makroCell },
    { key: "__makroVat",      label: "VAT",              headerClass: MAKRO_HDR, cellClass: makroCell },
    { key: "__makroInVat",    label: "ราคาขาย (In. VAT)",  headerClass: MAKRO_HDR, cellClass: makroCell },
    { key: "__makroStatus",   label: "สถานะ",             headerClass: MAKRO_HDR, cellClass: makroCell },
    // ----- VAT calculation (yellow) -----
    { key: "__makroVatPct",        label: "VAT%",                   headerClass: VAT_HDR, cellClass: vatCell, cellHtml: fmtPct },
    { key: "__makroPriceExVat",    label: "Price Exclude VAT",      headerClass: VAT_HDR, cellClass: vatCell },
    { key: "__makroPriceInVat",    label: "Price Include VAT",      headerClass: VAT_HDR, cellClass: vatCell },
    { key: "__makroDiffDecimal",   label: "Diff (Decimal)",         headerClass: VAT_HDR, cellClass: vatCell },
    { key: "__makroExVatAdj",      label: "Price Exclude VAT(Adj)", headerClass: VAT_HDR, cellClass: vatCell },
    { key: "__makroPriceInVatAdj", label: "Price Include VAT(Adj)", headerClass: VAT_HDR, cellClass: vatCell },
    { key: "__makroCheckDiff",     label: "Check Diff",             headerClass: VAT_HDR, cellClass: vatCell },
    // ----- Result -----
    { key: "priceOU",      label: "New Price/OU" },
    { key: "__diff",       label: "Diff (Old − New)" },
    { key: "availability", label: "Lead time (Adj)" },
    { key: "customerId",   label: "Customer ID" },
    {
      key: "status",
      label: "Status",
      cellHtml: (r) => {
        const info = STATUS_PILL[r.status || ""] || STATUS_PILL[""];
        return `<span class="status-pill ${info.cls}">${escapeHtml(info.label)}</span>`;
      },
    },
  ];

  // Columns for the "Makro not in 1145" viewer — ONLY Makro-side columns (these
  // rows have no Report 1145 data). Keys map to raw parseMakroFile fields.
  const fmtPctRaw = (r) => (typeof r.vatPct === "number" ? `${(r.vatPct * 100).toFixed(2)}%` : "");
  const MAKRO_ONLY_COLS = [
    { key: "pos",      label: "#" },
    // ----- Raw Makro columns (red) -----
    { key: "artCode",  label: "รหัสสินค้า",         headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    { key: "itemName", label: "ชื่อสินค้า",          headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    { key: "artGroup", label: "Art. Group",        headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    { key: "srcExVat", label: "ราคาขาย (Ex. VAT)",  headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    { key: "vatAmt",   label: "VAT",               headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    { key: "inVat",    label: "ราคาขาย (In. VAT)",  headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    { key: "status",   label: "สถานะ",              headerClass: MAKRO_HDR, cellClass: () => "makro-cell" },
    // ----- VAT calculation (yellow) -----
    { key: "vatPct",        label: "VAT%",                   headerClass: VAT_HDR, cellClass: () => "vat-cell", cellHtml: fmtPctRaw },
    { key: "priceExVat",    label: "Price Exclude VAT",      headerClass: VAT_HDR, cellClass: () => "vat-cell" },
    { key: "priceInVat",    label: "Price Include VAT",      headerClass: VAT_HDR, cellClass: () => "vat-cell" },
    { key: "diffDecimal",   label: "Diff (Decimal)",         headerClass: VAT_HDR, cellClass: () => "vat-cell" },
    { key: "priceExVatAdj", label: "Price Exclude VAT(Adj)", headerClass: VAT_HDR, cellClass: () => "vat-cell" },
    { key: "priceInVatAdj", label: "Price Include VAT(Adj)", headerClass: VAT_HDR, cellClass: () => "vat-cell" },
    { key: "checkDiff",     label: "Check Diff",             headerClass: VAT_HDR, cellClass: () => "vat-cell" },
    { key: "newPrice",      label: "New Price (Ex.VAT Adj)", headerClass: VAT_HDR, cellClass: () => "vat-cell" },
  ];

  function fmtMoney(v) {
    if (v === "" || v === null || v === undefined) return "—";
    if (typeof v !== "number") return String(v);
    return v.toFixed(2);
  }

  function diffText(r) {
    const v = r.__diff;
    if (typeof v !== "number" || !Number.isFinite(v)) return "—";
    return v === 0 ? "0" : v.toFixed(2);   // __diff is already snapped to a clean 0
  }

  function renderPreview(rows, invalidCells = new Map()) {
    const showRows = rows.slice(0, 200);
    const head = `<thead><tr>${PREVIEW_COLS.map((c) => `<th class="${c.cls}">${c.label}</th>`).join("")}</tr></thead>`;
    const body = "<tbody>" + showRows.map((r, idx) => {
      const invalid = invalidCells.get(idx) || new Set();
      return "<tr>" + PREVIEW_COLS.map((c) => {
        let shown;
        if (c.key === "__oldPrice" || c.key === "priceOU") shown = fmtMoney(r[c.key]);
        else if (c.key === "__diff") shown = diffText(r);
        else if (c.key === "status") {
          const info = STATUS_PILL[r.status || ""] || STATUS_PILL[""];
          return `<td class="${c.cls} status-cell"><span class="status-pill ${info.cls}">${escapeHtml(info.label)}</span></td>`;
        } else {
          const v = r[c.key];
          shown = v === NA_MARKER ? "#N/A" : (v === null || v === undefined ? "" : String(v));
        }

        let extraCls = "";
        if (c.key === "priceOU") {
          if (r.status === "Price updated from Makro") extraCls = " price-updated";
          else if (r.status === "No Information" || r.status === "Discontinued" || r.status === "No Makro price") extraCls = " price-fallback";
        }

        const invalidCls = invalid.has(c.key) ? " invalid-cell" : "";
        return `<td class="${c.cls}${invalidCls}${extraCls}" title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
      }).join("") + "</tr>";
    }).join("") + "</tbody>";
    els.previewTable.innerHTML = head + body;

    els.previewSummary.textContent = rows.length > 200
      ? `Showing first 200 of ${rows.length} rows — click "Show Full Data" to see all.`
      : `Showing all ${rows.length} row${rows.length === 1 ? "" : "s"}.`;
    els.previewCard.classList.remove("hidden");
  }

  // ---------- Generate XML ----------
  function getParams(companyId) {
    return {
      companyId,
      supplierNo: els.supplierNo.value.trim(),
      language: els.language.value.trim(),
      validityDate: els.validityDate.value.trim(),
    };
  }

  async function runGenerate() {
    const companies = companyPicker.getSelected();
    if (companies.length === 0) {
      setStatus("error", `<h3>Select a Company ID</h3>Please select at least one WebShop Company ID.`);
      return;
    }

    const validationParams = getParams(companies[0] || "000");
    const { errors } = validate(state.mergedRows, validationParams);
    if (errors.length) {
      const list = errors.map((e) => `<li>${escapeHtml(e).replace(/\n/g, "<br>")}</li>`).join("");
      setStatus("error", `<h3>Validation failed — ${errors.length} issue${errors.length === 1 ? "" : "s"}</h3><ul>${list}</ul>`);
      setGenerateReady("error", errors.length);
      return;
    }

    const createdFiles = [];
    try {
      for (let i = 0; i < companies.length; i++) {
        const companyId = companies[i];
        const params = getParams(companyId);
        const result = await runWithLoading(
          `Generating XML ${i + 1} of ${companies.length}…`,
          `Company ${companyId} · ${state.mergedRows.length.toLocaleString()} article${state.mergedRows.length === 1 ? "" : "s"}`,
          () => generateXml(state.mergedRows, params)
        );
        const { xml, filename } = result;
        downloadBlob("﻿" + xml, filename, "application/xml;charset=utf-8");
        createdFiles.push(filename);
        if (i < companies.length - 1) await delay(350);
      }
      const fileList = createdFiles.map((f) => `<li><code>${escapeHtml(f)}</code></li>`).join("");
      setStatus("success",
        `<h3>XML generated for ${createdFiles.length} compan${createdFiles.length === 1 ? "y" : "ies"}</h3>` +
        `<ul>${fileList}</ul>`
      );
    } catch (err) {
      console.error(err);
      setStatus("error", `<h3>Generation failed</h3>${escapeHtml(err.message || String(err))}`);
    }
  }

  // ---------- Event wiring ----------
  // Makro drop zone
  els.dropZone.addEventListener("click", () => els.fileInput.click());
  els.dropZone.addEventListener("dragover", (e) => { e.preventDefault(); els.dropZone.classList.add("dragging"); });
  els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragging"));
  els.dropZone.addEventListener("drop", (e) => {
    e.preventDefault(); els.dropZone.classList.remove("dragging");
    const f = e.dataTransfer.files[0]; if (f) handleMakroFile(f);
  });
  els.fileInput.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) handleMakroFile(f); });

  // R1145 drop zone
  els.r1145Drop.addEventListener("click", () => els.r1145FileInput.click());
  els.r1145Drop.addEventListener("dragover", (e) => { e.preventDefault(); els.r1145Drop.classList.add("dragging"); });
  els.r1145Drop.addEventListener("dragleave", () => els.r1145Drop.classList.remove("dragging"));
  els.r1145Drop.addEventListener("drop", (e) => {
    e.preventDefault(); els.r1145Drop.classList.remove("dragging");
    const f = e.dataTransfer.files[0]; if (f) handleR1145File(f);
  });
  els.r1145FileInput.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) handleR1145File(f); });

  // Expected-headers modal
  els.showHeadersBtn.addEventListener("click", () => {
    renderHeaderBox();
    els.hdrModal.classList.toggle("hidden");
  });
  els.hdrClose.addEventListener("click", () => els.hdrModal.classList.add("hidden"));
  els.hdrCopyBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(els.hdrBox.innerText);
      const orig = els.hdrCopyBtn.textContent;
      els.hdrCopyBtn.textContent = "✓ Copied";
      setTimeout(() => { els.hdrCopyBtn.textContent = orig; }, 1200);
    } catch (_) { /* no clipboard access — user can copy manually */ }
  });

  // Params & generate
  els.generateBtn.addEventListener("click", runGenerate);
  els.resetBtn.addEventListener("click", resetAll);

  els.showFullBtn.addEventListener("click", () => {
    openFullDataModal({
      rows: state.mergedRows,
      columns: FULL_COLS,
      invalidCells: state.invalidCells,
      statusPillMap: STATUS_PILL,
      exportFilename: "Makro_Merge_Export",
      exportSheetName: "Makro Merge",
    });
  });

  // "Makro not in 1145" — Makro products with no matching Report 1145 article.
  // Reuses the full-data modal (search across all columns + Export to Excel),
  // showing only Makro-side columns.
  els.unmatchedBtn.addEventListener("click", () => {
    if (!state.makroOnlyRows.length) return;
    openFullDataModal({
      rows: state.makroOnlyRows,
      columns: MAKRO_ONLY_COLS,
      loadingHint: `Preparing ${state.makroOnlyRows.length.toLocaleString()} unmatched Makro product(s).`,
      exportFilename: "Makro_Not_In_1145",
      exportSheetName: "Makro not in 1145",
    });
  });

  // Re-run param-level validation whenever a Step-3 input changes.
  let revalidateTimer = null;
  function revalidateParams() {
    if (!state.mergedRows.length) return;
    const companies = companyPicker.getSelected();
    const validationParams = getParams(companies[0] || "000");
    const { errors } = validate(state.mergedRows, validationParams);
    if (errors.length) {
      const hint = summarizeErrorForHint(errors);
      setGenerateReady("error", hint || errors.length);
    } else {
      setGenerateReady("ready");
    }
  }
  function scheduleRevalidate() {
    clearTimeout(revalidateTimer);
    revalidateTimer = setTimeout(revalidateParams, 250);
  }
  els.supplierNo.addEventListener("input", scheduleRevalidate);
  els.validityDate.addEventListener("input", scheduleRevalidate);
  els.language.addEventListener("change", scheduleRevalidate);
  companyPicker.onChange(scheduleRevalidate);

  restrictToDigits(els.supplierNo);
  restrictToDigits(els.validityDate);
  els.validityDate.placeholder = todayDDMMYYYY();
  renderHeaderBox();
  setGenerateReady("empty");
}
