// Tab 2: P2P → XML converter.
// Uploads a P2P file (supplier list from the hotel) plus the matching Report
// 1145 file, merges them based on Article No., then runs validation and XML
// generation on the merged rows — same XML pipeline as tab 1.

import * as XLSX from "xlsx";
import { parseReport1145, NA_MARKER } from "./reportParser.js";
import { validate } from "./validator.js";
import { generateXml } from "./xmlGenerator.js";
import { parseP2PFile, parseFirstParseableSheet } from "./p2pParser.js";
import { mergeP2PAndReport1145 } from "./p2pMerger.js";
import { P2P_HEADERS, P2P_FIELD_KEYS, headerDisplayList } from "./p2pHeaders.js";
import {
  $, escapeHtml, formatBytes, todayDDMMYYYY, delay,
  runWithLoading, downloadBlob, buildCompanyMultiselect, restrictToDigits,
} from "./shared.js";
import { openFullDataModal } from "./fullDataModal.js";

export function initP2PTab() {
  const els = {
    // Upload P2P
    resetBtn:      $("#p2pResetBtn"),
    dropZone:      $("#p2pDropZone"),
    fileInput:     $("#p2pFileInput"),
    fileInfo:      $("#p2pFileInfo"),
    supplierBanner:$("#p2pSupplierBanner"),
    supplierIcon:  $("#p2pSupplierIcon"),
    supplierName:  $("#p2pSupplierName"),
    supplierMeta:  $("#p2pSupplierMeta"),
    langPick:      $("#p2pLangPick"),
    showHeadersBtn:$("#p2pShowHeadersBtn"),
    hdrModal:      $("#p2pHdrModal"),
    hdrBox:        $("#p2pHdrBox"),
    hdrClose:      $("#p2pHdrClose"),
    hdrCopyBtn:    $("#p2pHdrCopyBtn"),
    // Upload R1145
    r1145Drop:     $("#p2pR1145DropZone"),
    r1145FileInput:$("#p2pR1145FileInput"),
    r1145FileInfo: $("#p2pR1145FileInfo"),
    // Options
    newPriceCheckbox: $("#p2pNewPriceCheckbox"),
    openLeadCheckbox: $("#p2pOpenLeadCheckbox"),
    mergeStatus:      $("#p2pMergeStatus"),
    // Preview
    previewCard:    $("#p2pPreviewCard"),
    previewSummary: $("#p2pPreviewSummary"),
    previewTable:   $("#p2pPreviewTable"),
    showFullBtn:    $("#p2pShowFullBtn"),
    statusFilter:   $("#p2pStatusFilter"),
    // Params
    supplierNo:   $("#p2pSupplierNo"),
    language:     $("#p2pLanguage"),
    validityDate: $("#p2pValidityDate"),
    generateBtn:  $("#p2pGenerateBtn"),
    genHint:      $("#p2pGenHint"),
  };

  const companyPicker = buildCompanyMultiselect({
    rootId: "p2pCompanyMultiselect",
    btnId: "p2pCompanyBtn",
    labelId: "p2pCompanyBtnLabel",
    menuId: "p2pCompanyMenu",
    optionsId: "p2pCompanyOptions",
    selectAllId: "p2pSelectAllCompanies",
    clearAllId: "p2pClearAllCompanies",
  });

  const state = {
    lang: "EN",
    p2pParsed: null,   // {supplier, division, rows}
    r1145Rows: null,   // rows[]
    mergedRows: [],
    invalidCells: new Map(),
    summary: null,
    statusFilter: "__all__", // currently selected value in the Status dropdown
  };

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
      const n = typeof detail === "number" ? detail : 0;
      els.genHint.textContent = n > 0
        ? `⚠ Fix the ${n} validation issue${n === 1 ? "" : "s"} above before generating.`
        : "⚠ Fix the validation issues above before generating.";
      els.genHint.className = "gen-hint warn";
    } else if (kind === "ready") {
      els.generateBtn.disabled = false;
      els.genHint.textContent = detail || "Ready — select one or more Company IDs and click Generate.";
      els.genHint.className = "gen-hint ready";
    }
  }

  function resetAll() {
    state.p2pParsed = null;
    state.r1145Rows = null;
    state.mergedRows = [];
    state.invalidCells = new Map();
    state.summary = null;
    state.statusFilter = "__all__";
    els.statusFilter.innerHTML = '<option value="__all__">All statuses</option>';
    els.statusFilter.value = "__all__";
    els.fileInput.value = "";
    els.r1145FileInput.value = "";
    els.fileInfo.classList.add("hidden");
    els.fileInfo.innerHTML = "";
    els.r1145FileInfo.classList.add("hidden");
    els.r1145FileInfo.innerHTML = "";
    els.supplierBanner.classList.add("hidden");
    els.previewCard.classList.add("hidden");
    els.hdrModal.classList.add("hidden");
    companyPicker.reset();
    clearStatus();
    setGenerateReady("empty");
  }

  // ---------- Language picker + header modal ----------
  function renderHeaderBox() {
    const useNewPriceCol = !!els.newPriceCheckbox.checked;
    const list = headerDisplayList(state.lang, useNewPriceCol);
    // Tab-separated line, with required fields wrapped in <span class="req">
    els.hdrBox.innerHTML = list
      .map((h) => (h.required ? `<span class="req">${escapeHtml(h.label)}</span>` : escapeHtml(h.label)))
      .join("\t");
  }

  // ---------- Supplier banner ----------
  function initialsFromName(name) {
    if (!name) return "—";
    const words = name.trim().split(/\s+/).slice(0, 2);
    return words.map((w) => w[0] || "").join("").toUpperCase() || "—";
  }

  function showSupplierBanner(supplier, division) {
    // Hide the banner entirely when nothing was detected in the header rows.
    // Hotels sometimes ship files with missing or messed-up supplier lines —
    // that's OK, parsing still works, user just fills in the number manually.
    if (!supplier && !division) {
      els.supplierBanner.classList.add("hidden");
      return;
    }

    // Prefer supplier as the "primary" name displayed. If only division was
    // detected, surface that instead of a blank banner so the user still gets
    // some context about which hotel/supplier this file relates to.
    if (supplier && supplier.name) {
      els.supplierName.textContent = supplier.name;
      els.supplierIcon.textContent = initialsFromName(supplier.name);
    } else if (division && division.name) {
      els.supplierName.textContent = "(supplier name not detected)";
      els.supplierIcon.textContent = initialsFromName(division.name);
    } else {
      els.supplierName.textContent = "(partial info detected)";
      els.supplierIcon.textContent = "?";
    }

    const bits = [];
    if (supplier?.num) bits.push(`Supplier no. ${supplier.num}`);
    if (division) {
      const divLabel =
        division.num && division.name ? `${division.name} (Division ${division.num})` :
        division.name ? division.name :
        division.num ? `Division ${division.num}` : "";
      if (divLabel) bits.push(divLabel);
    }
    if (!bits.length) bits.push("Enter the supplier number manually below.");
    els.supplierMeta.textContent = bits.join(" · ");

    els.supplierBanner.classList.remove("hidden");

    // Auto-fill the supplier number param only when we have a clean 6-digit number.
    if (supplier?.num && !els.supplierNo.value && /^\d{6}$/.test(supplier.num)) {
      els.supplierNo.value = supplier.num;
    }
  }

  // ---------- File handling ----------
  async function handleP2PFile(file) {
    clearStatus();
    try {
      const data = await file.arrayBuffer();
      const parsed = await runWithLoading(
        "Parsing P2P file…",
        "Detecting supplier and headers.",
        () => {
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          // Try each sheet in order — some workbooks have empty or unrelated
          // leftover sheets before the real P2P data. parseFirstParseableSheet
          // picks the first one that parses cleanly.
          const sheets = wb.SheetNames.map((name) => ({
            name,
            aoa: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: true }),
          }));
          return parseFirstParseableSheet(sheets, {
            useNewPriceCol: !!els.newPriceCheckbox.checked,
          });
        }
      );

      if (parsed.rows.length === 0) throw new Error("No data rows found in the P2P file.");
      state.p2pParsed = parsed;

      showSupplierBanner(parsed.supplier, parsed.division);

      // When multiple sheets exist, let the user know which one we picked so
      // they can tell if we guessed wrong and need to reorder their workbook.
      const sheetNote = parsed.totalSheets > 1
        ? ` · sheet "${escapeHtml(parsed.sheetName)}" of ${parsed.totalSheets}`
        : "";

      els.fileInfo.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
        <span class="name">${escapeHtml(file.name)}</span>
        <span class="size">· ${formatBytes(file.size)} · ${parsed.rows.length} items (counted by Article No.)${sheetNote}</span>
      `;
      els.fileInfo.classList.remove("hidden");

      if (!els.validityDate.value) els.validityDate.value = todayDDMMYYYY();
      await tryMerge();
    } catch (err) {
      console.error(err);
      state.p2pParsed = null;
      els.fileInfo.classList.add("hidden");
      els.supplierBanner.classList.add("hidden");
      setStatus("error", `<h3>Could not read the P2P file</h3>${escapeHtml(err.message || String(err))}`);
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
    if (!state.p2pParsed || !state.r1145Rows) {
      // Not both files yet — keep the button disabled.
      const missing = [];
      if (!state.p2pParsed) missing.push("P2P file");
      if (!state.r1145Rows) missing.push("Report 1145 file");
      setGenerateReady("empty", `Upload the ${missing.join(" and ")} to continue.`);
      return;
    }

    const opts = {
      useNewPriceCol: !!els.newPriceCheckbox.checked,
      openLeadTime:   !!els.openLeadCheckbox.checked,
    };

    const { rows, summary } = await runWithLoading(
      "Merging P2P with Report 1145…",
      `Matching ${state.p2pParsed.rows.length.toLocaleString()} P2P rows against ${state.r1145Rows.length.toLocaleString()} Report 1145 articles.`,
      () => mergeP2PAndReport1145(state.r1145Rows, state.p2pParsed.rows, opts),
    );

    state.mergedRows = rows;
    state.summary = summary;

    // Re-run validation on the merged rows with dummy params (row-level only).
    const { invalidCells, errors, warnings } = await runWithLoading(
      "Validating merged data…",
      `Checking ${rows.length.toLocaleString()} row${rows.length === 1 ? "" : "s"} against all rules.`,
      () => validate(rows, { companyId: "000", supplierNo: "000000", language: "TH", validityDate: "01012026" })
    );
    state.invalidCells = invalidCells;

    renderPreview(rows, invalidCells);

    // Summary + status
    const summaryHtml =
      `<h3>Merge complete</h3>` +
      `<ul>` +
      `  <li><strong>${summary.matched}</strong> row${summary.matched === 1 ? "" : "s"} updated with P2P data</li>` +
      `  <li><strong>${summary.p2pOnly}</strong> P2P-only item${summary.p2pOnly === 1 ? "" : "s"} (not in Report 1145 — will need manual enrichment)</li>` +
      `  <li><strong>${summary.r1145Only}</strong> Report 1145 item${summary.r1145Only === 1 ? "" : "s"} appended with "No price update"</li>` +
      `</ul>`;

    if (errors.length) {
      const list = errors.map((e) => `<li>${escapeHtml(e).replace(/\n/g, "<br>")}</li>`).join("");
      const warnList = warnings.length
        ? `<p style="margin-top:10px"><strong>Warnings:</strong></p><ul>${warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("")}</ul>`
        : "";
      setStatus("error",
        summaryHtml +
        `<p style="margin-top:10px"><strong>Validation failed — ${errors.length} issue${errors.length === 1 ? "" : "s"}:</strong></p><ul>${list}</ul>${warnList}`
      );
      setGenerateReady("error", errors.length);
      return;
    }

    if (warnings.length) {
      const list = warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("");
      setStatus("warn", summaryHtml + `<p style="margin-top:10px"><strong>Warnings:</strong></p><ul>${list}</ul>`);
    } else {
      setStatus("success", summaryHtml + `<p style="margin-top:6px">All rows validated successfully.</p>`);
    }
    setGenerateReady("ready");
  }

  // ---------- Preview ----------
  const STATUS_PILL = {
    "":                         { label: "—",                    cls: "pill-none" },
    "Open lead time":           { label: "Open lead time",       cls: "pill-open" },
    "Price from report 1145":   { label: "Price from Report 1145", cls: "pill-p1145" },
    "No price update":          { label: "No price update",      cls: "pill-nopu" },
    "P2P-only item":            { label: "P2P-only item",        cls: "pill-p2ponly" },
  };

  const PREVIEW_COLS = [
    { key: "pos",          label: "#",         cls: "c-pos" },
    { key: "itemNo",       label: "Item",      cls: "c-item" },
    { key: "descGB",       label: "Description", cls: "c-en" },
    { key: "ou",           label: "OU",        cls: "c-unit" },
    { key: "__oldPrice",   label: "Old",       cls: "c-price" },
    { key: "priceOU",      label: "New",       cls: "c-price" },
    { key: "__diff",       label: "Diff",      cls: "c-cuou" },
    { key: "availability", label: "LT",        cls: "c-avail" },
    { key: "status",       label: "Status",    cls: "c-cust" },
  ];

  // For the full-data modal — ALL fields + the P2P-specific columns
  const FULL_COLS = [
    { key: "pos", label: "#" },
    { key: "itemNo", label: "Article no." },
    { key: "ean", label: "EAN / GTIN" },
    { key: "descGB", label: "Description (EN)" },
    { key: "descExtra", label: "Description (Local)" },
    { key: "ou", label: "OU" },
    { key: "cu", label: "CU" },
    { key: "cuou", label: "CU/OU" },
    { key: "__oldPrice", label: "Old Price" },
    { key: "priceOU", label: "New Price" },
    { key: "__diff", label: "Diff" },
    { key: "origin", label: "Origin" },
    { key: "customsNo", label: "Customs No" },
    { key: "availability", label: "Lead time" },
    { key: "customerId", label: "Customer ID" },
    {
      key: "status",
      label: "Status",
      cellHtml: (r) => {
        const info = STATUS_PILL[r.status || ""] || STATUS_PILL[""];
        return `<span class="status-pill ${info.cls}">${escapeHtml(info.label)}</span>`;
      },
    },
  ];

  function fmtMoney(v) {
    if (v === "" || v === null || v === undefined) return "—";
    if (typeof v !== "number") return String(v);
    return v.toFixed(2);
  }

  function diffFor(r) {
    const a = typeof r.__oldPrice === "number" ? r.__oldPrice : null;
    const b = typeof r.priceOU === "number" ? r.priceOU : null;
    if (a === null || b === null) return "—";
    const d = b - a;
    if (Math.abs(d) < 0.005) return "0.00";
    return (d > 0 ? "+" : "") + d.toFixed(2);
  }

  // Map the empty-status pill to its preview label once, so the dropdown
  // can refer to it consistently. VBA flows tag normal matched rows with "".
  const NORMAL_LABEL = "Normal (no change)";

  // Rebuild the status-filter dropdown from the set of statuses present in
  // the current merged data. Only statuses that actually appear are listed,
  // each with a live count. Preserves the user's selection when possible.
  function updateStatusFilterOptions(rows) {
    const counts = new Map(); // statusValue -> count
    for (const r of rows) {
      const key = r.status || "";
      counts.set(key, (counts.get(key) || 0) + 1);
    }

    // Build option list, preferring a stable order: Normal first, then the
    // others in the order they appear in the STATUS_PILL map.
    const orderedKeys = [];
    for (const k of Object.keys(STATUS_PILL)) {
      if (counts.has(k)) orderedKeys.push(k);
    }
    // Surface any unknown statuses we didn't predefine (future-proofing)
    for (const k of counts.keys()) {
      if (!orderedKeys.includes(k)) orderedKeys.push(k);
    }

    const prevValue = state.statusFilter;
    const opts = [`<option value="__all__">All statuses (${rows.length})</option>`];
    for (const k of orderedKeys) {
      const label = (STATUS_PILL[k]?.label === "—" ? NORMAL_LABEL : STATUS_PILL[k]?.label) || k || NORMAL_LABEL;
      opts.push(`<option value="${escapeHtml(k)}">${escapeHtml(label)} (${counts.get(k)})</option>`);
    }
    els.statusFilter.innerHTML = opts.join("");

    // Restore previous selection if it still exists; otherwise reset to All.
    const valid = prevValue === "__all__" || counts.has(prevValue);
    state.statusFilter = valid ? prevValue : "__all__";
    els.statusFilter.value = state.statusFilter;
  }

  function renderPreview(rows, invalidCells = new Map()) {
    // Inject the diff into each row for preview/modal lookup
    rows.forEach((r) => { r.__diff = diffFor(r); });

    // Refresh the dropdown options to match the current dataset
    updateStatusFilterOptions(rows);

    // Apply the active filter — but keep the ORIGINAL row index for
    // invalidCells lookup, which is indexed against state.mergedRows.
    const filtered = state.statusFilter === "__all__"
      ? rows.map((r, idx) => ({ r, idx }))
      : rows.map((r, idx) => ({ r, idx })).filter(({ r }) => (r.status || "") === state.statusFilter);

    const showRows = filtered.slice(0, 200);
    const head = `<thead><tr>${PREVIEW_COLS.map((c) => `<th class="${c.cls}">${c.label}</th>`).join("")}</tr></thead>`;
    const body = "<tbody>" + showRows.map(({ r, idx }) => {
      const invalid = invalidCells.get(idx) || new Set();
      return "<tr>" + PREVIEW_COLS.map((c) => {
        let shown;
        if (c.key === "__oldPrice" || c.key === "priceOU") shown = fmtMoney(r[c.key]);
        else if (c.key === "__diff") shown = diffFor(r);
        else if (c.key === "status") {
          const info = STATUS_PILL[r.status || ""] || STATUS_PILL[""];
          return `<td class="${c.cls} status-cell"><span class="status-pill ${info.cls}">${escapeHtml(info.label)}</span></td>`;
        }
        else {
          const v = r[c.key];
          shown = v === NA_MARKER ? "#N/A" : (v === null || v === undefined ? "" : String(v));
        }

        // Paint the New-price cell with highlight colors based on status
        let extraCls = "";
        if (c.key === "priceOU") {
          if (r.status === "Price from report 1145") extraCls = " price-fallback";
          else if (r.status === "Open lead time" || (r.status === "" && r.__source === "matched")) extraCls = " price-updated";
        }

        const invalidCls = invalid.has(c.key) ? " invalid-cell" : "";
        return `<td class="${c.cls}${invalidCls}${extraCls}" title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
      }).join("") + "</tr>";
    }).join("") + "</tbody>";
    els.previewTable.innerHTML = head + body;

    // Summary reflects the current filter
    const filterActive = state.statusFilter !== "__all__";
    if (filterActive) {
      const label = (STATUS_PILL[state.statusFilter]?.label === "—" ? NORMAL_LABEL : STATUS_PILL[state.statusFilter]?.label) || state.statusFilter;
      els.previewSummary.textContent = filtered.length > 200
        ? `Filter: "${label}" — showing first 200 of ${filtered.length} matching rows (of ${rows.length} total).`
        : `Filter: "${label}" — ${filtered.length} matching row${filtered.length === 1 ? "" : "s"} (of ${rows.length} total).`;
    } else {
      els.previewSummary.textContent = rows.length > 200
        ? `Showing first 200 of ${rows.length} rows — click "Show Full Data" to see all.`
        : `Showing all ${rows.length} row${rows.length === 1 ? "" : "s"}.`;
    }
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

    // Re-validate with real params so supplier-no/validity-date shape checks fire.
    const validationParams = getParams(companies[0] || "000");
    const { errors, warnings } = validate(state.mergedRows, validationParams);
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
        downloadBlob("\uFEFF" + xml, filename, "application/xml;charset=utf-8");
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
  // P2P drop zone
  els.dropZone.addEventListener("click", () => els.fileInput.click());
  els.dropZone.addEventListener("dragover", (e) => { e.preventDefault(); els.dropZone.classList.add("dragging"); });
  els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragging"));
  els.dropZone.addEventListener("drop", (e) => {
    e.preventDefault(); els.dropZone.classList.remove("dragging");
    const f = e.dataTransfer.files[0]; if (f) handleP2PFile(f);
  });
  els.fileInput.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) handleP2PFile(f); });

  // R1145 drop zone
  els.r1145Drop.addEventListener("click", () => els.r1145FileInput.click());
  els.r1145Drop.addEventListener("dragover", (e) => { e.preventDefault(); els.r1145Drop.classList.add("dragging"); });
  els.r1145Drop.addEventListener("dragleave", () => els.r1145Drop.classList.remove("dragging"));
  els.r1145Drop.addEventListener("drop", (e) => {
    e.preventDefault(); els.r1145Drop.classList.remove("dragging");
    const f = e.dataTransfer.files[0]; if (f) handleR1145File(f);
  });
  els.r1145FileInput.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) handleR1145File(f); });

  // Language picker
  els.langPick.querySelectorAll(".lang-btn").forEach((b) => {
    b.addEventListener("click", () => {
      els.langPick.querySelectorAll(".lang-btn").forEach((x) => x.classList.toggle("active", x === b));
      state.lang = b.dataset.lang === "VN" ? "VN" : "EN";
      renderHeaderBox();
      // If P2P file is already loaded, re-parse it with the new language
      // so the user can switch without re-uploading.
      if (els.fileInput.files && els.fileInput.files[0]) {
        handleP2PFile(els.fileInput.files[0]);
      }
    });
  });

  // Header-example modal
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
    } catch (_) {
      // No clipboard access — do nothing loud; user can still copy manually.
    }
  });

  // Toggle changes trigger re-merge (and re-parse if needed, because NEW PRICE
  // is a required column when the toggle is on).
  els.newPriceCheckbox.addEventListener("change", async () => {
    renderHeaderBox();
    if (els.fileInput.files && els.fileInput.files[0]) {
      await handleP2PFile(els.fileInput.files[0]);
    } else {
      await tryMerge();
    }
  });
  els.openLeadCheckbox.addEventListener("change", () => { tryMerge(); });

  // Params & generate
  els.generateBtn.addEventListener("click", runGenerate);
  els.resetBtn.addEventListener("click", resetAll);

  els.showFullBtn.addEventListener("click", () => {
    openFullDataModal({
      rows: state.mergedRows,
      columns: FULL_COLS,
      invalidCells: state.invalidCells,
      statusPillMap: STATUS_PILL,
    });
  });

  // Status filter — re-renders the preview using the currently-selected status.
  // Only options present in the current round are shown (see updateStatusFilterOptions).
  els.statusFilter.addEventListener("change", (e) => {
    state.statusFilter = e.target.value;
    if (state.mergedRows.length) {
      renderPreview(state.mergedRows, state.invalidCells);
    }
  });

  restrictToDigits(els.supplierNo);
  restrictToDigits(els.validityDate);
  els.validityDate.placeholder = todayDDMMYYYY();
  renderHeaderBox();
  setGenerateReady("empty");
}
