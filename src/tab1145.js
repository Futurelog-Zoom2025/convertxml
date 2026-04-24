// Tab 1: Report 1145 → XML converter.
// Ported from the monolithic main.js. Exposes initR1145Tab() which wires up
// all event listeners for the DOM ids inside #panel1145.

import * as XLSX from "xlsx";
import { parseReport1145, NA_MARKER } from "./reportParser.js";
import { validate } from "./validator.js";
import { generateXml } from "./xmlGenerator.js";
import { downloadTemplate } from "./templateGenerator.js";
import {
  $, escapeHtml, formatBytes, todayDDMMYYYY, delay,
  runWithLoading, downloadBlob, buildCompanyMultiselect, restrictToDigits,
} from "./shared.js";
import { openFullDataModal } from "./fullDataModal.js";

export function initR1145Tab() {
  const els = {
    dropZone: $("#dropZone"),
    fileInput: $("#fileInput"),
    fileInfo: $("#fileInfo"),
    supplierNo: $("#supplierNo"),
    language: $("#language"),
    validityDate: $("#validityDate"),
    generateBtn: $("#generateBtn"),
    resetBtn: $("#resetBtn"),
    templateBtn: $("#templateBtn"),
    genHint: $("#genHint"),
    status: $("#status"),
    previewCard: $("#previewCard"),
    previewSummary: $("#previewSummary"),
    previewTable: $("#previewTable"),
    showFullBtn: $("#showFullBtn"),
  };

  const companyPicker = buildCompanyMultiselect({
    rootId: "companyMultiselect",
    btnId: "companyBtn",
    labelId: "companyBtnLabel",
    menuId: "companyMenu",
    optionsId: "companyOptions",
    selectAllId: "selectAllCompanies",
    clearAllId: "clearAllCompanies",
  });

  const state = {
    rows: [],
    fileName: null,
    invalidCells: new Map(),
  };

  // ---------- Status + generate-button gating ----------
  function setStatus(kind, html) {
    els.status.className = `status ${kind}`;
    els.status.innerHTML = html;
    els.status.classList.remove("hidden");
  }
  function clearStatus() {
    els.status.className = "status hidden";
    els.status.innerHTML = "";
  }

  function setGenerateReady(kind, detail) {
    if (kind === "empty") {
      els.generateBtn.disabled = true;
      els.genHint.textContent = "Upload a file first.";
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
    state.rows = [];
    state.fileName = null;
    state.invalidCells = new Map();
    els.fileInput.value = "";
    els.fileInfo.classList.add("hidden");
    els.fileInfo.innerHTML = "";
    els.previewCard.classList.add("hidden");
    companyPicker.reset();
    setGenerateReady("empty");
    clearStatus();
  }

  function getParams(companyId) {
    return {
      companyId,
      supplierNo: els.supplierNo.value.trim(),
      language: els.language.value.trim(),
      validityDate: els.validityDate.value.trim(),
    };
  }

  // ---------- Preview table ----------
  const PREVIEW_COLS = [
    { key: "pos",          label: "#",       cls: "c-pos" },
    { key: "itemNo",       label: "Item",    cls: "c-item" },
    { key: "descDE",       label: "German",  cls: "c-de" },
    { key: "descGB",       label: "English", cls: "c-en" },
    { key: "descExtra",    label: "Local",   cls: "c-local" },
    { key: "ean",          label: "EAN",     cls: "c-ean" },
    { key: "ou",           label: "OU",      cls: "c-unit" },
    { key: "cu",           label: "CU",      cls: "c-unit" },
    { key: "cuou",         label: "CU/OU",   cls: "c-cuou" },
    { key: "priceOU",      label: "Price",   cls: "c-price" },
    { key: "origin",       label: "Orig",    cls: "c-origin" },
    { key: "availability", label: "Avail",   cls: "c-avail" },
    { key: "customerId",   label: "Cust",    cls: "c-cust" },
  ];

  const FULL_COLS = [
    { key: "pos",          label: "#" },
    { key: "itemNo",       label: "Article no." },
    { key: "ean",          label: "EAN / GTIN" },
    { key: "manArtId",     label: "Mfg Item No" },
    { key: "descDE",       label: "Name (DE)" },
    { key: "descFR",       label: "Name (FR)" },
    { key: "descIT",       label: "Name (IT)" },
    { key: "descGB",       label: "Name (GB)" },
    { key: "descExtra",    label: "Name (Local)" },
    { key: "ou",           label: "OU" },
    { key: "cu",           label: "CU" },
    { key: "cuou",         label: "CU/OU" },
    { key: "priceOU",      label: "Price" },
    { key: "origin",       label: "Origin" },
    { key: "customsNo",    label: "Customs No" },
    { key: "leadTimeRaw",  label: "Lead Time" },
    { key: "availability", label: "Availability" },
    { key: "specUrl",      label: "Spec URL" },
    { key: "offerStart",   label: "Offer Start" },
    { key: "offerEnd",     label: "Offer End" },
    { key: "customerId",   label: "Customer ID" },
  ];

  function displayValue(v) {
    if (v === NA_MARKER) return "#N/A";
    if (v === null || v === undefined) return "";
    if (typeof v === "number") return Number.isInteger(v) ? String(v) : v.toFixed(2);
    return String(v);
  }

  function renderPreview(rows, invalidCells = new Map()) {
    const showRows = rows.slice(0, 200);
    const head = `<thead><tr>${PREVIEW_COLS.map((c) => `<th class="${c.cls}">${c.label}</th>`).join("")}</tr></thead>`;
    const body = "<tbody>" + showRows.map((r, idx) => {
      const invalid = invalidCells.get(idx) || new Set();
      return "<tr>" + PREVIEW_COLS.map((c) => {
        const v = r[c.key];
        const shown = displayValue(v);
        const isNA = v === NA_MARKER;
        const invalidCls = invalid.has(c.key) ? (isNA ? " invalid-cell na-cell" : " invalid-cell") : "";
        return `<td class="${c.cls}${invalidCls}" title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
      }).join("") + "</tr>";
    }).join("") + "</tbody>";
    els.previewTable.innerHTML = head + body;

    els.previewSummary.textContent = rows.length > 200
      ? `Showing first 200 of ${rows.length} rows — click "Show Full Data" to see all.`
      : `Showing all ${rows.length} row${rows.length === 1 ? "" : "s"}.`;
    els.previewCard.classList.remove("hidden");
  }

  // ---------- File handling ----------
  async function handleFile(file) {
    clearStatus();
    state.fileName = file.name;
    try {
      const data = await file.arrayBuffer();
      const rows = await runWithLoading(
        "Parsing Excel file…",
        "This may take a moment for large files.",
        () => {
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          const firstSheet = wb.SheetNames[0];
          const ws = wb.Sheets[firstSheet];
          const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
          return parseReport1145(aoa);
        }
      );

      if (rows.length === 0) throw new Error("No data rows found below the header in this file.");
      state.rows = rows;

      els.fileInfo.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
        <span class="name">${escapeHtml(file.name)}</span>
        <span class="size">· ${formatBytes(file.size)} · ${rows.length} article${rows.length === 1 ? "" : "s"}</span>
      `;
      els.fileInfo.classList.remove("hidden");
      renderPreview(rows);

      if (!els.validityDate.value) els.validityDate.value = todayDDMMYYYY();
      await runValidation(true);
    } catch (err) {
      console.error(err);
      state.rows = [];
      els.fileInfo.classList.add("hidden");
      els.previewCard.classList.add("hidden");
      setGenerateReady("empty");
      setStatus("error", `<h3>Could not read the file</h3>${escapeHtml(err.message || String(err))}`);
    }
  }

  // ---------- Validation ----------
  async function runValidation(auto = false) {
    const params = auto
      ? { companyId: "000", supplierNo: "000000", language: "TH", validityDate: "01012026" }
      : getParams(companyPicker.getSelected()[0] || "000");

    const result = await runWithLoading(
      "Validating…",
      `Checking ${state.rows.length.toLocaleString()} row${state.rows.length === 1 ? "" : "s"} against all rules.`,
      () => {
        const v = validate(state.rows, params);
        renderPreview(state.rows, v.invalidCells);
        state.invalidCells = v.invalidCells;
        return v;
      }
    );
    const { errors, warnings } = result;

    if (errors.length) {
      const list = errors.map((e) => `<li>${escapeHtml(e).replace(/\n/g, "<br>")}</li>`).join("");
      const warnList = warnings.length
        ? `<p style="margin-top:10px"><strong>Warnings:</strong></p><ul>${warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("")}</ul>`
        : "";
      setStatus("error", `<h3>Validation failed — ${errors.length} issue${errors.length === 1 ? "" : "s"}</h3><ul>${list}</ul>${warnList}`);
      setGenerateReady("error", errors.length);
      return false;
    }

    if (warnings.length) {
      const list = warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("");
      setStatus("warn", `<h3>Validation passed with warnings</h3><ul>${list}</ul>`);
    } else {
      setStatus("success", `<h3>Validation passed</h3>All ${state.rows.length} rows look good.`);
    }
    setGenerateReady("ready");
    return true;
  }

  // ---------- Generate ----------
  async function runGenerate() {
    const companies = companyPicker.getSelected();
    if (companies.length === 0) {
      setStatus("error", `<h3>Select a Company ID</h3>Please select at least one WebShop Company ID.`);
      return;
    }
    const ok = await runValidation(false);
    if (!ok) return;

    const createdFiles = [];
    try {
      for (let i = 0; i < companies.length; i++) {
        const companyId = companies[i];
        const params = getParams(companyId);
        const result = await runWithLoading(
          `Generating XML ${i + 1} of ${companies.length}…`,
          `Company ${companyId} · ${state.rows.length.toLocaleString()} article${state.rows.length === 1 ? "" : "s"}`,
          () => generateXml(state.rows, params)
        );
        const { xml, filename } = result;
        downloadBlob("\uFEFF" + xml, filename, "application/xml;charset=utf-8");
        createdFiles.push(filename);
        if (i < companies.length - 1) await delay(350);
      }
      const fileList = createdFiles.map((f) => `<li><code>${escapeHtml(f)}</code></li>`).join("");
      setStatus("success",
        `<h3>XML generated for ${createdFiles.length} compan${createdFiles.length === 1 ? "y" : "ies"}</h3>` +
        `<ul>${fileList}</ul>` +
        `<p class="muted small" style="margin-top:8px">If your browser only downloaded one file, check its download settings — it may be blocking multiple automatic downloads.</p>`
      );
    } catch (err) {
      console.error(err);
      setStatus("error", `<h3>Generation failed</h3>${escapeHtml(err.message || String(err))}`);
    }
  }

  // ---------- Event wiring ----------
  els.dropZone.addEventListener("click", () => els.fileInput.click());
  els.dropZone.addEventListener("dragover", (e) => { e.preventDefault(); els.dropZone.classList.add("dragging"); });
  els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragging"));
  els.dropZone.addEventListener("drop", (e) => {
    e.preventDefault(); els.dropZone.classList.remove("dragging");
    const f = e.dataTransfer.files[0]; if (f) handleFile(f);
  });
  els.fileInput.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) handleFile(f); });

  els.generateBtn.addEventListener("click", runGenerate);
  els.resetBtn.addEventListener("click", resetAll);
  els.templateBtn.addEventListener("click", (e) => {
    e.preventDefault();
    try {
      downloadTemplate();
      setStatus("success", `<h3>Template downloaded</h3>Open <code>Report_1145_Template.xlsx</code>, fill in rows starting from row 5, then drop it onto the upload area above.`);
    } catch (err) {
      console.error(err);
      setStatus("error", `<h3>Could not create template</h3>${escapeHtml(err.message || String(err))}`);
    }
  });

  els.showFullBtn.addEventListener("click", () => {
    openFullDataModal({
      rows: state.rows,
      columns: FULL_COLS,
      invalidCells: state.invalidCells,
    });
  });

  restrictToDigits(els.supplierNo);
  restrictToDigits(els.validityDate);
  els.validityDate.placeholder = todayDDMMYYYY();
  setGenerateReady("empty");
}
