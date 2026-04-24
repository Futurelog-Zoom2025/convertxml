import * as XLSX from "xlsx";
import { parseReport1145, NA_MARKER } from "./reportParser.js";
import { validate } from "./validator.js";
import { generateXml } from "./xmlGenerator.js";
import { downloadTemplate } from "./templateGenerator.js";

const $ = (sel) => document.querySelector(sel);

// The exact list and order requested by the business.
// Duplicates in the source list (230 appeared twice) are deduped here.
const COMPANY_IDS = [
  "169", "215", "233", "247", "278", "257", "262", "230", "315",
  "101", "265", "225", "296", "285",
];

const els = {
  dropZone: $("#dropZone"),
  fileInput: $("#fileInput"),
  fileInfo: $("#fileInfo"),
  paramsCard: $("#paramsCard"),
  actionCard: $("#actionCard"),
  // Multi-select
  companyMultiselect: $("#companyMultiselect"),
  companyBtn: $("#companyBtn"),
  companyBtnLabel: $("#companyBtnLabel"),
  companyMenu: $("#companyMenu"),
  companyOptions: $("#companyOptions"),
  selectAllCompanies: $("#selectAllCompanies"),
  clearAllCompanies: $("#clearAllCompanies"),
  // Other params
  supplierNo: $("#supplierNo"),
  language: $("#language"),
  validityDate: $("#validityDate"),
  // Actions
  validateBtn: $("#validateBtn"),
  generateBtn: $("#generateBtn"),
  resetBtn: $("#resetBtn"),
  templateBtn: $("#templateBtn"),
  status: $("#status"),
  // Preview
  previewCard: $("#previewCard"),
  previewSummary: $("#previewSummary"),
  previewTable: $("#previewTable"),
  showFullBtn: $("#showFullBtn"),
  // Full-data modal
  fullDataModal: $("#fullDataModal"),
  fullDataTable: $("#fullDataTable"),
  fullDataSummary: $("#fullDataSummary"),
  fullDataSearch: $("#fullDataSearch"),
  closeFullBtn: $("#closeFullBtn"),
  // Loading
  loadingOverlay: $("#loadingOverlay"),
  loadingMsg: $("#loadingMsg"),
  loadingSub: $("#loadingSub"),
};

// -------- Loading overlay helpers --------
function showLoading(msg, sub) {
  if (msg) els.loadingMsg.textContent = msg;
  if (sub !== undefined) els.loadingSub.textContent = sub;
  els.loadingOverlay.classList.remove("hidden");
}
function hideLoading() {
  els.loadingOverlay.classList.add("hidden");
}
// Double rAF ensures overlay paints before heavy synchronous work
function runWithLoading(msg, sub, fn) {
  return new Promise((resolve, reject) => {
    showLoading(msg, sub);
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        try {
          const result = fn();
          hideLoading();
          resolve(result);
        } catch (err) {
          hideLoading();
          reject(err);
        }
      });
    });
  });
}

// In-memory state
const state = {
  rows: [],
  fileName: null,
  selectedCompanies: new Set(),
};

// --------------- Helpers ---------------

function todayDDMMYYYY() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}${mm}${d.getFullYear()}`;
}

function setStatus(kind, html) {
  els.status.className = `status ${kind}`;
  els.status.innerHTML = html;
  els.status.classList.remove("hidden");
}

function clearStatus() {
  els.status.className = "status hidden";
  els.status.innerHTML = "";
}

function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

function enableStep(n) {
  if (n >= 2) els.paramsCard.setAttribute("aria-disabled", "false");
  if (n >= 3) {
    els.actionCard.setAttribute("aria-disabled", "false");
    els.validateBtn.disabled = false;
    els.generateBtn.disabled = false;
  }
}

function resetAll() {
  state.rows = [];
  state.fileName = null;
  state.selectedCompanies.clear();
  els.fileInput.value = "";
  els.fileInfo.classList.add("hidden");
  els.fileInfo.innerHTML = "";
  els.paramsCard.setAttribute("aria-disabled", "true");
  els.actionCard.setAttribute("aria-disabled", "true");
  els.validateBtn.disabled = true;
  els.generateBtn.disabled = true;
  els.previewCard.classList.add("hidden");
  // Uncheck all company checkboxes
  els.companyOptions.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; });
  updateCompanyLabel();
  clearStatus();
}

function getSelectedCompanies() {
  return Array.from(state.selectedCompanies);
}

function getParams(companyId) {
  return {
    companyId,
    supplierNo: els.supplierNo.value.trim(),
    language: els.language.value.trim(),
    validityDate: els.validityDate.value.trim(),
  };
}

// --------------- Multi-select Company ID ---------------

function buildCompanyOptions() {
  const html = COMPANY_IDS.map((id) => `
    <label class="multiselect-option">
      <input type="checkbox" value="${id}" />
      <span class="company-code">${id}</span>
    </label>
  `).join("");
  els.companyOptions.innerHTML = html;

  els.companyOptions.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.addEventListener("change", () => {
      if (cb.checked) state.selectedCompanies.add(cb.value);
      else state.selectedCompanies.delete(cb.value);
      updateCompanyLabel();
    });
  });
}

function updateCompanyLabel() {
  const selected = getSelectedCompanies();
  if (selected.length === 0) {
    els.companyBtnLabel.textContent = "Select companies…";
    els.companyBtnLabel.classList.remove("has-selection");
    els.companyBtnLabel.classList.add("muted");
  } else if (selected.length <= 4) {
    els.companyBtnLabel.textContent = selected.join(", ");
    els.companyBtnLabel.classList.add("has-selection");
    els.companyBtnLabel.classList.remove("muted");
  } else {
    els.companyBtnLabel.textContent = `${selected.length} companies selected (${selected.slice(0, 3).join(", ")}…)`;
    els.companyBtnLabel.classList.add("has-selection");
    els.companyBtnLabel.classList.remove("muted");
  }
}

function toggleCompanyMenu(open) {
  const isOpen = !els.companyMenu.classList.contains("hidden");
  const shouldOpen = open === undefined ? !isOpen : open;
  if (shouldOpen) {
    els.companyMenu.classList.remove("hidden");
    els.companyMultiselect.classList.add("open");
  } else {
    els.companyMenu.classList.add("hidden");
    els.companyMultiselect.classList.remove("open");
  }
}

// --------------- File handling ---------------

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

    if (rows.length === 0) {
      throw new Error("No data rows found below the header in this file.");
    }
    state.rows = rows;

    els.fileInfo.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
      <span class="name">${escapeHtml(file.name)}</span>
      <span class="size">· ${formatBytes(file.size)} · ${rows.length} article${rows.length === 1 ? "" : "s"}</span>
    `;
    els.fileInfo.classList.remove("hidden");

    renderPreview(rows);
    enableStep(3);

    if (!els.validityDate.value) els.validityDate.value = todayDDMMYYYY();

    setStatus("info", `<strong>File parsed.</strong> ${rows.length} article${rows.length === 1 ? "" : "s"} loaded. Select one or more Company IDs, then click <em>Generate &amp; Download XML</em>.`);
  } catch (err) {
    console.error(err);
    state.rows = [];
    els.fileInfo.classList.add("hidden");
    els.previewCard.classList.add("hidden");
    setStatus("error", `<h3>Could not read the file</h3>${escapeHtml(err.message || String(err))}`);
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// --------------- Preview table (compact, fits card width) ---------------

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

// Columns shown in the full-data modal (every field)
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
  if (typeof v === "number") {
    return Number.isInteger(v) ? String(v) : v.toFixed(2);
  }
  return String(v);
}

function renderPreview(rows, invalidCells = new Map()) {
  const showRows = rows.slice(0, 200);
  const head = `<thead><tr>${PREVIEW_COLS.map((c) => `<th class="${c.cls}">${c.label}</th>`).join("")}</tr></thead>`;
  const body =
    "<tbody>" +
    showRows
      .map((r, idx) => {
        const invalid = invalidCells.get(idx) || new Set();
        return (
          "<tr>" +
          PREVIEW_COLS.map((c) => {
            const v = r[c.key];
            const shown = displayValue(v);
            const isNA = v === NA_MARKER;
            const invalidCls = invalid.has(c.key) ? (isNA ? " invalid-cell na-cell" : " invalid-cell") : "";
            return `<td class="${c.cls}${invalidCls}" title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
          }).join("") +
          "</tr>"
        );
      })
      .join("") +
    "</tbody>";
  els.previewTable.innerHTML = head + body;

  els.previewSummary.textContent =
    rows.length > 200
      ? `Showing first 200 of ${rows.length} rows — click "Show Full Data" to see all.`
      : `Showing all ${rows.length} row${rows.length === 1 ? "" : "s"}.`;
  els.previewCard.classList.remove("hidden");
}

// --------------- Full-data modal ---------------

let fullDataInvalidCells = new Map();

function renderFullData(filterText = "") {
  const q = filterText.trim().toLowerCase();
  const head = `<thead><tr>${FULL_COLS.map((c) => `<th>${escapeHtml(c.label)}</th>`).join("")}</tr></thead>`;

  let matchCount = 0;
  const body = "<tbody>" + state.rows
    .map((r, idx) => {
      if (q) {
        // Filter: include row if any field contains the query
        const hay = FULL_COLS.map((c) => displayValue(r[c.key])).join(" ").toLowerCase();
        if (!hay.includes(q)) return "";
      }
      matchCount++;
      const invalid = fullDataInvalidCells.get(idx) || new Set();
      return "<tr>" +
        FULL_COLS.map((c) => {
          const v = r[c.key];
          const shown = displayValue(v);
          const isNA = v === NA_MARKER;
          const cls = invalid.has(c.key) ? (isNA ? ' class="invalid-cell na-cell"' : ' class="invalid-cell"') : "";
          return `<td${cls} title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
        }).join("") +
        "</tr>";
    })
    .join("") + "</tbody>";

  els.fullDataTable.innerHTML = head + body;
  els.fullDataSummary.textContent = q
    ? `Showing ${matchCount} of ${state.rows.length} rows (filter: "${filterText}")`
    : `All ${state.rows.length} row${state.rows.length === 1 ? "" : "s"} · ${FULL_COLS.length} columns`;
}

function openFullDataModal() {
  runWithLoading(
    "Loading full data…",
    `Preparing ${state.rows.length.toLocaleString()} rows for display.`,
    () => {
      els.fullDataSearch.value = "";
      renderFullData("");
    }
  ).then(() => {
    els.fullDataModal.classList.remove("hidden");
    // Focus search for quick filtering
    setTimeout(() => els.fullDataSearch.focus(), 100);
  });
}

function closeFullDataModal() {
  els.fullDataModal.classList.add("hidden");
}

// --------------- Validate / Generate ---------------

async function runValidation(silent = false) {
  // Company ID only affects the output filename, not the row-level checks.
  // Use the first selected company if any, otherwise a placeholder so the
  // param-shape check still passes when the user runs "Validate Only" early.
  const selected = getSelectedCompanies();
  const companyForValidation = selected[0] || "000";
  const params = getParams(companyForValidation);

  const result = await runWithLoading(
    "Validating…",
    `Checking ${state.rows.length.toLocaleString()} row${state.rows.length === 1 ? "" : "s"} against all rules.`,
    () => {
      const v = validate(state.rows, params);
      renderPreview(state.rows, v.invalidCells);
      fullDataInvalidCells = v.invalidCells;
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
    return false;
  }

  if (warnings.length) {
    const list = warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("");
    setStatus("warn", `<h3>Validation passed with warnings</h3><ul>${list}</ul>`);
  } else if (!silent) {
    setStatus("success", `<h3>Validation passed</h3>All ${state.rows.length} rows look good.`);
  }
  return true;
}

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

// Small delay helper so browsers don't block rapid sequential downloads
function delay(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function runGenerate() {
  const companies = getSelectedCompanies();
  if (companies.length === 0) {
    setStatus("error", `<h3>Select a Company ID</h3>Please select at least one WebShop Company ID.`);
    return;
  }

  const ok = await runValidation(true);
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

      // Tiny gap so browsers don't dedupe or block consecutive downloads
      if (i < companies.length - 1) await delay(350);
    }

    const fileList = createdFiles.map((f) => `<li><code>${escapeHtml(f)}</code></li>`).join("");
    setStatus(
      "success",
      `<h3>XML generated for ${createdFiles.length} compan${createdFiles.length === 1 ? "y" : "ies"}</h3>` +
        `<ul>${fileList}</ul>` +
        `<p class="muted small" style="margin-top:8px">If your browser only downloaded one file, check its download settings — it may be blocking multiple automatic downloads.</p>`
    );
  } catch (err) {
    console.error(err);
    setStatus("error", `<h3>Generation failed</h3>${escapeHtml(err.message || String(err))}`);
  }
}

// --------------- Event wiring ---------------

// Drop zone
els.dropZone.addEventListener("click", () => els.fileInput.click());
els.dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  els.dropZone.classList.add("dragging");
});
els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragging"));
els.dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  els.dropZone.classList.remove("dragging");
  const f = e.dataTransfer.files[0];
  if (f) handleFile(f);
});
els.fileInput.addEventListener("change", (e) => {
  const f = e.target.files[0];
  if (f) handleFile(f);
});

// Multi-select dropdown
buildCompanyOptions();
els.companyBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  toggleCompanyMenu();
});
els.companyMenu.addEventListener("click", (e) => e.stopPropagation());
document.addEventListener("click", () => toggleCompanyMenu(false));
els.selectAllCompanies.addEventListener("click", () => {
  els.companyOptions.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.checked = true;
    state.selectedCompanies.add(cb.value);
  });
  updateCompanyLabel();
});
els.clearAllCompanies.addEventListener("click", () => {
  els.companyOptions.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.checked = false;
  });
  state.selectedCompanies.clear();
  updateCompanyLabel();
});

// Action buttons
els.validateBtn.addEventListener("click", () => runValidation(false));
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

// Full-data modal
els.showFullBtn.addEventListener("click", openFullDataModal);
els.closeFullBtn.addEventListener("click", closeFullDataModal);
els.fullDataModal.addEventListener("click", (e) => {
  if (e.target === els.fullDataModal) closeFullDataModal();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !els.fullDataModal.classList.contains("hidden")) {
    closeFullDataModal();
  }
});

// Debounced filter in full-data modal
let filterTimer = null;
els.fullDataSearch.addEventListener("input", (e) => {
  const q = e.target.value;
  clearTimeout(filterTimer);
  filterTimer = setTimeout(() => renderFullData(q), 120);
});

// Numeric-only filters for supplier/date fields
["supplierNo", "validityDate"].forEach((id) => {
  document.getElementById(id).addEventListener("input", (e) => {
    e.target.value = e.target.value.replace(/\D/g, "");
  });
});

// Prefill helpful defaults
els.validityDate.placeholder = todayDDMMYYYY();

// Unlock parameters panel immediately so users can fill them early
enableStep(2);
updateCompanyLabel();
