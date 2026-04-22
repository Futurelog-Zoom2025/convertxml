import * as XLSX from "xlsx";
import { parseReport1145, NA_MARKER } from "./reportParser.js";
import { validate } from "./validator.js";
import { generateXml } from "./xmlGenerator.js";
import { downloadTemplate } from "./templateGenerator.js";

const $ = (sel) => document.querySelector(sel);

const els = {
  dropZone: $("#dropZone"),
  fileInput: $("#fileInput"),
  fileInfo: $("#fileInfo"),
  paramsCard: $("#paramsCard"),
  actionCard: $("#actionCard"),
  companyId: $("#companyId"),
  supplierNo: $("#supplierNo"),
  language: $("#language"),
  validityDate: $("#validityDate"),
  validateBtn: $("#validateBtn"),
  generateBtn: $("#generateBtn"),
  resetBtn: $("#resetBtn"),
  templateBtn: $("#templateBtn"),
  status: $("#status"),
  previewCard: $("#previewCard"),
  previewSummary: $("#previewSummary"),
  previewTable: $("#previewTable"),
};

// In-memory state
const state = {
  rows: [],
  fileName: null,
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
  els.fileInput.value = "";
  els.fileInfo.classList.add("hidden");
  els.fileInfo.innerHTML = "";
  els.paramsCard.setAttribute("aria-disabled", "true");
  els.actionCard.setAttribute("aria-disabled", "true");
  els.validateBtn.disabled = true;
  els.generateBtn.disabled = true;
  els.previewCard.classList.add("hidden");
  clearStatus();
}

function getParams() {
  return {
    companyId: els.companyId.value.trim(),
    supplierNo: els.supplierNo.value.trim(),
    language: els.language.value.trim(),
    validityDate: els.validityDate.value.trim(),
  };
}

// --------------- File handling ---------------

async function handleFile(file) {
  clearStatus();
  state.fileName = file.name;

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array", cellDates: true });
    const firstSheet = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheet];
    // header:1 -> array of arrays; defval:"" to keep column alignment on empty cells
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    const rows = parseReport1145(aoa);
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

    // If user hasn't set a date yet, default to today
    if (!els.validityDate.value) els.validityDate.value = todayDDMMYYYY();

    setStatus("info", `<strong>File parsed.</strong> ${rows.length} article${rows.length === 1 ? "" : "s"} loaded. Fill in the XML parameters below, then click <em>Generate &amp; Download XML</em> in Step 3.`);
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

// --------------- Preview table ---------------

const PREVIEW_COLS = [
  { key: "pos", label: "#" },
  { key: "itemNo", label: "Item No" },
  { key: "descDE", label: "German" },
  { key: "descGB", label: "English" },
  { key: "descExtra", label: "Local" },
  { key: "ean", label: "EAN" },
  { key: "ou", label: "OU" },
  { key: "cu", label: "CU" },
  { key: "cuou", label: "CU/OU" },
  { key: "priceOU", label: "Price" },
  { key: "origin", label: "Origin" },
  { key: "availability", label: "Avail." },
  { key: "customerId", label: "Cust ID" },
];

function displayValue(v) {
  if (v === NA_MARKER) return "#N/A";
  if (v === null || v === undefined) return "";
  return String(v);
}

function renderPreview(rows, invalidCells = new Map()) {
  const showRows = rows.slice(0, 200);
  const head = `<thead><tr>${PREVIEW_COLS.map((c) => `<th>${c.label}</th>`).join("")}</tr></thead>`;
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
            const cls = invalid.has(c.key) ? (isNA ? ' class="invalid-cell na-cell"' : ' class="invalid-cell"') : "";
            return `<td${cls} title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
          }).join("") +
          "</tr>"
        );
      })
      .join("") +
    "</tbody>";
  els.previewTable.innerHTML = head + body;

  els.previewSummary.textContent =
    rows.length > 200
      ? `Showing first 200 of ${rows.length} rows.`
      : `Showing all ${rows.length} row${rows.length === 1 ? "" : "s"}.`;
  els.previewCard.classList.remove("hidden");
}

// --------------- Validate / Generate ---------------

function runValidation(silent = false) {
  const params = getParams();
  const { errors, warnings, invalidCells } = validate(state.rows, params);
  renderPreview(state.rows, invalidCells);

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

function runGenerate() {
  if (!runValidation(true)) return;
  const params = getParams();
  const { xml, filename } = generateXml(state.rows, params);
  // Prepend UTF-8 BOM so some downstream tools (and Excel) detect encoding correctly,
  // and keep the raw XML declaration intact.
  downloadBlob("\uFEFF" + xml, filename, "application/xml;charset=utf-8");
  setStatus("success",
    `<h3>XML generated successfully</h3>File: <code>${escapeHtml(filename)}</code> · ${state.rows.length} articles`
  );
}

// --------------- Event wiring ---------------

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

// Numeric-only filters for company/supplier/date fields
["companyId", "supplierNo", "validityDate"].forEach((id) => {
  document.getElementById(id).addEventListener("input", (e) => {
    e.target.value = e.target.value.replace(/\D/g, "");
  });
});

// Prefill helpful defaults
els.validityDate.placeholder = todayDDMMYYYY();

// On first load, Step 3 (params) is visually unlocked as soon as a file is chosen.
// But we allow user to fill params early, so unlock right away:
enableStep(2);
