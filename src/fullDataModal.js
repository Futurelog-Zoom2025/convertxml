// Shared full-data modal. Each tab calls openFullDataModal(...) with its own
// rows, column definitions, and invalidCells map. The modal supports:
//   - column-header sorting (asc → desc → none)
//   - text search across all columns
//   - "Errors only" toggle (visible only when there are any invalid cells)

import { escapeHtml, runWithLoading } from "./shared.js";
import { NA_MARKER } from "./reportParser.js";

const els = {
  modal:              document.getElementById("fullDataModal"),
  table:              document.getElementById("fullDataTable"),
  summary:            document.getElementById("fullDataSummary"),
  search:             document.getElementById("fullDataSearch"),
  closeBtn:           document.getElementById("closeFullBtn"),
  errorFilterLabel:   document.getElementById("errorFilterLabel"),
  errorFilterCheckbox:document.getElementById("errorFilterCheckbox"),
  errorRowCount:      document.getElementById("errorRowCount"),
};

// Per-open state — reset whenever a new dataset is shown
let view = {
  rows: [],
  columns: [],
  invalidCells: new Map(),
  sortCol: null,
  sortDir: null,
  showOnlyErrors: false,
  filter: "",
};

function displayValue(v) {
  if (v === NA_MARKER) return "#N/A";
  if (v === null || v === undefined) return "";
  if (typeof v === "number") {
    return Number.isInteger(v) ? String(v) : v.toFixed(2);
  }
  return String(v);
}

// Robust comparator for sorting — handles numbers, numeric-looking strings,
// plain text, empties/NA always to the bottom.
function compareValues(a, b) {
  const isEmpty = (x) => x === null || x === undefined || x === "" || x === NA_MARKER;
  const ae = isEmpty(a); const be = isEmpty(b);
  if (ae && be) return 0;
  if (ae) return 1;
  if (be) return -1;
  const toNum = (x) => {
    if (typeof x === "number") return x;
    const s = String(x).replace(/,/g, "").trim();
    if (s === "") return NaN;
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  };
  const an = toNum(a); const bn = toNum(b);
  if (!Number.isNaN(an) && !Number.isNaN(bn)) return an - bn;
  return String(a).toLowerCase().localeCompare(String(b).toLowerCase());
}

function render() {
  const q = view.filter.trim().toLowerCase();
  let entries = view.rows.map((r, idx) => ({ r, idx }));

  if (view.showOnlyErrors) {
    entries = entries.filter((e) => view.invalidCells.has(e.idx));
  }

  if (q) {
    entries = entries.filter((e) =>
      view.columns.map((c) => displayValue(e.r[c.key])).join(" ").toLowerCase().includes(q)
    );
  }

  if (view.sortCol && view.sortDir) {
    const key = view.sortCol;
    const mul = view.sortDir === "asc" ? 1 : -1;
    entries.sort((a, b) => mul * compareValues(a.r[key], b.r[key]));
  }

  const head = `<thead><tr>${view.columns.map((c) => {
    const isSorted = view.sortCol === c.key;
    const arrow = !isSorted ? "⇅" : (view.sortDir === "asc" ? "↑" : "↓");
    const ariaSort = !isSorted ? "none" : (view.sortDir === "asc" ? "ascending" : "descending");
    const sortedCls = isSorted ? " sorted" : "";
    return `<th class="sortable${sortedCls}" data-col="${c.key}" aria-sort="${ariaSort}" title="Click to sort">` +
      `<span class="th-label">${escapeHtml(c.label)}</span>` +
      `<span class="sort-arrow">${arrow}</span>` +
      `</th>`;
  }).join("")}</tr></thead>`;

  const body = "<tbody>" + entries.map(({ r, idx }) => {
    const invalid = view.invalidCells.get(idx) || new Set();
    return "<tr>" + view.columns.map((c) => {
      const v = r[c.key];
      const shown = displayValue(v);
      const isNA = v === NA_MARKER;
      let cls = "";
      if (invalid.has(c.key)) cls = isNA ? "invalid-cell na-cell" : "invalid-cell";
      // Let columns inject extra classes (for P2P status highlighting)
      if (c.cellClass) cls = (cls + " " + c.cellClass(r)).trim();
      const attr = cls ? ` class="${cls}"` : "";
      // Columns may also inject custom HTML (e.g. status pills)
      const html = c.cellHtml ? c.cellHtml(r) : escapeHtml(shown);
      return `<td${attr} title="${escapeHtml(shown)}">${html}</td>`;
    }).join("") + "</tr>";
  }).join("") + "</tbody>";

  els.table.innerHTML = head + body;

  const parts = [`${entries.length} of ${view.rows.length} row${view.rows.length === 1 ? "" : "s"}`];
  if (view.showOnlyErrors) parts.push("errors only");
  if (q) parts.push(`search: "${view.filter}"`);
  if (view.sortCol) {
    const lbl = view.columns.find((c) => c.key === view.sortCol)?.label || view.sortCol;
    parts.push(`sorted by ${lbl} ${view.sortDir === "asc" ? "↑" : "↓"}`);
  }
  els.summary.textContent = parts.join(" · ") + ` · ${view.columns.length} columns`;
}

function refreshErrorToggle() {
  const n = view.invalidCells.size;
  if (n === 0) {
    els.errorFilterLabel.classList.add("hidden");
    view.showOnlyErrors = false;
    els.errorFilterCheckbox.checked = false;
  } else {
    els.errorFilterLabel.classList.remove("hidden");
    els.errorRowCount.textContent = n;
  }
}

export function openFullDataModal({ rows, columns, invalidCells = new Map(), loadingHint = "" }) {
  view = {
    rows,
    columns,
    invalidCells,
    sortCol: null,
    sortDir: null,
    showOnlyErrors: false,
    filter: "",
  };
  els.errorFilterCheckbox.checked = false;
  els.search.value = "";
  refreshErrorToggle();

  runWithLoading(
    "Loading full data…",
    loadingHint || `Preparing ${rows.length.toLocaleString()} rows for display.`,
    () => { render(); }
  ).then(() => {
    els.modal.classList.remove("hidden");
    setTimeout(() => els.search.focus(), 100);
  });
}

function closeModal() {
  els.modal.classList.add("hidden");
}

// One-time wiring (module loaded once at startup)
els.closeBtn.addEventListener("click", closeModal);
els.modal.addEventListener("click", (e) => { if (e.target === els.modal) closeModal(); });
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !els.modal.classList.contains("hidden")) closeModal();
});
els.table.addEventListener("click", (e) => {
  const th = e.target.closest("th.sortable");
  if (!th) return;
  const col = th.dataset.col;
  if (!col) return;
  if (view.sortCol !== col) {
    view.sortCol = col;
    view.sortDir = "asc";
  } else if (view.sortDir === "asc") {
    view.sortDir = "desc";
  } else {
    view.sortCol = null;
    view.sortDir = null;
  }
  render();
});
els.errorFilterCheckbox.addEventListener("change", (e) => {
  view.showOnlyErrors = e.target.checked;
  render();
});
let filterTimer = null;
els.search.addEventListener("input", (e) => {
  const q = e.target.value;
  clearTimeout(filterTimer);
  filterTimer = setTimeout(() => { view.filter = q; render(); }, 120);
});
