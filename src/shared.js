// Small helpers reused by both tabs. Keeps main.js and the per-tab modules slim.

export const $ = (sel) => document.querySelector(sel);

// Exact list and order requested by the business.
// Duplicates in the source list (230 appeared twice) are deduped here.
// New IDs are appended at the end to avoid shifting existing positions in
// case anyone has bookmarked the order.
export const COMPANY_IDS = [
  "169", "215", "233", "247", "278", "257", "262", "230", "315",
  "101", "265", "225", "296", "285", "192",
];

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

export function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

export function todayDDMMYYYY() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}${mm}${d.getFullYear()}`;
}

export function delay(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// Given a validation error list, return a short single-line hint suitable
// for inline display next to the Generate button. Used by the param-level
// revalidate path in both tabs so the user sees WHY Generate is disabled
// without scrolling up to the validation panel.
//
// Strategy: prioritize parameter-shape errors (validity date, supplier no.,
// company id, language) since those are exactly what the user is trying to
// fix. If only row-level errors remain, fall back to a count.
//
// Returns null when there are no errors (caller should switch to "ready").
export function summarizeErrorForHint(errors) {
  if (!errors || errors.length === 0) return null;

  // Patterns for param-level errors, ordered by user priority. Each entry has
  // a `match` regex and a short `hint` (or a function that builds one from
  // the matched error string).
  const PARAM_PATTERNS = [
    { match: /Buddhist Era|พ\.ศ/i,           hint: (e) => firstLine(e) },
    { match: /Validity Date must be/i,       hint: () => "Validity Date must be 8 digits (DDMMYYYY)." },
    { match: /Supplier Number must be/i,     hint: () => "Supplier Number must be 6 digits." },
    { match: /Language must be/i,            hint: (e) => firstLine(e) },
    { match: /WebShop Company ID|companyId/i, hint: () => "Company ID must be 3 digits." },
  ];

  for (const p of PARAM_PATTERNS) {
    const hit = errors.find((e) => p.match.test(e));
    if (hit) return p.hint(hit);
  }

  // No param-level error matched — fall back to a generic count so the user
  // knows the issue is a row-level data problem (not Step 3).
  const n = errors.length;
  return `Fix the ${n} validation issue${n === 1 ? "" : "s"} above before generating.`;
}

function firstLine(s) {
  const i = s.indexOf("\n");
  return i === -1 ? s : s.slice(0, i);
}


// ---------- Loading overlay (single shared overlay in index.html) ----------
const overlay = {
  el: () => document.getElementById("loadingOverlay"),
  msg: () => document.getElementById("loadingMsg"),
  sub: () => document.getElementById("loadingSub"),
};

export function showLoading(msg, sub) {
  if (msg) overlay.msg().textContent = msg;
  if (sub !== undefined) overlay.sub().textContent = sub;
  overlay.el().classList.remove("hidden");
}
export function hideLoading() {
  overlay.el().classList.add("hidden");
}

// Wraps a synchronous heavy task and guarantees the overlay has painted before
// the task starts blocking. Two requestAnimationFrame hops is the reliable way.
export function runWithLoading(msg, sub, fn) {
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

// ---------- Download helper ----------
export function downloadBlob(content, filename, mimeType) {
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

// ---------- Company multi-select builder ----------
//
// Each tab has its own multi-select but they all use the same COMPANY_IDS list
// and the same UX. Build one here with namespaced DOM ids.
//
// Returns:
//   {
//     getSelected(): string[],
//     setLabel(): void,            // re-renders the button label from state
//     reset(): void                // unchecks everything
//   }
export function buildCompanyMultiselect({ rootId, btnId, labelId, menuId, optionsId, selectAllId, clearAllId }) {
  const root    = document.getElementById(rootId);
  const btn     = document.getElementById(btnId);
  const label   = document.getElementById(labelId);
  const menu    = document.getElementById(menuId);
  const options = document.getElementById(optionsId);
  const selAll  = document.getElementById(selectAllId);
  const clear   = document.getElementById(clearAllId);

  const selected = new Set();

  // Listeners registered via the returned `onChange()` API. Fired any time
  // the selection set changes (per-checkbox click, Select All, Clear All,
  // and reset()).
  const changeListeners = [];
  function fireChange() { for (const fn of changeListeners) fn(); }

  options.innerHTML = COMPANY_IDS.map((id) => `
    <label class="multiselect-option">
      <input type="checkbox" value="${id}" />
      <span class="company-code">${id}</span>
    </label>
  `).join("");

  function updateLabel() {
    const xs = Array.from(selected);
    if (xs.length === 0) {
      label.textContent = "Select companies…";
      label.classList.remove("has-selection");
      label.classList.add("muted");
    } else if (xs.length <= 4) {
      label.textContent = xs.join(", ");
      label.classList.add("has-selection");
      label.classList.remove("muted");
    } else {
      label.textContent = `${xs.length} companies selected (${xs.slice(0, 3).join(", ")}…)`;
      label.classList.add("has-selection");
      label.classList.remove("muted");
    }
  }

  options.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.addEventListener("change", () => {
      if (cb.checked) selected.add(cb.value);
      else selected.delete(cb.value);
      updateLabel();
      fireChange();
    });
  });

  function toggleMenu(open) {
    const isOpen = !menu.classList.contains("hidden");
    const shouldOpen = open === undefined ? !isOpen : open;
    if (shouldOpen) { menu.classList.remove("hidden"); root.classList.add("open"); }
    else            { menu.classList.add("hidden");    root.classList.remove("open"); }
  }

  btn.addEventListener("click", (e) => { e.stopPropagation(); toggleMenu(); });
  menu.addEventListener("click", (e) => e.stopPropagation());
  document.addEventListener("click", () => toggleMenu(false));

  selAll.addEventListener("click", () => {
    options.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
      cb.checked = true;
      selected.add(cb.value);
    });
    updateLabel();
    fireChange();
  });

  clear.addEventListener("click", () => {
    options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
    selected.clear();
    updateLabel();
    fireChange();
  });

  updateLabel();

  return {
    getSelected: () => Array.from(selected),
    reset: () => {
      options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
      selected.clear();
      updateLabel();
      fireChange();
    },
    onChange: (fn) => { if (typeof fn === "function") changeListeners.push(fn); },
  };
}

// ---------- Numeric-only input filter ----------
export function restrictToDigits(inputEl) {
  inputEl.addEventListener("input", (e) => {
    e.target.value = e.target.value.replace(/\D/g, "");
  });
}
