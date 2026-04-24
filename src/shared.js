// Small helpers reused by both tabs. Keeps main.js and the per-tab modules slim.

export const $ = (sel) => document.querySelector(sel);

// Exact list and order requested by the business.
// Duplicates in the source list (230 appeared twice) are deduped here.
export const COMPANY_IDS = [
  "169", "215", "233", "247", "278", "257", "262", "230", "315",
  "101", "265", "225", "296", "285",
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
  });

  clear.addEventListener("click", () => {
    options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
    selected.clear();
    updateLabel();
  });

  updateLabel();

  return {
    getSelected: () => Array.from(selected),
    reset: () => {
      options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
      selected.clear();
      updateLabel();
    },
  };
}

// ---------- Numeric-only input filter ----------
export function restrictToDigits(inputEl) {
  inputEl.addEventListener("input", (e) => {
    e.target.value = e.target.value.replace(/\D/g, "");
  });
}
