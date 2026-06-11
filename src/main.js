// Thin orchestrator: wires the tab switcher and initializes each tab module.
// All real logic lives in tab1145.js, tabP2P.js, and tabXmlToR1145.js.

import { initR1145Tab } from "./tab1145.js";
import { initP2PTab } from "./tabP2P.js";
import { initXmlToR1145Tab } from "./tabXmlToR1145.js";

// ---------- Tab switching ----------
const tabs = document.querySelectorAll(".tabs .tab");
const panels = {
  r1145: document.getElementById("panel1145"),
  p2p:   document.getElementById("panelP2P"),
  xml:   document.getElementById("panelXml"),
};

tabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    const target = tab.dataset.tab;
    tabs.forEach((t) => {
      const isActive = t === tab;
      t.classList.toggle("active", isActive);
      t.setAttribute("aria-selected", isActive ? "true" : "false");
    });
    Object.entries(panels).forEach(([key, panel]) => {
      panel.classList.toggle("active", key === target);
    });
  });
});

// ---------- Initialize all tabs up-front ----------
// All tabs are wired immediately so state is preserved when the user flips
// between tabs. Inactive panels are just hidden via CSS, not unmounted.
initR1145Tab();
initP2PTab();
initXmlToR1145Tab();

// ---------- First-visit news popup ----------
// Bump the version string when announcing a new feature so previously-dismissed
// users see the new announcement once.
const NEWS_KEY = "fl_news_seen_v1_xml_to_1145";
const newsModal = document.getElementById("newsModal");
if (newsModal && !localStorage.getItem(NEWS_KEY)) {
  newsModal.classList.remove("hidden");

  const dismiss = () => {
    newsModal.classList.add("hidden");
    try { localStorage.setItem(NEWS_KEY, "1"); } catch {}
  };

  document.getElementById("newsModalCloseBtn")?.addEventListener("click", dismiss);
  document.getElementById("newsModalTryBtn")?.addEventListener("click", () => {
    document.getElementById("tabBtnXml")?.click();
    dismiss();
  });
  newsModal.addEventListener("click", (e) => {
    if (e.target === newsModal) dismiss();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !newsModal.classList.contains("hidden")) dismiss();
  });
}
