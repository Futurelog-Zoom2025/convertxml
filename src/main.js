// Thin orchestrator: wires the tab switcher and initializes each tab module.
// All real logic lives in tab1145.js and tabP2P.js.

import { initR1145Tab } from "./tab1145.js";
import { initP2PTab } from "./tabP2P.js";

// ---------- Tab switching ----------
const tabs = document.querySelectorAll(".tabs .tab");
const panels = {
  r1145: document.getElementById("panel1145"),
  p2p:   document.getElementById("panelP2P"),
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

// ---------- Initialize both tabs up-front ----------
// Both tabs are wired immediately so state is preserved when the user flips
// between tabs. The inactive panel is just hidden via CSS, not unmounted.
initR1145Tab();
initP2PTab();
